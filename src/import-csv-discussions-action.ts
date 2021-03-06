import * as SDK from 'azure-devops-extension-sdk';
import {
  CommonServiceIds,
  ILocationService,
  IProjectPageService,
} from 'azure-devops-extension-api';
import * as csv from 'csvtojson';
import { Parser } from 'json2csv';
import Logger, { LogLevel } from './logger';
import * as originalFetch from 'isomorphic-fetch';
import * as fetchBuilder from 'fetch-retry';

interface IContributedMenuSource {
  execute?(actionContext: any): any;
}

class ImportCSVDiscussionsAction implements IContributedMenuSource {
  _logger: Logger = new Logger(LogLevel.Info);

  _statusCodes = [409, 503, 504];
  _options = {
    retries: 5,
    retryDelay: (attempt: any, error: any, response: any) => {
      return Math.pow(2, attempt) * 1000;
    },
    retryOn: (attempt: any, error: any, response: any) => {
      // retry on any network error, or specific status codes
      if (error !== null || this._statusCodes.includes(response.status)) {
        this._logger.info(`retrying, attempt number ${attempt + 1}`);
        return true;
      }
    },
  };

  _fetch: any = fetchBuilder(originalFetch, this._options);
  _json2csvParser = new Parser();
  _failures: any[] = [];

  public showFileUpload() {
    SDK.getService('ms.vss-web.dialog-service').then(
      async (dialogService: any) => {
        // pointer to our form
        let fileUploadForm: any;

        // Get our extension details
        const extensionCtx = SDK.getExtensionContext();
        // Build absolute contribution ID for dialogContent
        const contributionId =
          extensionCtx.publisherId +
          '.' +
          extensionCtx.extensionId +
          '.file-upload';

        // Show dialog
        const dialogOptions = {
          title: 'Import CSV to discussions',
          width: 800,
          height: 600,
          okText: 'Import',
          cancelText: 'Cancel',
          getDialogResult: async () => {
            // Get the file contents
            return fileUploadForm ? await fileUploadForm.getFileContents() : '';
          },
          okCallback: async (result: string) => {
            // clear any existing failures
            this._failures = [];

            // do we have some data
            if (result) {
              this._logger.info('Started Import.');

              // Get the HOST URI
              const service: ILocationService = await SDK.getService(
                CommonServiceIds.LocationService
              );
              const hostBaseUrl = await service.getResourceAreaLocation(
                '5264459e-e5e0-4bd8-b118-0985e68a4ec5' // WIT
              );

              const projectService = await SDK.getService<IProjectPageService>(
                CommonServiceIds.ProjectPageService
              );
              const project = await projectService.getProject();

              // Get Access Token as we will execute simple rest call
              const accessToken = await SDK.getAccessToken();

              // convert our csv data to json array
              const records: any[] = await csv().fromString(result);

              // do we have an array?
              if (records) {
                this._logger.info(`'${records.length}' records to import.`);

                const promises: Promise<boolean>[] = [];

                const batches = records.reduce((r, a) => {
                  // Let's make sure we already have a return array intitialized
                  if (!r || !Array.isArray(r)) {
                    r = [];
                  }

                  // Handle Mutiple id's which a ';' delimited
                  const ids: string[] = a.WorkItemId.split(';');

                  this._logger.debug('ids', ids);

                  // Loop through the supplied ids
                  ids.forEach((id: string) => {
                    let added = false;

                    // ensure we have a number and we need to add seperate records for each id
                    const n = Object.assign({}, a);
                    n.WorkItemId = id.replace(/\D/g, '');

                    this._logger.debug('n', n);

                    // If we already have items in r then lets check if our new item can be placed in one of the existing arrays
                    r.every((i: any) => {
                      // Does this item already contain this WorkItemId?
                      if (
                        i.filter((f: any) => {
                          return f.WorkItemId === n.WorkItemId;
                        }).length === 0
                      ) {
                        // this id doesnt exist so place it here
                        i.push(n);
                        added = true;
                        // End our search as we have inserted this item into our return array
                        return false;
                      } else {
                        // Look in the next array
                        return true;
                      }
                    });

                    // Our item was not added to any existing array therefore lets add a new array and add the current item
                    if (!added) {
                      r.push([n]);
                    }
                  });

                  return r;
                }, Object.create(null));

                this._logger.debug('batches', batches);

                // loop through each batch and apply the update to ADO
                batches.forEach(async (batch: any) => {
                  // let batch_payload : any = [];

                  // Loop through each record in this batch and generate the JSON
                  batch.forEach(async (record: any) => {
                    this._logger.debug('record', record);

                    /* batch_payload.push({
                                                        "method": "PATCH",
                                                        "uri": `/_apis/wit/workitems/${record.WorkItemId}?api-version=4.1`,
                                                        "headers": {
                                                            "Content-Type": "application/json-patch+json"
                                                        },
                                                        "body": [{
                                                            "op": "add",
                                                            "path": "/fields/System.History",
                                                            "value": `${discussion_comment}`
                                                            }
                                                        ]
                                                    }); */

                    promises.push(
                      this.createComment(
                        `${hostBaseUrl}${project.name}/_apis/wit/workItems/${record.WorkItemId}/comments?api-version=6.0-preview.3`,
                        accessToken,
                        record
                      )
                    );
                  });

                  /* // Finally apply the batched updates
                                    if(batch_payload.length > 0)
                                    {
                                        try {
                                            await fetch(`${hostBaseUrl}_apis/wit/$batch?api-version=4.1`, {
                                                method: 'PATCH',
                                                headers: {
                                                    'Authorization': `Bearer ${accessToken}`,
                                                    'Content-Type': 'application/json-patch+json',
                                                },
                                                body: JSON.stringify( batch_payload )
                                            });
                                        } catch (error){
                                            this._logger.error(`Failed to add comments`, error);
                                        }
                                    } */
                });

                // wait for promises
                await Promise.all(promises.map((p) => p.catch((e) => e))).then(
                  () => {
                    this._logger.info(
                      `'${this._failures.length}' records failed to import.`
                    );

                    if (this._failures.length > 0) {
                      try {
                        // convert our array of failed imports back into csv format
                        const csv = this._json2csvParser.parse(this._failures);

                        // create a buffer for our csv string
                        const buff = Buffer.from(csv, 'utf-8');

                        // decode buffer as Base64
                        const base64 = buff.toString('base64');

                        // Attempt to send our file containing failures back to the user
                        const a: HTMLAnchorElement =
                          document.createElement('a');
                        document.body.appendChild(a);
                        a.download = 'import-failed.csv';
                        a.href = `data:text/plain;base64,${base64}`;
                        a.click();
                      } catch (error) {
                        this._logger.error(error);
                      }
                    }

                    this._logger.info('Ended Import.');
                  }
                );
              } else {
                alert('Unable to parse CSV file contents.');
                this._logger.error('Unable to parse CSV file contents.');
              }
            } else {
              alert('CSV File is Empty.');
              this._logger.error('CSV File is Empty.');
            }
          },
        };

        dialogService
          .openDialog(contributionId, dialogOptions)
          .then((dialog: any) => {
            // Get fileUploadForm instance which is registered in file-upload-dialog.html
            dialog
              .getContributionInstance('file-upload')
              .then((fileUploadFormInstance: any) => {
                // Keep a reference of fileUpload form instance (to be used previously in dialog options)
                fileUploadForm = fileUploadFormInstance;

                // Subscribe to form input changes and update the Ok enabled state
                fileUploadForm.attachFileChanged((isValid: boolean) => {
                  dialog.updateOkButton(isValid);
                });

                // Set the initial ok enabled state
                const isValid: boolean = fileUploadForm.isFileValid();
                dialog.updateOkButton(isValid);
              });
          });
      }
    );
  }

  async createComment(
    url: string,
    accessToken: string,
    record: any
  ): Promise<boolean> {
    return new Promise((resolve, reject) => {
      const header: Array<string> = [];
      const cols: Array<string> = [];

      // build our html table header rows and body rows
      Object.keys(record).forEach((key) => {
        if (key !== 'Title' && key !== 'WorkItemId') {
          header.push(`<th>${key}</th>`);
          cols.push(`<td>${record[key]}</td>`);
        }
      });

      this._logger.debug('header', header);
      this._logger.debug('cols', cols);

      // put the table together with its title
      const discussion_comment = {
        text: `<table style="width:100%">
                            <thead>
                                <tr>
                                    <th  colspan="${header.length}">${
          record.Title
        }</th>
                                </tr>
                                <tr>
                                    ${header.join('\n')}
                                </tr>
                            </thead>
                            <tbody>
                                <tr>
                                    ${cols.join('\n')}
                                </tr>
                            </tbody>
                        </table>`,
      };

      this._logger.debug('discussion_comment', discussion_comment);

      this._logger.info(`Adding comment for id '${record.WorkItemId}'`);

      this._fetch(url, {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(discussion_comment),
      })
        .then(async (response: Response) => {
          if (response.status >= 200 && response.status < 300) {
            const json: string = await response.json();
            this._logger.info(
              `Successfully added comment for id '${record.WorkItemId}'`
            );
            // log any JSON to debug
            this._logger.debug('json', json);
            resolve(true);
          } else {
            this._logger.info(
              `Failed adding comment for id '${record.WorkItemId}'`
            );
            this._failures.push(Object.assign({}, record));
            resolve(false);
          }
        })
        .catch((error: Error) => {
          // Save this failure for later
          this._logger.error('Unhandled Error.', error);
          this._failures.push(Object.assign({}, record));
          reject(error);
        });
    });
  }

  public execute(actionContext: any) {
    this.showFileUpload();
  }

  private isNumber(n: any) {
    if (n) return !isNaN(parseFloat(n)) && !isNaN(n - 0);
    else return false;
  }
}

export async function main(): Promise<void> {
  await SDK.init();

  // wait until we are ready
  await SDK.ready();

  SDK.register(SDK.getContributionId(), () => {
    return new ImportCSVDiscussionsAction();
  });
}

// execute our entrypoint
main().catch((error) => {
  console.error(error);
});
