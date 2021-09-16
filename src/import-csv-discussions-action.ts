import * as SDK from "azure-devops-extension-sdk";
import { CommonServiceIds, ILocationService, IProjectPageService } from "azure-devops-extension-api";
import * as csv from 'csvtojson';
import { Parser } from 'json2csv';
import Logger, { LogLevel } from "./logger";
import * as originalFetch from 'isomorphic-fetch';
import * as fetchBuilder from 'fetch-retry';

interface IContributedMenuSource {
    execute?(actionContext: any) : any;
}

class ImportCSVDiscussionsAction implements IContributedMenuSource {
    _logger : Logger = new Logger(LogLevel.Info);

    _statusCodes = [409, 503, 504];
    _options = {
        retries: 5,
        retryDelay: (attempt : any, error  : any, response  : any) => {
            return Math.pow(2, attempt) * 1000;
        },
        retryOn: (attempt  : any, error  : any, response  : any)  => {
            // retry on any network error, or specific status codes
            if (error !== null || this._statusCodes.includes(response.status)) {
                this._logger.info(`retrying, attempt number ${attempt + 1}`);
                return true;
            }
        }
    };

    _fetch : any = fetchBuilder(originalFetch, this._options);
    _json2csvParser = new Parser();


    constructor()
    {

    }

    public showFileUpload () {
        SDK
        .getService("ms.vss-web.dialog-service")
        .then(async(dialogService : any) => {
            // pointer to our form
            let fileUploadForm : any;

            // Get our extension details
            let extensionCtx = SDK.getExtensionContext();
            // Build absolute contribution ID for dialogContent
            let contributionId = extensionCtx.publisherId + "." + extensionCtx.extensionId + ".file-upload";
    
            // Show dialog
            let dialogOptions = {
                title: "Import CSV to discussions",
                width: 800,
                height: 600,
                okText: "Import",
                cancelText: "Cancel" ,
                getDialogResult: async () => {
                    // Get the file contents
                    return fileUploadForm ? await fileUploadForm.getFileContents() : "";
                },
                okCallback: async (result : string) => {
                    // do we have some data
                    if(result)
                    {
                        this._logger.info(`Started Import.`);

                        // Get the HOST URI
                        const service : ILocationService = await SDK.getService(CommonServiceIds.LocationService);
                        const hostBaseUrl = await service.getResourceAreaLocation(
                            '5264459e-e5e0-4bd8-b118-0985e68a4ec5' // WIT
                        );

                        const projectService = await SDK.getService<IProjectPageService>(CommonServiceIds.ProjectPageService);
                        const project = await projectService.getProject();

                        // Get Access Token as we will execute simple rest call
                        const accessToken = await SDK.getAccessToken();

                        // convert our csv data to json array
                        const records = await csv().fromString(result);
                        let failures : any = [];

                        // do we have an array?
                        if(records)
                        {
                            this._logger.info(`'${records.length}' records to import.`);

                            let batches = records.reduce((r, a) => {
                                
                                // Let's make sure we already have a return array intitialized
                                if(!r ||
                                    !Array.isArray(r))
                                {
                                    r = [];
                                }
                                
                                // Handle Mutiple id's which a ';' delimited
                                let ids : string[] = a.WorkItemId.split(";");
    
                                this._logger.debug("ids", ids);

                                // Loop through the supplied ids
                                ids.forEach((id : string)=>{
                                    let added = false;

                                    // ensure we have a number and we need to add seperate records for each id
                                    let n = Object.assign({}, a);
                                    n.WorkItemId = id.replace(/\D/g,'');

                                    this._logger.debug("n", n);
                                    
                                    // If we already have items in r then lets check if our new item can be placed in one of the existing arrays
                                    r.every((i : any) => {
                                        // Does this item already contain this WorkItemId?
                                        if(i.filter((f : any) => {
                                                return f.WorkItemId === n.WorkItemId;
                                            }).length === 0)
                                        {
                                            // this id doesnt exist so place it here
                                            i.push(n);
                                            added = true;
                                            // End our search as we have inserted this item into our return array
                                            return false;
                                        }
                                        else
                                        {
                                            // Look in the next array
                                            return true;
                                        }
                                    });
                                    
                                    // Our item was not added to any existing array therefore lets add a new array and add the current item
                                    if(!added)
                                    {
                                        r.push([n]);
                                    }
                                });
                                
                                return r;
                            }, Object.create(null));

                            this._logger.debug("batches", batches);

                            // loop through each batch and apply the update to ADO
                            batches.forEach(async (batch : any) => {
                                // let batch_payload : any = [];

                                // Loop through each record in this batch and generate the JSON
                                batch.forEach(async (record : any) => {
                                    this._logger.debug("record", record);

                                    let header : Array<string> = [];
                                    let cols : Array<string> = [];
    
                                    // build our html table header rows and body rows
                                    Object.keys(record).forEach(key => {
                                        if(key !== "Title" &&
                                            key !== "WorkItemId")
                                        {
                                            header.push(`<th>${key}</th>`);
                                            cols.push(`<td>${record[key]}</td>`);
                                        }
                                    });
    
                                    this._logger.debug("header", header);
                                    this._logger.debug("cols", cols);
    
                                    // put the table together with its title
                                    let discussion_comment = {
                                        text : `<table style="width:100%">
                                                    <thead>
                                                        <tr>
                                                            <th  colspan="${header.length}">${record.Title}</th>
                                                        </tr>
                                                        <tr>
                                                            ${header.join("\n")}
                                                        </tr>
                                                    </thead>
                                                    <tbody>
                                                        <tr>
                                                            ${cols.join("\n")}
                                                        </tr>
                                                    </tbody>
                                                </table>`
                                    }
    
                                    this._logger.debug("discussion_comment", discussion_comment);

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
                                    try {
                                        this._logger.info(`Adding comment for id '${record.WorkItemId}'`);

                                        // Make sure we retry
                                        let response : Response = await this._fetch(`${hostBaseUrl}${project.name}/_apis/wit/workItems/${record.WorkItemId}/comments?api-version=6.0-preview.3`, {
                                            method: 'POST',
                                            headers: {
                                                'Authorization': `Bearer ${accessToken}`,
                                                'Content-Type': 'application/json',
                                            },
                                            body: JSON.stringify( discussion_comment )
                                        });

                                        if(response.status >= 200 && response.status < 300)
                                        {
                                            this._logger.info(`Successfully added comment for id '${record.WorkItemId}'`);
                                        }
                                        else
                                        {
                                            this._logger.info(`Failed to add comment for id '${record.WorkItemId}', Http Response Status '${response.status}'`);
                                            this._logger.debug("response", response);
                                            failures.pop(record);
                                        }

                                    }
                                    catch(error)
                                    {
                                        this._logger.info(`Failed to add comment for id '${record.WorkItemId}'`);
                                        this._logger.error(error);
                                        failures.pop(record);
                                    }
                                });

                                if(failures.length > 0)
                                {
                                    try {
                                        this._logger.info(`'${failures.length}' records failed to import.`);

                                        // convert our array of failed imports back into csv format
                                        const csv = this._json2csvParser.parse(failures);
    
                                        // create a buffer for our csv string
                                        const buff = Buffer.from(csv, 'utf-8');
    
                                        // decode buffer as Base64
                                        const base64 = buff.toString('base64');
    
                                        // Attempt to send our file containing failures back to the user
                                        let a : HTMLAnchorElement = document.createElement('a');
                                        document.body.appendChild(a);
                                        a.download = "import-failed.csv";
                                        a.href = `data:text/plain;base64,${base64}`;
                                        a.click();

                                    } catch (error) {
                                        this._logger.error(error);
                                    }
                                }

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
                        }
                        this._logger.info(`Ended Import.`);
                    }
                    else
                    {
                        alert("Error : CSV File is Empty.")
                        this._logger.error("Error : CSV File is Empty.");
                    }
                }
            };

            dialogService
            .openDialog(contributionId, dialogOptions)
            .then((dialog : any) => {
                // Get fileUploadForm instance which is registered in file-upload-dialog.html
                dialog
                .getContributionInstance("file-upload")
                .then((fileUploadFormInstance : any) => {
                    // Keep a reference of fileUpload form instance (to be used previously in dialog options)
                    fileUploadForm = fileUploadFormInstance;

                    // Subscribe to form input changes and update the Ok enabled state
                    fileUploadForm.attachFileChanged((isValid : boolean) => {
                        dialog.updateOkButton(isValid);
                    });
                    
                    // Set the initial ok enabled state
                    let isValid : boolean = fileUploadForm.isFileValid();
                    dialog.updateOkButton(isValid);
                });
            });
        });
    }

    public execute(actionContext: any)  {
        this.showFileUpload();
    }

    private isNumber(n : any) { 
        if(n)
            return !isNaN(parseFloat(n)) && !isNaN(n - 0) 
        else
            return false;
    }
}

export async function main(): Promise<void> {
    await SDK.init();

    // wait until we are ready
    await SDK.ready();

    SDK.register(SDK.getContributionId(), () => {
        return new ImportCSVDiscussionsAction();
    });
};

// execute our entrypoint
main().catch((error) => { console.error(error); });