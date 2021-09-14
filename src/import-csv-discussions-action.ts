import * as SDK from "azure-devops-extension-sdk";
import { CommonServiceIds, ILocationService, IProjectPageService } from "azure-devops-extension-api";

import * as csv from 'csvtojson';
import Logger, { LogLevel } from "./logger";


interface IContributedMenuSource {
    execute?(actionContext: any) : any;
}

class ImportCSVDiscussionsAction implements IContributedMenuSource {
    _logger : Logger = new Logger(LogLevel.All);

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

                        // do we have an array?
                        if(records)
                        {
                            try {
                                let new_records = records.reduce((r, a) => {
                                    this._logger.debug("a", a);

                                    let ids : string[] = a.WorkItemId.split(";");
                                    
                                    this._logger.debug("ids", ids);

                                    ids.forEach(async (id : string)=>{
                                        this._logger.debug("id", ids);
                                        r[id] = r[id] || [];
                                        r[id].push(a);
                                    });

                                    return r;
                                }, Object.create(null));

                                console.dir(new_records);
                            } catch (error) {
                                this._logger.error(error);
                            }

                            // for each array item lets create a new comment
                            records.forEach(async (record : any) => {

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

                                // if multiple ids are specified for this row lets split them
                                let ids : string[] = record.WorkItemId.split(";");

                                this._logger.debug("ids", ids);

                                // finally create the comment for each workitem
                                ids.forEach(async (id : string)=>{
                                    let clean_id : string = id.replace(/\D/g,'');

                                    if(this.isNumber(clean_id))
                                    {
                                        this._logger.debug(`Adding comment for id '${clean_id}'`);

                                        try {
                                            await fetch(`${hostBaseUrl}${project.name}/_apis/wit/workItems/${clean_id}/comments?api-version=6.0-preview.3`, {
                                                method: 'POST',
                                                headers: {
                                                    'Authorization': `Bearer ${accessToken}`,
                                                    'Content-Type': 'application/json',
                                                },
                                                body: JSON.stringify( discussion_comment )
                                            });
                                        } catch (error){
                                            this._logger.error(`Failed to add comment id '${clean_id}'`, error);
                                        }
                                    }
                                    else
                                    {
                                        this._logger.debug(`id is not a number`, id);
                                    }
                                });
                            });
                        }
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