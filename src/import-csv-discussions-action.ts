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

                        // Get Access Token as we will execute simple rest call
                        const accessToken = await SDK.getAccessToken();

                        // convert our csv data to json array
                        const records = await csv().fromString(result);

                        // do we have an array?
                        if(records)
                        {
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
                                ids.forEach(async (id : string)=>{
                                    let added = false;
                                    
                                    // ensure we have a number and we need to add seperate records for each id
                                    a.WorkItemId = id.replace(/\D/g,'');
                                    
                                    this._logger.debug("a", a);
                                    
                                    // If we already have items in r then lets check if our new item can be placed in one of the existing arrays
                                    r.every((i : any) => {
                                        // Does this item already contain this WorkItemId?
                                        if(i.filter((f : any) => {
                                                return f.WorkItemId === a.WorkItemId;
                                            }).length === 0)
                                        {
                                            // this id doesnt exist so place it here
                                            i.push(a);
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
                                        r.push([a]);
                                    }
                                });
                                
                                return r;
                            }, Object.create(null));

                            console.dir(batches);

                            // loop through each batch and apply the update to ADO
                            batches.forEach(async (batch : any) => {
                                let batch_payload : any = [];

                                // Loop through each record in this batch and generate the JSON
                                batch.forEach(async (record : any) => {
                                    this._logger.debug(`Adding comment for id '${record.WorkItemId}'`);
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

                                    batch_payload.push({
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
                                                });
                                });

                                // Finally apply the batched updates
                                if(batch_payload.length > 0)
                                {
                                    try {
                                        await fetch(`${hostBaseUrl}_apis/wit/$batch?api-version=4.1`, {
                                            method: 'PATCH',
                                            headers: {
                                                'Authorization': `Bearer ${accessToken}`,
                                                'Content-Type': 'application/json',
                                            },
                                            body: JSON.stringify( batch_payload )
                                        });
                                    } catch (error){
                                        this._logger.error(`Failed to add comments`, error);
                                    }
                                }
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