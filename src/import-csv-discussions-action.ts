import * as SDK from "azure-devops-extension-sdk";
import { CommonServiceIds, ILocationService, IProjectPageService } from "azure-devops-extension-api";

import * as csv from 'csvtojson';


interface IContributedMenuSource {
    execute?(actionContext: any) : any;
}

class ImportCSVDiscussionsAction implements IContributedMenuSource {
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
                            // for each array item lets create a new comment
                            records.forEach(async (record : any) => {

                                let header : Array<string> = [];
                                let cols : Array<string> = [];

                                // build ou html table header rows and body rows
                                Object.keys(record).forEach(key => {
                                    if(key !== "Title" &&
                                        key !== "WorkItemId")
                                    {
                                        header.push(`<th>${key}</th>`);
                                        cols.push(`<td>${record[key]}</td>`);
                                    }
                                });

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

                                // if multiple ids are specified for this row lets split them
                                let ids : string[] = record.WorkItemId.split(";");

                                // finally create the comment for each workitem
                                ids.forEach(async (id : string)=>{
                                    if(id && 
                                        id.trim() !== '')
                                    {
                                        await fetch(`${hostBaseUrl}${project.name}/_apis/wit/workItems/${id.trim()}/comments?api-version=6.0-preview.3`, {
                                            method: 'POST',
                                            headers: {
                                                'Authorization': `Bearer ${accessToken}`,
                                                'Content-Type': 'application/json',
                                            },
                                            body: JSON.stringify( discussion_comment )
                                        });
                                    }
                                });
                            });
                        }
                    }
                    else
                    {
                        alert("Error : CSV File is Empty.")
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
}


SDK.init();
SDK.ready().then(() =>{
    SDK.register(SDK.getContributionId(), () => {
        return new ImportCSVDiscussionsAction();
    });
});