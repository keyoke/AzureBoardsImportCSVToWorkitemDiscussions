import * as SDK from "azure-devops-extension-sdk";

SDK.init();

SDK.register("file-upload", () => {
    let callbacks : any = [];

    let input : HTMLInputElement = document.getElementById('file-input') as HTMLInputElement;

    function isValid() : boolean{
        // HAs the user selected a file?
        return input.files.length == 1;
    }

    input.addEventListener("change", () =>{
        // Execute registered callbacks
        for(var i = 0; i < callbacks.length; i++) {
            callbacks[i](isValid());
        }
    }, false);

    return {
        getFileContents : async () => {
            return new Promise((res,rej)=>{
                const fileList : FileList = input.files;

                console.log(fileList.length + " files selected.");

                // If we have a file lets read the contents
                if(fileList.length == 1)
                {
                    let file = fileList[0];

                    console.log("Selected File Name : " + file.name);

                    // new FileReader object
                    let reader = new FileReader();

                    // event fired when file reading finished
                    reader.addEventListener('loadend', function(e) {
                        // contents of the file
                        res(e.target.result.toString());
                    });

                    // event fired when file reading failed
                    reader.addEventListener('error', function() {
                        rej('Error : Failed to read file.');
                    });

                    // read file as text file
                    reader.readAsText(file);
                }
                else
                {
                    rej('Error : No File Selected.');
                }
            });
        },
        attachFileChanged: (cb : any) => {
            callbacks.push(cb);
       },
       isFileValid: () => {
            return isValid();
        }
    };
});