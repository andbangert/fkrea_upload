declare function LoadSodByKey(scriptName: string, func: () => void): void;
declare function registerDragUpload(dropTargetElement: Element,
    serverUrl: string,
    siteRelativeUrl: string,
    listName: string,
    rootFolder: string,
    overwriteAll: boolean,
    hideProgressBar: boolean,
    refreshFunction: Function,
    preUploadFunction: (a: FileElement[]) => void,
    postUploadFunction: (a: FileElement[]) => void,
    checkPermissionFunction: () => boolean)
    : void;

declare interface FileElement {
    droppedFile: File;
    fileName: string;//"printFormProjectDisp - Copy (2).pdf"
    fileSize: Number;
    isFolder: boolean;
    overwrite: boolean;
    relativeFolder: string;
    status: Number;
    statusText: string;
    subFolder: string;
    type: string; //"application/pdf"
}

declare function NotifyScriptLoadedAndExecuteWaitingJobs(scriptFileName: string): any;