export interface AppOptions {
    serverUrl: string;
    siteUrl: string;
    listId: string;
    projectListId: string;
    folderUrl: string;
    docPartNameListId: string;
    projectId: number;
    returnUrl?: string;
}

export interface RootState {
    version?: string;
}

export interface DocTypeSearchPattern {
    pattern: string;
    itemId: number;
}

export interface DocType {
    id: number;
    title: string;
    partNum?: string,
    sortNum?: number, 
}

export interface FetchedFileData {
    file: SP.File;
    listItem: SP.ListItem;
    iconUrl: string;
}

export interface ItemData {
    id: number;
    [field: string]: any;
}

export interface StateFile {
    id: number;
    fileName: string;
    docType: DocType;
    changedDocType: DocType;
    info?: string;
    changedInfo?: string;
    saved: boolean;
    changed: boolean;
    majorVersion: number;
    minorVersion: number;
    checkinComment?: string;
    changedCheckinComment?: string;
    serverRelativeUrl: string;
    overwriteVersion: boolean;
    item: ItemData;
    iconUrl: string;
    overwrite: boolean;
    lastModified: Date;
    fileSize: number;
}

export type StateFileType = StateFile;

export interface StateObject {
    mode: UploadMode;
    loading?: boolean;
    files: StateFile[];
    counter: number;
    docTypes: DocType[];
    docTypeSearchPatterns: DocTypeSearchPattern[];
    projectItem: ItemData;
    options: AppOptions;
}

export interface DialogText{
    instructions: string,
    buttonOk: string,
    subject: string,
}

export enum SaveEditButtonKey {
    Unknown = 0,
    Save = 1,
    Edit = 2,
    Remove = 3,
    Cancel = 4,
}

export enum UploadMode {
    Unknown = 0,
    New = 1,
    Edit = 2,
}