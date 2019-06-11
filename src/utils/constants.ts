
export const SPScripts = {
    Reputation: {
        Script: 'reputation.js',
        ReputationType: 'Microsoft.Office.Server.ReputationModel.Reputation',
        LayoutPath: '/_layouts/15/reputation.js',
    },
    SP: {
        Script: 'sp.js',
        ClientContextType: 'SP.ClientContext',
        UI: {
            Status: 'SP.UI.Status',
            Notify: 'SP.UI.Notify',
            Dialog: {
                Script: 'sp.ui.dialog.js',
                ModalDialog: 'SP.UI.ModalDialog',
                ShowModalDialog: 'SP.UI.ModalDialog.showModalDialog',
            },
        },
    },
    DragDrop: {
        Script: 'dragdrop.js',
    },
    SP_UI_Dialog: {
        Script: 'sp.ui.dialog.js',
        ShowModalDialog: 'SP.UI.ModalDialog.showModalDialog',
    },
};

export const Fields = {
    TypeOfJobs: 'TypeOfJobs',
    BuildObj: 'BuildObj',
    Path: 'Path',
    ProjectCardID: 'ProjectCardID',
    PartName: 'PartName',
    Comment: 'Comment',
    BuildObjForFile: 'BuildObjForFile',
    Modified: 'Modified',
    FileSize: 'File_x0020_Size',
};
