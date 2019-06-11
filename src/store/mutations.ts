import {
    SET_LOADING_STATE,
    UPLOAD_FILE,
    SET_PART_FILE_NAMES,
    SET_DOC_TYPE_PATTERN,
    CHANGE_FILE,
    SET_PROJECT_ITEM,
    SET_FILE_ITEM_DATA,
    SET_UPLOAD_MODE,
    DELETE_FILE,
    SET_FILE_UNSAVED,
    SAVE_FILE, 
    APPLY_CHANGES,
    SET_OPTIONS,
} from './mutation-types';
import { StateObject, StateFile, DocTypeSearchPattern, DocType, ItemData, UploadMode, AppOptions } from '../types';
import { Fields } from '../utils/constants';

export const mutations = {
    [SET_OPTIONS](state: StateObject, options: AppOptions) {
        state.options = options;
    },
    [SET_LOADING_STATE](state: StateObject, loading: boolean) {
        state.loading = loading;
    },
    [UPLOAD_FILE](state: StateObject, stateFile: StateFile) {
        const oldFile = state.files.find((f) => f.fileName === stateFile.fileName);
        // if (stateFile.file.overwrite) {
        if (stateFile.overwrite) {
            if (oldFile) {
                // just oerwrite file
                oldFile.changed = true;
                oldFile.saved = false;
                oldFile.lastModified = stateFile.lastModified;
                oldFile.majorVersion = stateFile.majorVersion;
                oldFile.minorVersion = stateFile.minorVersion;
                oldFile.fileSize = stateFile.fileSize;
                // oldFile.file = stateFile.file;
            } else {
                stateFile.changed = true;
                stateFile.saved = false;
                state.files.push(stateFile);
            }
        } else if (!oldFile) {
            stateFile.changed = true;
            state.files.push(stateFile);
        }
    },
    [SET_PART_FILE_NAMES](state: StateObject, items: DocType[]) {
        state.docTypes = items;
    },
    [SET_DOC_TYPE_PATTERN](state: StateObject, docTypePatterns: DocTypeSearchPattern[]) {
        state.docTypeSearchPatterns = docTypePatterns;
    },
    [CHANGE_FILE](state: StateObject, payload: { fileName: string, docType: DocType, info?: string }) {
        const file = state.files.find((f) => f.fileName === payload.fileName);
        if (!file) {
            throw new Error('File can not be null!');
        }

        if (file.docType.id !== payload.docType.id) {
            file.changedDocType = payload.docType;
            file.changed = true;
            file.saved = false;
        }

        if (file.info !== payload.info) {
            file.changedInfo = payload.info;
            file.changed = true;
            file.saved = false;
        }
    },
    [SAVE_FILE](state: StateObject, file: StateFile) {
        const sf = state.files.find((f) => f.fileName === file.fileName);
        if (!sf) {
            throw new Error('File can not be undefined or null!');
        }

        sf.majorVersion += 1;
        sf.minorVersion = 0;
        sf.changed = false;
        sf.saved = true;
    },
    [SET_PROJECT_ITEM](state: StateObject, project: ItemData) {
        state.projectItem = project;
    },
    [SET_FILE_ITEM_DATA](state: StateObject, payload: { fileName: string, item: ItemData }) {
        const file = state.files.find((f) => f.fileName === payload.fileName);
        if (!file) {
            throw new Error('File can not be null!');
        }
        file.item = payload.item;
    },
    [SET_UPLOAD_MODE](state: StateObject, mode: UploadMode) {
        if (mode !== UploadMode.Unknown) {
            state.mode = mode;
        } else {
            throw new Error('Unknown Upload mode');
        }
    },
    [DELETE_FILE](state: StateObject, payload: { file: StateFile }) {
        const sf = state.files.find((f) => f.fileName === payload.file.fileName);
        if(!sf) {
            throw new Error('File with name ' + payload.file.fileName + 
            ' could not be found in the collection.');
        }
        const index = state.files.indexOf(sf);
        state.files.splice(index, 1);
    },
    [SET_FILE_UNSAVED](state: StateObject, payload: { file: StateFile }) {
        const sf = state.files.find((f) => f.fileName === payload.file.fileName);
        if(!sf) {
            throw new Error('File with name ' + payload.file.fileName + 
            ' could not be found in the collection.');
        }
        sf.saved = false;
    },
    [APPLY_CHANGES](state: StateObject, payload: StateFile) {
        const file = state.files.find((f) => f.fileName === payload.fileName);
        if (!file) {
            throw new Error('File can not be null!');
        }

        console.log(payload.changedDocType)
        console.log(file.docType)
        if (file.docType.id !== payload.changedDocType.id) {
            file.docType = payload.changedDocType;
            if (file.docType.id <= 0) {
                file.item[Fields.PartName] = null;
            } else {
                const lv = file.item[Fields.PartName] as SP.FieldLookupValue;
                if (lv) {
                    lv.set_lookupId(payload.docType.id);
                } else {
                    const pn = new SP.FieldLookupValue();
                    pn.set_lookupId(payload.docType.id);
                    file.item[Fields.PartName] = pn;
                }
            }
        }

        if (file.info !== payload.changedInfo) {
            file.info = payload.changedInfo;
            file.checkinComment = payload.changedInfo;
            file.item[Fields.Comment] = file.info;
        }
    },
};
