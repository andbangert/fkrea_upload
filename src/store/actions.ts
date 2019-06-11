import actionTypes from './action-types';
import {
    UPLOAD_FILE,
    SET_DOC_TYPE_PATTERN,
    SET_PART_FILE_NAMES,
    CHANGE_FILE, SAVE_FILE,
    SET_PROJECT_ITEM,
    SET_FILE_ITEM_DATA,
    SET_UPLOAD_MODE,
    DELETE_FILE,
    APPLY_CHANGES,
    SET_OPTIONS,
} from './mutation-types';
import { ActionTree } from 'vuex';
import {
    StateObject,
    DocTypeSearchPattern,
    DocType,
    ItemData,
    StateFile,
    UploadMode,
    AppOptions,
} from '../types';
import { SPDataService } from '@/utils';
import { Fields } from '../utils/constants';

export const actions: ActionTree<StateObject, StateObject> = {
    async [actionTypes.FETCH_FILES]({ commit, dispatch, state }, payload: {
        siteUrl: string, listId: string, folderUrl: string,
    }) {
        const files = await SPDataService.Current()
            .getFiles(payload.siteUrl, `${payload.siteUrl}${payload.folderUrl}`);

        if (!state.docTypes) {
            throw Error('docTypes must be initialized first.');
        }

        if (files && files.get_count() > 0) {
            const num = files.get_count();
            const iterator = files.getEnumerator();
            const sfarr: StateFile[] = new Array<StateFile>(num);

            while (iterator.moveNext()) {
                const file: SP.File = iterator.get_current();
                const item: SP.ListItem = file.get_listItemAllFields();
                const itemId = item.get_id();
                const mjversion = file.get_majorVersion();
                const mnversion = file.get_minorVersion();
                const checkinComment = item.get_item(Fields.Comment); //file.get_checkInComment();
                const serverRelativeUrl = file.get_serverRelativeUrl();
                const fileName = file.get_name();
                const docType: SP.FieldLookupValue = item.get_item(Fields.PartName) as SP.FieldLookupValue;
                const lastModified: Date = item.get_item(Fields.Modified) as Date;
                const icon = await SPDataService.mapToIcon(payload.siteUrl, fileName, SP.Utilities.IconSize.size16);
                const fileSize = item.get_item(Fields.FileSize);

                const dtId = docType ? docType.get_lookupId() : -1;
                const dtv: DocType | undefined = state.docTypes.find(d => d.id === dtId);

                // then commit
                commit(UPLOAD_FILE, {
                    id: itemId,
                    majorVersion: mjversion,
                    minorVersion: mnversion,
                    checkinComment, // : checkinComment,
                    serverRelativeUrl, // : serverRelativeUrl,
                    fileName, // : fileName,
                    lastModified, // : lastModified,
                    overwrite: false,
                    docType: dtv ? dtv : { id: -1, title: '' },
                    iconUrl: icon,
                    saved: true,
                    fileSize, // : fileSize,
                    info: checkinComment,
                });
                // Set Item
                await dispatch(actionTypes.SET_FILE_ITEM_DATA,
                    {
                        fileName, // : fileName,
                        itemId, // : itemId,
                        docType, // : docType,
                    });
            }
        }
    },
    // Upload files
    [actionTypes.UPLOAD_FILES]({ commit, dispatch, state }, payload: {
        siteUrl: string, listId: string, files: FileElement[], folderUrl: string,
    }) {
        const files = payload.files;

        if (files && files.length !== 0) {
            files.sort((ele1, ele2) => {
                if (ele1.fileName > ele2.fileName) {
                    return 1;
                }
                if (ele1.fileName < ele2.fileName) {
                    return -1;
                }
                return 0;
            }).forEach(async (element) => {
                // Define default document type by name.
                let docTypeId: number = -1;
                if (state.docTypeSearchPatterns && state.docTypeSearchPatterns.length > 0) {
                    for (let i = 0; i < state.docTypeSearchPatterns.length; i++) {
                        const p = state.docTypeSearchPatterns[i].pattern;
                        if (element.fileName.startsWith(p)) {
                            docTypeId = state.docTypeSearchPatterns[i].itemId;
                            break;
                        }
                    }
                }
                let docType: DocType = { id: -1, title: '' };
                if (state.docTypes && state.docTypes.length) {
                    const doctc = state.docTypes.find((docTypeElement) => {
                        return docTypeElement.id === docTypeId;
                    });
                    docType = doctc ? doctc : { id: -1, title: '' };
                }
                const fileUrl = `${payload.siteUrl}${payload.folderUrl}/${element.fileName}`;
                // Element
                const prfile = SPDataService.Current().getFile(payload.siteUrl, fileUrl);
                // fetch uploaded file
                prfile.then((a) => {
                    const itemId = a.listItem.get_id();
                    const mjversion = a.file.get_majorVersion();
                    const mnversion = a.file.get_minorVersion();
                    const checkinComment = a.file.get_checkInComment();
                    const serverRelativeUrl = a.file.get_serverRelativeUrl();
                    // then commit
                    commit(UPLOAD_FILE, {
                        id: itemId,
                        majorVersion: mjversion,
                        minorVersion: mnversion,
                        checkinComment, // : checkinComment,
                        serverRelativeUrl, // : serverRelativeUrl,
                        fileName: element.fileName,
                        lastModified: new Date(element.droppedFile.lastModified),
                        overwrite: element.overwrite,
                        docType, // : docType,
                        changedDocType: docType,
                        changedInfo: checkinComment,
                        iconUrl: a.iconUrl,
                        fileSize: element.fileSize,
                    });

                    let dtv; // = undefined;
                    if (docType.id > 0) {
                        dtv = new SP.FieldLookupValue();
                        dtv.set_lookupId(docType.id);
                    }

                    dispatch(actionTypes.SET_FILE_ITEM_DATA,
                        {
                            fileName: element.fileName,
                            itemId, // : itemId,
                            docType: dtv,
                        });
                }).catch((e) => {
                    throw e;
                });
            });
        }
    },
    [actionTypes.SET_FILE_ITEM_DATA]({ commit, state }, payload: {
        fileName: string, itemId: number, docType: SP.FieldLookupValue, comment: string,
    }) {
        const projcard = new SP.FieldLookupValue();
        projcard.set_lookupId(state.projectItem.id);

        const item: ItemData = {
            id: payload.itemId,
            [Fields.ProjectCardID]: projcard,
            [Fields.Comment]: payload.comment,
            [Fields.BuildObjForFile]: state.projectItem[Fields.BuildObj],
        };

        if (payload.docType) {
            item[Fields.PartName] = payload.docType;
        }
        commit(SET_FILE_ITEM_DATA, { fileName: payload.fileName, item });
    },
    [actionTypes.SET_DOC_TYPES]({ commit }, items: SP.ListItem[]) {
        const docTypePatterns = new Array<DocTypeSearchPattern>();
        const docTypes = new Array<DocType>();
        // Add default empty DocType
        docTypes.push({ id: -1, title: '', sortNum: -1 });
        items.forEach((element) => {
            // Get pattern string from listitem
            const pnum: string = element.get_item('PartNum');
            const pnums: string[] = pnum.split('.');

            docTypes.push({
                id: element.get_id(),
                title: element.get_item('Title'),
                partNum: pnum,
                sortNum: pnums && pnums.length > 0 ? Number(pnums[0]) : Number(0),
            });
            const pattern: string = element.get_item('SearchPattern');
            if (pattern && pattern !== '') {
                // Split pattern string
                const patterns = pattern.split(';') as string[];
                if (patterns !== null && patterns.length > 0) {
                    const id = element.get_id();
                    patterns.forEach((p) => {
                        if (p !== '') {
                            docTypePatterns.push({ itemId: id, pattern: p });
                        }
                    });
                }
            }
        });
        commit(SET_DOC_TYPE_PATTERN, docTypePatterns);
        commit(SET_PART_FILE_NAMES, docTypes.
            // Sort by number
            sort((a, b) => {
            const an: number = Number(a.sortNum);
            const bn: number = Number(b.sortNum);
            if (an > bn) {
                return 1;
            }
            if (an < bn) {
                return -1;
            }
            if (an === bn) {
                return 0;
            }
            return -1;
        }));
    },
    [actionTypes.CHANGE_FILE_ACTION]({ commit, state }, payload: {
        fileName: string, docType: DocType, info?: string,
    }) {
        const file = state.files.find((f) => f.fileName === payload.fileName);
        if (!file) {
            throw new Error('File can not be null!');
        }
        commit(CHANGE_FILE, payload);
    },
    [actionTypes.SAVE_FILE]({ commit, state }, payload: {
        siteUrl: string, listId: string, checkinType: SP.CheckinType, file: StateFile
    }) {
        if (!payload.file.saved) {
            // Apply changes first.
            commit(APPLY_CHANGES, payload.file);
            SPDataService.Current().saveFile(
                payload.siteUrl,
                payload.listId,
                payload.file,
                payload.checkinType,
            ).then((f) => {
                commit(SAVE_FILE, payload.file);
            }).catch((e) => {
                throw e;
            });
        }
    },
    [actionTypes.SAVE_FILES]({ commit, state }, payload: {
        siteUrl: string, listId: string, checkinType: SP.CheckinType,
    }) {
        if (state.files && state.files.length > 0) {
            state.files.forEach((file) => {
                if (!file.saved) {
                    // Apply changes first.
                    commit(APPLY_CHANGES, file);
                    SPDataService.Current().saveFile(
                        payload.siteUrl,
                        payload.listId,
                        file,
                        payload.checkinType,
                    ).then((f) => {
                        commit(SAVE_FILE, file);
                    }).catch((e) => {
                        throw e;
                    });
                }
            });
        }
    },
    [actionTypes.SET_PROJECT_ITEM]({ commit, state }, item: SP.ListItem) {
        const project: ItemData = {
            id: item.get_id(),
        };
        const oi = item.get_fieldValues();
        Object.keys(item.get_fieldValues()).forEach((key) => {
            if (key !== 'id') {
                project[key] = oi[key];
            }
        });
        commit(SET_PROJECT_ITEM, project);
    },
    [actionTypes.SET_UPLOAD_MODE]({ commit }, mode: UploadMode) {
        if (mode !== UploadMode.Unknown) {
            commit(SET_UPLOAD_MODE, mode);
        } else {
            throw new Error('Unknown Upload mode');
        }
    },
    async [actionTypes.DELETE_FILE]({ commit, state }, file: StateFile) {
        if (!file) {
            throw new ReferenceError('File can not be null or undefined.');
        }
        const sf = state.files.find((f) => f.fileName === file.fileName);
        if (!sf) {
            throw new ReferenceError('File with name ' + file.fileName +
                'does not exists in collection.');
        }
        await SPDataService.Current().deleteFile(state.options.siteUrl, file.serverRelativeUrl);
        commit(DELETE_FILE, { file });
    },
    [actionTypes.SET_OPTIONS]({ commit, state }, options: AppOptions) {
        commit(SET_OPTIONS, options);
    },
    // [actionTypes.SET_FILE_UNSAVED]({ commit, state }, file: StateFile){
    //     if (!file) {
    //         throw new ReferenceError('File can not be null or undefined.');
    //     }
    //     const sf = state.files.find((f) => f.fileName === file.fileName);
    //     if (!sf) {
    //         throw new ReferenceError('File with name ' + file.fileName +
    //             'does not exists in collection.');
    //     }
    //     commit('', { file });
    // },
};
