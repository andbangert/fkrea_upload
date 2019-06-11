import { SPScripts } from './constants';
import { applyPolyfills } from './polyfills';
import axios from 'axios';
import { FetchedFileData, ItemData, StateFile } from '@/types';

function trimHtmlTags(html: string) {
    const ele = document.createElement('div');
    ele.innerHTML = html;
    return ele.innerText;
}

function getUrlParameter(name: string): string {
    name = name.replace(/[\[]/, '\\[').replace(/[\]]/, '\\]');
    const regex = new RegExp('[\\?&]' + name + '=([^&#]*)');
    const results = regex.exec(location.search);
    return results === null ? '' : decodeURIComponent(results[1].replace(/\+/g, ' '));
}

function executionContext(fileName: string, func: () => void): void;
function executionContext(fileName: string, func: () => void, typeName?: string): void {
    if (SP && SP.SOD) {
        SP.SOD.executeFunc(fileName, typeName ? typeName : '', func);
    }
}

declare interface ListItemCollection {
    items: SP.ListItem[];
    position?: string;
}

declare interface SPQuery {
    query: string;
    rowLimit?: number;
    position?: string;
}

export class SPDataService {
    public static Current() {
        if (!SPDataService.current) {
            SPDataService.current = new SPDataService();
        }
        return SPDataService.current;
    }

    public static async mapToIcon(siteUrl: string, fileName: string, size: SP.Utilities.IconSize): Promise<string> {
        return new Promise<string>((resolve, reject) => {
            const ctx = new SP.ClientContext(siteUrl);
            const iconobj = ctx.get_web().mapToIcon(fileName, '', size);
            ctx.executeQueryAsync(() => {
                const icon = iconobj.get_value();
                resolve(icon && icon !== '0' ?
                    `${_spPageContextInfo.webAbsoluteUrl}/_layouts/15/images/${icon}` :
                    `${_spPageContextInfo.webAbsoluteUrl}/_layouts/15/images/icgen.gif`);
            });
            ctx.dispose();
        }).then((r) => {
            return r;
        });
    }

    private static current: SPDataService;

    public async getItemsAsync(siteUrl: string, listId: string, query: SPQuery): Promise<ListItemCollection> {
        return await this.fetchItemsByJSOMAsync(siteUrl, listId, query.query, query.rowLimit, query.position);
    }

    public deleteFile(siteUrl: string, serverRelativeFileUrl: string): Promise<string | SP.ClientRequestFailedEventArgs> {
        // async jquery
        const promise: Promise<string | SP.ClientRequestFailedEventArgs> = new Promise<string | SP.ClientRequestFailedEventArgs>((resolve, reject) => {

            SP.SOD.executeFunc(SPScripts.SP.Script, SPScripts.SP.ClientContextType, () => {
                const clientcontext = new SP.ClientContext(siteUrl);
                const web = clientcontext.get_web();
                // load root folder
                const file: SP.File = web.getFileByServerRelativeUrl(serverRelativeFileUrl);
                file.deleteObject();
                clientcontext.executeQueryAsync((sender, args) => {
                    // success
                    resolve(serverRelativeFileUrl);
                }, (sender, args: SP.ClientRequestFailedEventArgs) => {
                    console.log(args);
                    // failed
                    reject(args);
                    // console.error('(fetchFileByJSOM) Error occured while fetch data from SharePoint.', args);
                });
                clientcontext.dispose();
            });
        });
        return promise;
    }

    public getFiles(siteUrl: string, folderServerRelativeUrl: string): Promise<SP.FileCollection> {
        return new Promise((resolve, reject) => {
            const ctx = new SP.ClientContext(siteUrl);
            const web = ctx.get_web();
            const folder = web.getFolderByServerRelativeUrl(folderServerRelativeUrl);
            const files = folder.get_files();
            ctx.load(folder);
            ctx.load(files);
            ctx.executeQueryAsync((sender, args) => {
                // load all items
                if (files && files.get_count() > 0) {
                    const num = files.get_count();
                    const iterator = files.getEnumerator();
                    while (iterator.moveNext()) {
                        ctx.load(iterator.get_current().get_listItemAllFields());
                    }
                    iterator.reset();
                    ctx.executeQueryAsync((s, a) => {
                        resolve(files);
                    }, (s, a) => {
                        throw new Error(a.get_errorDetails() + ' ' + a.get_stackTrace());
                    });
                }
            }, (sender, args) => {
                reject(args);
            });
            ctx.dispose();
        });
    }

    public createFolder(siteUrl: string, listId: string, folderUrl: string): Promise<SP.Folder> {
        return new Promise<SP.Folder>((resolve, reject) => {
            const ctx = new SP.ClientContext(siteUrl);
            const list = ctx.get_web().get_lists().getById(listId);
            const createFolderInternal = (parentFolder: SP.Folder, fUrl: string) => {
                const ctx1 = parentFolder.get_context();
                const folderNames = fUrl.split('/');
                const folderName = folderNames[0];
                const curFolder = parentFolder.get_folders().add(folderName);
                ctx1.load(curFolder);
                ctx1.executeQueryAsync((s, a) => {
                    if (folderNames.length > 1) {
                        const subFolderUrl = folderNames.slice(1, folderNames.length).join('/');
                        createFolderInternal(curFolder, subFolderUrl);
                    }
                    resolve(curFolder);
                }, (s, a) => {
                    reject(a);
                });
                // end
                ctx1.dispose();
            };
            createFolderInternal(list.get_rootFolder(), folderUrl);
            ctx.dispose();
        });
    }

    public saveFile(siteUrl: string, listId: string, sfile: StateFile, checkinType?: SP.CheckinType): Promise<SP.File | SP.ClientRequestFailedEventArgs> {
        const promise = new Promise<SP.File | SP.ClientRequestFailedEventArgs>((resolve, reject) => {
            const item = sfile.item;
            // Initialize context
            const clientContext = new SP.ClientContext(siteUrl);
            const web = clientContext.get_web();
            const list = web.get_lists().getById(listId);
            // Get File
            const file: SP.File = web.getFileByServerRelativeUrl(sfile.serverRelativeUrl);
            // load file
            clientContext.load(file);
            clientContext.executeQueryAsync((sender, args) => {
                // Check out
                if (file.get_checkOutType() === SP.CheckOutType.none) {
                    // checkout for edit
                    file.checkOut();
                }
                // Get file sp list item
                const spitem = file.get_listItemAllFields();
                // Set items Properties
                if (item && item !== null) {
                    Object.keys(item).forEach((key) => {
                        if (key !== 'id') {
                            spitem.set_item(key, item[key]);
                        }
                    });
                    spitem.update();
                }
                // Load Item
                clientContext.load(spitem);
                clientContext.executeQueryAsync((s, a) => {
                    // Checkin and publish
                    const chc = sfile.info ? sfile.info : '';
                    file.checkIn(chc, checkinType ? checkinType : SP.CheckinType.majorCheckIn);
                    file.publish(chc);
                    clientContext.executeQueryAsync((s1, a1) => {
                        resolve(file);
                    }, (s1, a1) => {
                        reject(a1);
                        // console.error('(CheckinPublishJSOM) Error occured while fetch data from SharePoint.', a1);
                    });
                }, (s, a) => {
                    reject(a);
                    // console.error('(SetItemJSOM) Error occured while fetch data from SharePoint.', a);
                });
                // Load iten
            }, (sender, args) => {
                // failed
                reject(args);
                // console.error('(FetchFileJSOM) Error occured while fetch data from SharePoint.', args);
            });
            // load file
        });
        return promise;
    }

    public saveItem(siteUrl: string, listId: string, item: ItemData): Promise<SP.ListItem> {
        const promise = new Promise<SP.ListItem>((resolve, reject) => {
            const clientContext = new SP.ClientContext(siteUrl);
            const list = clientContext.get_web().get_lists().getById(listId);
            const spitem = list.getItemById(item.id);
            Object.keys(item).forEach((key) => {
                if (key !== 'id') {
                    spitem.set_item(key, item[key]);
                }
            });
            spitem.update();
            clientContext.executeQueryAsync((sender, args) => {
                resolve(spitem);
            }, (sender, args) => {
                // console.error(args);
            });
            clientContext.dispose();
        });
        return promise;
    }

    public getFile(siteUrl: string, serverRelativeFileUrl: string): Promise<FetchedFileData> {
        // async jquery
        const promise: Promise<FetchedFileData> = new Promise<FetchedFileData>((resolve, reject) => {

            SP.SOD.executeFunc(SPScripts.SP.Script, SPScripts.SP.ClientContextType, () => {
                const clientcontext = new SP.ClientContext(siteUrl);
                const web = clientcontext.get_web();

                // load root folder
                const file: SP.File = web.getFileByServerRelativeUrl(serverRelativeFileUrl);
                const item = file.get_listItemAllFields();
                // load file
                clientcontext.load(file);
                clientcontext.load(item);

                const iconobj = web.mapToIcon(serverRelativeFileUrl, '', SP.Utilities.IconSize.size16);
                clientcontext.executeQueryAsync((sender, args) => {
                    // success
                    const icon = iconobj.get_value();
                    const iconUrl = icon && icon !== '0' ?
                        `${_spPageContextInfo.webAbsoluteUrl}/_layouts/15/images/${icon}` :
                        `${_spPageContextInfo.webAbsoluteUrl}/_layouts/15/images/icgen.gif`;

                    resolve({
                        file, // : file,
                        listItem: item,
                        iconUrl, // : iconUrl
                    });
                }, (sender, args) => {
                    // failed
                    reject(args);
                    // console.error('(fetchFileByJSOM) Error occured while fetch data from SharePoint.', args);
                });
                clientcontext.dispose();
            });
        });
        return promise;
    }

    //#region [ private methods ]
    private fetchItemsByJSOMAsync(siteUrl: string, listId: string, query: string,
        rowLimit?: number, position?: string): Promise<ListItemCollection> {
        // async jquery
        const promise: Promise<ListItemCollection> = new Promise<ListItemCollection>((resolve, reject) => {

            SP.SOD.executeFunc(SPScripts.SP.Script, SPScripts.SP.ClientContextType, () => {
                const clientContext = new SP.ClientContext(siteUrl);
                const web = clientContext.get_web();

                const list = web.get_lists().getById(listId);
                const caml = new SP.CamlQuery();

                caml.set_viewXml(this.buildCaml(query, rowLimit));
                // Sets current position
                if (!position || position !== '') {
                    const posObj = new SP.ListItemCollectionPosition();
                    posObj.set_pagingInfo(position + '');
                    caml.set_listItemCollectionPosition(posObj);
                }

                const colItems = list.getItems(caml);

                clientContext.load(colItems);
                clientContext.executeQueryAsync((sender, args) => {
                    // success
                    const iterator = colItems.getEnumerator();
                    const curPosObj = colItems.get_listItemCollectionPosition();
                    const curPos: string = curPosObj ? curPosObj.get_pagingInfo() : '';

                    const data = new Array<SP.ListItem>();
                    // iterate through items
                    while (iterator.moveNext()) {
                        const item = iterator.get_current();
                        // console.log(item);
                        // console.log(item.get_id());
                        data.push(item);
                    }
                    // Resolve operation
                    // d.resolve({ posts: data, position: curPos });
                    resolve({ items: data, position: curPos });
                }, (sender, args) => {
                    // failed
                    reject(args);
                    // console.error('(fetchItemsByJSOM) Error occured while fetch data from SharePoint.', args);
                });
                clientContext.dispose();
            });
        });
        return promise;
    }

    private buildCaml(caml: string, rowLimit?: number): string {
        if (rowLimit && rowLimit > 0) {
            return `<View>${caml}<RowLimit>${rowLimit}</RowLimit></View>`;
        } else {
            return `<View>${caml}</View>`;
        }
    }

    // REST API
    private getListItemsByRestApi(baseUrl: string, listid: string, baseQuery: string, columns: string) {
        const query = JSON.stringify({ ViewXml: baseQuery });
        const ajaxurl = `${baseUrl}'/_api/web/lists(guid\''${listid}'\')/
            getitems(query=@q)?@q='${query}'&$select='${columns}`;

        let data: any = null;
        this.spSecurityContextAsync(baseUrl).then((securityToken: string) => {
            axios.post(ajaxurl, null, {
                headers: {
                    'content-type': 'application/json; odata=verbose',
                    'Accept': 'application/json; odata=verbose',
                    'X-RequestDigest': securityToken,
                },
                responseType: 'json',
            }).then((d: any) => {
                data = d;
                // console.log(data);
            });
        });
        return data;
    }

    private spSecurityContext(baseUrl: string): string {
        let token: string = '';
        axios.post(baseUrl + '/_api/contextinfo', {},
            {
                headers: { 'Accept': 'application/json; odata=verbose' },
            }).then((tokenData: any) => {
                // console.log(tokenData);
                token = tokenData.d.GetContextWebInformation.FormDigestValue;
            }).catch((e) => {
                throw new Error(e);
            });
        return token;
    }

    private spSecurityContextAsync(baseUrl: string): Promise<string> {
        const promise: Promise<string> = new Promise<string>((resolve, reject) => {
            axios.post(baseUrl + '/_api/contextinfo', {},
                {
                    headers: { 'Accept': 'application/json; odata=verbose' },
                }).then((tokenData: any) => {
                    // console.log(tokenData);
                    const token: string = tokenData.d.GetContextWebInformation.FormDigestValue;
                    if (!token || token === '') {
                        reject(new Error('Security token is empty'));
                    }
                    resolve(token);
                }).catch((e) => {
                    reject(e);
                });
        });
        return promise;
    }
    //#endregion
}

export default {
    applyPolyfills,
    trimHtmlTags,
    executionContext,
    SPScripts,
    getUrlParameter,
};
