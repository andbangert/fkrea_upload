import Vue from 'vue';
import App from './App.vue';
import store from './store/store';
import { AppOptions, UploadMode } from './types';
import { SPDataService } from './utils';
import Utils from './utils';
import { applyPolyfills } from './utils/polyfills';
import actionTypes from './store/action-types';
import './scss/main.scss';
import { SPScripts } from './utils/constants';

Vue.config.productionTip = false;
applyPolyfills();


export function InitAppWithDialog(config?: AppOptions) {
  SP.SOD.executeFunc(SPScripts.SP.UI.Dialog.Script, 'SP.UI.ModalDialog.showWaitScreenWithNoClose', () => {
    const loadingDialog = SP.UI.ModalDialog.showWaitScreenWithNoClose('Пожалуйста подождите', 'Загружаем файлы', 150, 500);
    try {
      InitApp(config).then((val) => {
        loadingDialog.close(SP.UI.DialogResult.OK);
      }).catch((e) => {
        loadingDialog.close(SP.UI.DialogResult.cancel);
      });
    } catch (e) {
      console.error(e);
      loadingDialog.close(SP.UI.DialogResult.cancel);
    }
  });
}

// Main Entry
export async function InitApp(config?: AppOptions) {
  console.log('Initialize Application');
  if (!config) {
    config = {
      serverUrl: 'http://vm-arch',
      siteUrl: '/sites/documentation',
      listId: '{CA9F91AA-D391-41B8-A04C-C9D676CE2136}',
      folderUrl: '/documents',
      docPartNameListId: '{8b1d74ae-97bd-44a9-a314-6a48903df849}',
      projectListId: '{d0a9d56c-4d8a-43b1-9d0a-ceb123ec9b54}',
      projectId: -1,
    };
  }

  const rurl = GetUrlKeyValue('returnUrl');
  const mode = GetUrlKeyValue('mode');
  let path: string = '';

  if (rurl) {
    config.returnUrl = rurl;
  }

  await store.dispatch(actionTypes.SET_OPTIONS, config);
  // Fetch Project

  const r = await InitializeProjectItem(config.projectId, config.projectListId, config.siteUrl);
  if (r.items.length > 0) {
    path = r.items[0].get_item('Path') as string;
    if (path === '') {
      path = r.items[0].get_item('Title') as string;
    }
    await store.dispatch(actionTypes.SET_PROJECT_ITEM, r.items[0]);
  }
  // ========================

  // Initialize folder Url
  if (!path) {
    throw new Error('Folder path can not be null.');
  }

  const folderUrl = `${config.folderUrl}/${path}`;
  config.folderUrl = folderUrl;
  // ========================

  // Initialize Doc Types
  const spsvc = new SPDataService();
  const dct = await spsvc.getItemsAsync(config.siteUrl, config.docPartNameListId, {
    query: '<Query><OrderBy><FieldRef Name="PartNum" Ascending="True" /></OrderBy></Query>',
  })
  await store.dispatch('setDocTypes', dct.items);
  // ========================

  const titleDoc: HTMLElement | null | undefined = document.getElementById('DeltaPlaceHolderPageTitleInTitleArea');
  // Initialize primary files data
  console.log(mode)
  if (mode && (mode === 'Edit' || mode === '1')) {
    if (titleDoc) {
      titleDoc.innerText = `${path}`;
    }
    //Initialise files
    await store.dispatch(actionTypes.SET_UPLOAD_MODE, UploadMode.Edit);
    await store.dispatch(actionTypes.FETCH_FILES, {
      siteUrl: config.siteUrl, listId: config.listId, folderUrl,
    });
  } else {
    if (titleDoc) {
      titleDoc.innerText = `${path}`;
    }
    // Create folder
    await store.dispatch(actionTypes.SET_UPLOAD_MODE, UploadMode.New);
    const f = await SPDataService.Current().createFolder(config.siteUrl, config.listId, folderUrl);
  }
  // ========================

  // Initialize App
  const v1 = new Vue({
    el: '#app',
    components: {
      App,
    },
    store,
    render: (h) => h(App, {
      props: {
        config,
      },
    }),
  }).$mount('#app');
  // ========================
}

export function InitializeProjectItem(projectId: number, listId: string, siteUrl: string) {
  if (projectId === undefined || projectId <= 0) {
    throw Error('Project Id can not be null.');
  }
  let project: SP.ListItem;
  const spsvc = new SPDataService();
  return spsvc.getItemsAsync(siteUrl, listId, {
    query: '<Query><Where><Eq><FieldRef Name="ID" Ascending="True" />'
      + `<Value Type='Number'>${projectId}</Value></Eq></Where></Query>`,
  });
}

export function OpenUploadProjectFilesDialog() {
  SP.SOD.executeFunc(Utils.SPScripts.SP_UI_Dialog.Script, Utils.SPScripts.SP_UI_Dialog.ShowModalDialog, () => {
    const projectId = GetUrlKeyValue('ID');
    if (!projectId) {
      throw Error('Project Id can not be null.');
    }
    const listId = _spPageContextInfo.pageListId;
    const siteUrl = _spPageContextInfo.siteServerRelativeUrl;
    const absoluteSiteUrl = _spPageContextInfo.webAbsoluteUrl;
    const spsvc = new SPDataService();
    spsvc.getItemsAsync(siteUrl, listId, {
      query: '<Query><Where><Eq><FieldRef Name="ID" Ascending="True" />'
        + `<Value Type='Number'>${projectId}</Value></Eq></Where></Query>`,
    }).then((r) => {
      const options = SP.UI.$create_DialogOptions();
      options.autoSize = true;
      options.showClose = true;
      options.title = 'Добавить файлы к проекту';
      options.url = `${absoluteSiteUrl}/SitePages/UploadFilesPage.aspx`;
      SP.UI.ModalDialog.showModalDialog(options);
    }).catch((e) => {
      // TODO: Handle Error correctly
      throw new Error(e);
    });
  });
}

export function RedirectUploadFilesPage(projectId: number, scanLibraryId: string, scanFolderUrl: string, mode?: string) {
  const listId = _spPageContextInfo.pageListId;
  const siteUrl = _spPageContextInfo.siteServerRelativeUrl;
  const absoluteSiteUrl = _spPageContextInfo.webAbsoluteUrl;

  const libId = encodeURI(scanLibraryId);
  const scanF = encodeURI(scanFolderUrl);

  const returnUrl = encodeURI(_spPageContextInfo.serverRequestPath + '?ID=' + projectId);
  location.href = `${absoluteSiteUrl}/SitePages/UploadFilesPage.aspx?pid=${projectId}&mode=${mode}&returnUrl=${returnUrl}`;
}

(function () {
  RegisterModuleInit('app.js', () => {
    // Initialize properties here
  });
  NotifyScriptLoadedAndExecuteWaitingJobs('app.js');
});

// <script type="text/javascript" >
//   document.write('<div id="app"></div>');

// if (SP && SP.SOD) {
//   SP.SOD.executeFunc('app.js', null, function () {
//     try {
//       var projectId = Number(GetUrlKeyValue('pid'));
//       dynamicapp.InitApp({
//         serverUrl: "http://fkrea",
//         siteUrl: "/sites/documentation",
//         listId: "{F966AE2C-75F9-4E64-B37A-C5C2B9E408C8}",
//         folderUrl: "/ProjectScan",
//         docPartNameListId: "{74B4A326-8880-4D15-8907-A1904EEE13FD}",
//         projectListId: "{C3C9E598-C6F1-49EF-881B-DFA95013DF73}",
//         projectId: Number(projectId),
//       });
//     } catch (e) {
//       console.error(e);
//     }
//   });
// }
// </script>

// <script type="text/javascript">
// document.write('<div id="app"></div>');

// if (SP && SP.SOD) {
//   SP.SOD.executeFunc('app.js', null, function () {
//     var projectId = Number(GetUrlKeyValue('pid'));
//     dynamicapp.InitAppWithDialog({
//       serverUrl: "http://vm-arch",
//       siteUrl: "/sites/documentation",
//       listId: "{70e129bd-68ac-4b61-8242-133d6eb48c3e}",
//       folderUrl: "/ProjectScan",
//       docPartNameListId:"{8b1d74ae-97bd-44a9-a314-6a48903df849}",
//       projectListId: "{d0a9d56c-4d8a-43b1-9d0a-ceb123ec9b54}",
//       projectId: Number(projectId),
//     });
//   });
// }
// </script>


// <script type='text/javascript'>
// document.write('<div id='app'></div>');

// if (SP && SP.SOD) {
//   SP.SOD.executeFunc('app.js', null, function () {
//     dynamicapp.InitApp({
//       serverUrl: 'http://vm-arch',
//       siteUrl: '/sites/documentation',
//       listId: '{CA9F91AA-D391-41B8-A04C-C9D676CE2136}',
//       folderUrl: '/documents',
//       docPartNameListId:'{74B4A326-8880-4D15-8907-A1904EEE13FD}'
//     });
//   });
// }
// </script>

// <script type="text/javascript">
// document.write('<div id="app"></div>');

// if (SP && SP.SOD) {
//   SP.SOD.executeFunc('app.js', null, function () {
//     dynamicapp.InitApp({
//       serverUrl: "http://vm-arch",
//       siteUrl: "/sites/documentation",
//       listId: "{70e129bd-68ac-4b61-8242-133d6eb48c3e}",
//       folderUrl: "/ProjectScan",
//       docPartNameListId:"{8b1d74ae-97bd-44a9-a314-6a48903df849}"
//     });
//   });
// }
// </script>
