<template>
  <div>
    <div class="ms-toolbar ms-alignLeft" v-if="mode === 2">
      <button type="button" @click="onclose">OK</button>&nbsp;
      <!-- <button type="button" @click="onclose">Закрыть</button> -->
    </div>
    <div class="ms-toolbar ms-alignLeft" v-if="mode === 1">
      <button type="button" @click="onsave" :disabled="!hasUnsavedFiles">Сохранить файлы</button>&nbsp;
      <button type="button" @click="onclose">Закрыть</button>
    </div>
    <div class="ms-toolbar ms-alignLeft" v-if="mode === 1">
      <div class="ms-list-addnew ms-list-addnew-aligntop ms-textXLarge ms-alignLeft">
        <span>Перетащите файлы сюда</span>
      </div>
    </div>
    <FileList :siteUrl="config.siteUrl" :listId="config.listId"></FileList>
    <div class="ms-toolbar ms-alignLeft" v-if="mode === 2">
      <button type="button" @click="onclose">OK</button>&nbsp;
      <!-- <button type="button" @click="onclose">Закрыть</button> -->
    </div>
    <div class="ms-toolbar ms-alignLeft" v-if="mode === 1">
      <button type="button" @click="onsave" :disabled="!hasUnsavedFiles">Сохранить файлы</button>&nbsp;
      <button type="button" @click="onclose">Закрыть</button>
    </div>
  </div>
</template>
<script lang="ts">
import utils from "../utils";
import { AppOptions, StateFile, UploadMode } from "../types";
import { Component, Prop, Vue, Watch } from "vue-property-decorator";
import {
  mapState,
  mapActions,
  ActionMethod,
  mapGetters,
  Computed,
  Getter
} from "vuex";
import actions from "../store/action-types";
import store from "@/store/store";
import FileList from "./FileList.vue";
import OnSaveDialog from "./OnSaveDialog.vue";
import { SPScripts } from "@/utils/constants";
import { createOnSaveDialogElement } from "./inline";

@Component({
  computed: {
    ...mapState({
      counter: number => store.state.counter,
      mode: mode => store.state.mode
    }),
    hasUnsavedFiles() {
      const files: Array<StateFile> = this.$store.state.files as Array<
        StateFile
      >;
      const savedFiles = files.filter(f => !f.saved);
      return savedFiles.length > 0;
    },
    ...mapGetters(["getUnsavedFiles"])
  },
  methods: {
    ...mapActions({
      setCounter: "setCounter",
      uploadFiles: actions.UPLOAD_FILES,
      saveFiles: actions.SAVE_FILES
    })
  },
  components: {
    FileList
  }
})
export default class UploadFiles extends Vue {
  @Prop()
  public config!: AppOptions;
  // State computed
  public counter!: number;
  public hasUnsavedFiles!: boolean;
  // State Actions
  public setCounter!: ActionMethod;
  public uploadFiles!: ActionMethod;
  public saveFiles!: ActionMethod;
  public notify!: string;
  public mode!: UploadMode;

  public getUnsavedFiles!: () => Array<StateFile>;

  public mounted() {

    // Initialize uploader
    if (this.$store.state.mode !== UploadMode.Edit) {
      const self = this;
      SP.SOD.executeFunc(utils.SPScripts.DragDrop.Script, 'registerDragUpload', () => {
        //console.log('SOD WORKS')
        registerDragUpload(
          self.$el,
          self.config.serverUrl,
          self.config.siteUrl,
          self.config.listId,
          self.config.folderUrl,
          false,
          false,
          () => {},
          self.preUpload,
          self.postUpload,
          self.checkPermissions
        );
      });
    }
  }

  public preUpload(files: FileElement[]) {}
  public postUpload(files: FileElement[]) {
    const self = this;

    this.uploadFiles({
      siteUrl: self.$store.state.options.siteUrl,
      listId: self.$store.state.options.listId,
      files,
      folderUrl: self.$store.state.options.folderUrl
    })
      .then(value => {
        SP.SOD.executeFunc(SPScripts.SP.Script, SPScripts.SP.UI.Status, () => {
          const status = SP.UI.Status.removeAllStatus(true);
        });
      })
      .catch((e: any) => {
        console.error(e);
      });
  }

  public checkPermissions(): boolean {
    return true;
  }

  public onsave() {
    const self = this;
    SP.SOD.executeFunc(SPScripts.SP.Script, SPScripts.SP.UI.Notify, () => {
      self.notify = SP.UI.Notify.addNotification(
        "Выполняем загрузку файлов",
        true
      );
    });

    this.saveFiles({
      siteUrl: self.config.siteUrl,
      listId: self.config.listId,
      checkinType: SP.CheckinType.majorCheckIn
    })
      .then(a => {
        SP.SOD.executeFunc(SPScripts.SP.Script, SPScripts.SP.UI.Status, () => {
          SP.UI.Notify.removeNotification(self.notify);
          const status = SP.UI.Status.addStatus(
            "Файлы успешно сохранены в системе",
            ""
          );
          SP.UI.Status.setStatusPriColor(status, "green");
        });
      })
      .catch(e => {
        SP.SOD.executeFunc(SPScripts.SP.Script, SPScripts.SP.UI.Status, () => {
          SP.UI.Notify.removeNotification(self.notify);
          const status = SP.UI.Status.addStatus(
            "При сохранении файлов возникли неожиданные ошибки",
            e
          );
          SP.UI.Status.setStatusPriColor(status, "red");
        });
      });
  }

  public onclose() {
    if (this.hasUnsavedFiles) {
      const self = this;
      const files = this.getUnsavedFiles();
      SP.SOD.executeFunc(
        SPScripts.SP.UI.Dialog.Script,
        SPScripts.SP.UI.Dialog.ShowModalDialog,
        // Callback start
        () => {
          // Open Save Dialog
          SP.UI.ModalDialog.showModalDialog({
            autoSize: true,
            title: "Сохранить файлы",
            html: createOnSaveDialogElement(
              "_onSaveDialog",
              {
                instructions:
                  "Нажмите кнопку 'сохранить', чтобы сохранить изменения или отмена, чтобы выйти без сохранения.",
                buttonOk: "Сохранить",
                subject: "Следующие файлы были изменены:"
              },
              files.map(file => {
                return file.fileName;
              })
            ),
            // Callback dialog return
            dialogReturnValueCallback: (
              dialogResult: SP.UI.DialogResult,
              returnValue: any
            ) => {
              if (dialogResult === SP.UI.DialogResult.OK) {
                // Save Files
                this.saveFiles({
                  siteUrl: self.$store.state.options.siteUrl,
                  listId: self.$store.state.options.listId,
                  checkinType: SP.CheckinType.majorCheckIn
                })
                  .then(a => {
                    // saved
                    self.redirect();
                  })
                  .catch(e => {
                    SP.SOD.executeFunc(
                      SPScripts.SP.Script,
                      SPScripts.SP.UI.Status,
                      () => {
                        SP.UI.Notify.removeNotification(self.notify);
                        const status = SP.UI.Status.addStatus(
                          "При сохранении файлов возникли неожиданные ошибки",
                          e
                        );
                        // Set Error status
                        SP.UI.Status.setStatusPriColor(status, "red");
                      }
                    );
                  });
              } else {
                // without save
                self.redirect();
              }
            } // End callback dialog return
          });
        } // Callback ends
      );
    } else {
      this.redirect();
    }
  }

  redirect() {
    const retUrl = this.$store.state.options.returnUrl;
    if (retUrl) {
      window.location.href = retUrl;
    } else {
      window.location.href = "/";
    }
  }
}
</script>
<style scoped>
.db-margin-bottom-15 {
  margin-bottom: 15px;
}
.db-margin-top-15 {
  margin-top: 15px;
}
</style>
