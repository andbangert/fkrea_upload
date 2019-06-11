<template>
  <tr>
    <td class="ms-cellstyle ms-vb-icon" height="100%">
      <span>
        <img :src="file.iconUrl" alt>
      </span>
    </td>
    <td class="ms-cellstyle ms-vb-title" height="100%">
      <div class="ms-vb">
        <span>{{file.fileName}}</span>
      </div>
    </td>
    <td class="ms-cellstyle ms-vb-title" height="100%">
      <div class="ms-vb">
        <span>{{file.lastModified.toLocaleString()}}</span>
      </div>
    </td>
    <td class="ms-cellstyle ms-vb-title" height="100%">
      <div class="ms-vb">
        <span>{{file.fileSize / 1000}}&nbsp;kb</span>
      </div>
    </td>
    <td class="ms-cellstyle ms-vb-title" height="100%">
      <div class="ms-vb">
        <span>{{file.majorVersion}}.{{file.minorVersion}}</span>
      </div>
    </td>
    <td class="ms-cellstyle ms-vb-title" height="100%">
      <div class="ms-vb">
        <select
          name="docType"
          @change="onchange($event)"
          class="db-select ms-RadioText"
          :disabled="saved"
          v-model="docType"
        >
          <option
            v-for="item in docTypes"
            :key="item.id"
            :value="item"
            :selected="item.id === file.docType.id"
          >{{item.partNum}}&nbsp;-&nbsp;{{item.title}}</option>
        </select>
      </div>
    </td>
    <td class="ms-cellstyle ms-vb-title" height="100%">
      <div class="ms-vb">
        <input
          v-model.lazy="addInfo"
          @change="onchange"
          class="ms-input ms-inputBox db-input"
          :disabled="saved"
        >
      </div>
    </td>
    <td class="ms-cellstyle ms-vb-icon" height="100%" v-if="mode === 2">
      <div class="ms-vb">
        <SaveEditButton @active-changed="onEditChanged"></SaveEditButton>
      </div>
    </td>
  </tr>
</template>
<script lang="ts">
import utils, { SPDataService } from "../utils";
import {
  AppOptions,
  StateFile,
  DocType,
  UploadMode,
  SaveEditButtonKey,
  StateFileType
} from "../types";
import { Component, Prop, Vue, Watch, Model } from "vue-property-decorator";
import { mapState, mapActions, ActionMethod } from "vuex";
import SaveEditButton from "./SaveEditButton.vue";
import actions from "../store/action-types";
import store from "@/store/store";
import { SPScripts } from "@/utils/constants";
import { createOnSaveDialogElement } from "./inline";

@Component({
  // props: ['file'],
  computed: {
    ...mapState({
      docTypes: docTypes => store.state.docTypes,
      mode: mode => store.state.mode
    })
  },
  methods: {
    ...mapActions({
      changeFile: actions.CHANGE_FILE_ACTION,
      saveFile: actions.SAVE_FILE,
      deleteFile: actions.DELETE_FILE
    })
  },
  components: {
    SaveEditButton
  }
})
export default class File extends Vue {
  @Prop()
  public file!: StateFile;
  public mode!: UploadMode;
  public docTypes!: DocType[];
  public addInfo?: string = this.file.info;
  public docType: DocType = this.file.docType;

  public changeFile!: ActionMethod;
  public saveFile!: ActionMethod;
  public deleteFile!: ActionMethod;
  public saved: boolean = this.file.saved;
  public notify!: string;

  public onEditChanged(arg: SaveEditButtonKey) {
    if (arg === SaveEditButtonKey.Edit) {
      this.saved = false;
    } else if (arg === SaveEditButtonKey.Cancel) {
      const self = this;
      this.addInfo = self.file.info;
      this.docType = self.file.docType;
      this.saved = true;
    } else if (arg === SaveEditButtonKey.Remove) {
      const self = this;
      SP.UI.ModalDialog.showModalDialog({
        autoSize: true,
        title: "Удаление файла",
        html: createOnSaveDialogElement(
          "_onSaveDialog",
          {
            instructions:
              "Нажмите кнопку 'Удалить', чтобы удалить файл или нажмите 'Отмена'.",
            buttonOk: "Удалить",
            subject: "Вы действительно хотите удалить файл:"
          },
          [self.file.fileName]
        ),
        // Callback dialog return
        dialogReturnValueCallback: (
          dialogResult: SP.UI.DialogResult,
          returnValue: any
        ) => {
          if (dialogResult === SP.UI.DialogResult.OK) {
            // Set status
            SP.SOD.executeFunc(
              SPScripts.SP.Script,
              SPScripts.SP.UI.Notify,
              () => {
                self.notify = SP.UI.Notify.addNotification(
                  "Выполняем удаление файла",
                  true
                );
              }
            );
            // Delete File
            const fileName = self.file.fileName;
            self
              .deleteFile(self.file)
              .then(a => {
                // Set Status
                SP.SOD.executeFunc(
                  SPScripts.SP.Script,
                  SPScripts.SP.UI.Status,
                  () => {
                    SP.UI.Notify.removeNotification(self.notify);
                    SP.UI.Status.removeAllStatus(false);
                    const status = SP.UI.Status.addStatus(
                      "Файл " + fileName + " успешно удален.",
                      ""
                    );
                    SP.UI.Status.setStatusPriColor(status, "green");
                  }
                );
              })
              .catch((e: SP.ClientRequestFailedEventArgs) => {
                // Set Status
                SP.SOD.executeFunc(
                  SPScripts.SP.Script,
                  SPScripts.SP.UI.Status,
                  () => {
                    SP.UI.Notify.removeNotification(self.notify);
                    SP.UI.Status.removeAllStatus(false);
                    const status = SP.UI.Status.addStatus(
                      "Ошибка при удалении файла " + fileName + ".",
                      e.get_errorDetails() + ""
                    );
                    SP.UI.Status.setStatusPriColor(status, "red");
                  }
                );
              });
          }
        }
      });
    } else if (arg === SaveEditButtonKey.Save) {
      const self = this;
      // Notify
      SP.SOD.executeFunc(SPScripts.SP.Script, SPScripts.SP.UI.Notify, () => {
        self.notify = SP.UI.Notify.addNotification(
          "Выполняем сохранение файла",
          true
        );
      });
      // Change File Action
      this.changeFile({
        fileName: self.file.fileName,
        docType: self.docType,
        info: self.addInfo
      })
        .then(changeArgs => {
          // Save File Action
          self
            .saveFile({
              siteUrl: self.$store.state.options.siteUrl,
              listId: self.$store.state.options.listId,
              checkinType: SP.CheckinType.majorCheckIn,
              file: self.file
            })
            .then(args => {
              self.saved = true;
              // Set Status
              SP.SOD.executeFunc(
                SPScripts.SP.Script,
                SPScripts.SP.UI.Status,
                () => {
                  SP.UI.Notify.removeNotification(self.notify);
                  SP.UI.Status.removeAllStatus(false);
                  const status = SP.UI.Status.addStatus(
                    "Файл " +
                      self.file.fileName +
                      " успешно сохранен в системе",
                    ""
                  );
                  SP.UI.Status.setStatusPriColor(status, "green");
                }
              );
            })
            .catch(e => {
              SP.SOD.executeFunc(
                SPScripts.SP.Script,
                SPScripts.SP.UI.Status,
                () => {
                  SP.UI.Notify.removeNotification(self.notify);
                  SP.UI.Status.removeAllStatus(false);
                  const status = SP.UI.Status.addStatus(
                    "Ошибка при сохранении файла " + self.file.fileName + ".",
                    e + ""
                  );
                  SP.UI.Status.setStatusPriColor(status, "red");
                }
              );
            });
        })
        .catch(e => {
          SP.SOD.executeFunc(
            SPScripts.SP.Script,
            SPScripts.SP.UI.Status,
            () => {
              SP.UI.Notify.removeNotification(self.notify);
              SP.UI.Status.removeAllStatus(false);
              const status = SP.UI.Status.addStatus(
                "Ошибка при сохранении файла " + self.file.fileName + ".",
                e + ""
              );
              SP.UI.Status.setStatusPriColor(status, "red");
            }
          );
        });
    }
  }

  public onchange(event: any) {
    // or use $store.dispatch(methodName, payload);
    //if (this.$store.state.mode === UploadMode.New) {
    this.changeFile({
      fileName: this.file.fileName,
      docType: this.docType,
      info: this.addInfo
    });
    //}
  }
}
</script>
<style scoped>
.db-select {
  width: 250px;
  max-width: 300px;
  font-size: medium;
}
.db-input {
  width: 310px;
  max-width: 310px;
}
.ms-formbody {
  padding-left: 6px;
}
.db-iconCell {
  width: 24px;
}
.ms-cellstyle {
  text-align: left;
}
</style>
