<template>
  <div>
    <div class="ms-toolbar ms-alignLeft">
      <div class="ms-list-addnew ms-list-addnew-aligntop ms-textXLarge ">
        <!-- <slot name="opName">
          Следующие файлы были изменены:
        </slot> -->
        <slot v-bind:subject="dialogText">
          {{ dialogText.subject }}
        </slot>
      </div>
      <div class="ms-list-addnew ms-list-addnew-aligntop ms-textXLarge ms-alignLeft">
        <ul>
          <li v-for="item in unsavedFiles" :key="item" class="ms-textSmall">{{item}}</li>
        </ul>
      </div>
      <div class="ms-list-addnew ms-list-addnew-aligntop ms-textXLarge ">
        <!-- <slot name="instruction">
        Нажмите кнопку "сохранить", чтобы сохранить изменения или отмена, чтобы выйти без сохранения.
        </slot> -->
        <slot v-bind:instructions="dialogText">
          {{ dialogText.instructions }}
        </slot>
      </div>
    </div>
    <div class="ms-toolbar ms-alignLeft">
      <button type="button" @click="onclick(1)" id="btnSave" ref="btnSave">
        <slot v-bind:buttonOk="dialogText">
          {{ dialogText.buttonOk }}
        </slot>
        </button>&nbsp;
      <button type="button" @click="onclick(0)">Отмена</button>
    </div>
  </div>
</template>
<script lang="ts">
import { Component, Prop, Vue, Watch } from "vue-property-decorator";
import { SPScripts } from "@/utils/constants";
import { DialogText } from '@/types';

@Component
export default class OnSaveDialog extends Vue {
  @Prop() unsavedFiles!: string[];
  @Prop() dialogText!: DialogText;
mounted() {
    const btn = this.$refs["btnSave"] as HTMLButtonElement; // .focus();
    if (btn) {
      btn.focus();
    }
  }
  onclick(result: SP.UI.DialogResult) {
    SP.SOD.executeFunc(
      SPScripts.SP.UI.Dialog.Script,
      SPScripts.SP.UI.Dialog.ShowModalDialog,
      () => {
        SP.UI.ModalDialog.commonModalDialogClose(result, "");
      }
    );
  }
}
</script>
