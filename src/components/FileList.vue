<template>
  <div class="db-container">
    <table class="ms-listviewtable">
      <tr>
        <th class="ms-vh2" style="max-width: 24px;">
          <div class="ms-vh-div"></div>
        </th>
        <th class="ms-vh2" style="max-width: 500px;">
          <div class="ms-vh-div">
            <span>Название файла</span>
          </div>
        </th>

        <th class="ms-vh2" style="max-width: 500px;">
          <div class="ms-vh-div">
            <span>Дата</span>
          </div>
        </th>

        <th class="ms-vh2" style="max-width: 500px;">
          <div class="ms-vh-div">
            <span>Размер</span>
          </div>
        </th>

        <th class="ms-vh2" style="max-width: 500px;">
          <div class="ms-vh-div">
            <span>Версия</span>
          </div>
        </th>

        <th class="ms-vh2" style="max-width: 500px;">
          <div class="ms-vh-div">
            <span>Тип документа</span>
          </div>
        </th>

        <th class="ms-vh2" style="max-width: 500px;">
          <div class="ms-vh-div">
            <span>Примечание</span>
          </div>
        </th>
        <th class="ms-vh2" style="max-width: 500px;" v-if="mode === 2">
          <div class="ms-vh-div">
            <span></span>
          </div>
        </th>
      </tr>

      <File v-for="item in files" :key="item.fileName" v-bind:file="item"></File>

      <tr v-if="!files || files.length === 0">
        <td
          colspan="8"
          class="ms-list-emptyText-compact ms-textLarge ms-alignLeft db-empty"
        >
          <span>Нет файлов для отображения в таблице</span>
        </td>
      </tr>
    </table>
  </div>
</template>
<script lang="ts">
import utils from "../utils";
import { AppOptions, StateFile } from "../types";
import { Component, Prop, Vue, Watch } from "vue-property-decorator";
import { mapState, mapActions, ActionMethod } from "vuex";
import actionTypes from "../store/action-types";
import store from "@/store/store";
import File from "@/components/File.vue";
import { UploadMode } from "@/types";

@Component({
  components: {
    File
  },
  computed: {
    ...mapState({
      files: files => store.state.files,
      mode: mode => store.state.mode
    })
  }
})
export default class FileList extends Vue {
  public files!: StateFile[];
  public mode!: UploadMode;
  public onsaveclick() {}
}
</script>
<style>
.db-container {
  margin-top: 15px;
  margin-bottom: 15px;
  border: 1px #eeeeee solid;
}
.db-empty {
  padding: 10px;
}
.ms-tableCell {
  padding: 5px;
}
</style>
