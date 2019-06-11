<template>
  <div>
    <span v-if="!active">
      <a class="ms-draggable" @click="onclick(2)" target="_self">
        <img alt="edit" :src="editImgUrl" border="0">
      </a>
      &nbsp;&nbsp;
      <a class="ms-draggable" @click="onclick(3)" target="_self">
        <img alt="delete" :src="deleteImgUrl" border="0">
      </a>
    </span>
    <span v-if="active">
      <a class="ms-draggable" @click="onclick(1)" target="_self">
        <img alt="save" :src="saveImgUrl" border="0">
      </a>
      &nbsp;&nbsp;
      <a class="ms-draggable" @click="onclick(4)" target="_self">
        <img alt="cancel" :src="cancelImgUrl" border="0">
      </a>
    </span>
  </div>
</template>
<script lang="ts">
import utils, { SPDataService } from "../utils";
import { AppOptions, SaveEditButtonKey } from "../types";
import { Component, Prop, Vue, Watch, Model } from "vue-property-decorator";
import { mapState, mapActions, ActionMethod } from "vuex";
import actions from "../store/action-types";
import store from "@/store/store";
// Save = 1,
//     Edit = 2,
//     Remove = 3,
//     Cancel = 4,
@Component({})
export default class SaveEditButton extends Vue {
  @Prop() isActive!: boolean;
  public active: boolean = false;

  public cancelImgUrl: string = `${
    _spPageContextInfo.webAbsoluteUrl
  }/_layouts/15/images/CancelGlyph.16x16x32.png?rev=23`;

  public editImgUrl: string = `${
    _spPageContextInfo.webAbsoluteUrl
  }/_layouts/15/images/edititem.gif?rev=23`;

  public deleteImgUrl: string = `${
    _spPageContextInfo.webAbsoluteUrl
  }/_layouts/15/images/delitem.gif?rev=23`;

  public saveImgUrl: string = `${
    _spPageContextInfo.webAbsoluteUrl
  }/_layouts/15/images/saveitem.gif?rev=23`;

  onclick(command: SaveEditButtonKey) {
    // if command remove do nothing.
    console.log(command);
    if (command !== SaveEditButtonKey.Remove) {
      this.active = !this.active;
    }
    this.$emit("active-changed", command);
  }
}
</script>
<style scoped>
</style>
