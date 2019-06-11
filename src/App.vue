<template>
  <div id="app">
    <!-- <img alt="Vue logo" src="./assets/logo.png"> -->
    <!-- <HelloWorld msg="Welcome to Your Vue.js + TypeScript App"/> -->
    <UploadFiles v-bind:config='config'></UploadFiles>
    <FileList></FileList>
  </div>
</template>

<script lang="ts">
import { Component, Vue, Prop } from 'vue-property-decorator';
import { AppOptions } from './types';
import UploadFiles from './components/UploadFiles.vue';
import FileList from './components/FileList.vue';

@Component({
  components: {
    UploadFiles,
  }
})
export default class App extends Vue {
  @Prop() config!:AppOptions;
  mounted() {
  }
}
</script>

<style lang="scss">
#app {
  font-family: 'Avenir', Helvetica, Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  text-align: center;
  color: #2c3e50;
}
</style>
