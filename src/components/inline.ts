import Vue from 'vue'
import OnSaveDialog from './OnSaveDialog.vue';
import { DialogText } from '@/types';

export function createOnSaveDialogElement(elementId: string, dialogText: DialogText, files: string[]) {
    let element = document.getElementById(elementId);
    if (!element) {
        element = document.createElement('div');
        element.id = elementId;
    }
    const DialogComponent = Vue.extend({
        components: {
            OnSaveDialog,
        },
        render: (h, context) => {
            return h(OnSaveDialog, {
                props: {
                    unsavedFiles: files,
                    dialogText,
                },
                scopedSlots: {
                    ['instructions']: (props) => props.dialogText,
                    ['buttonOk']: (props) => props.dialogText,
                    ['subject']: (props) => props.dialogText,
                },
            });
        }
    });
    var component = new DialogComponent().$mount(element);
    element.appendChild(component.$el);
    return element;
}