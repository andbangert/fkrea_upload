import { GetterTree } from 'vuex';
import {
    StateObject,
} from '../types';

export const getters: GetterTree<StateObject, StateObject> = {
    getUnsavedFiles: (state) => () => {
        return state.files.filter((f) => !f.saved);
    },
    changedFiles: (state) => () => {
        return state.files.filter((f) => f.changed);
    },
    hasUnsavedFiles: (state, getrs) => () => {
        const files = getrs.unsavedFiles; // as Array<StateFile>;
        return files && files.length > 0;
    },
};
