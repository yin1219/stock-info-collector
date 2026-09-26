import { contextBridge, ipcRenderer } from 'electron';
import { createReporterApi } from './api';

const reporterApi = createReporterApi(process.versions.electron, ipcRenderer);

contextBridge.exposeInMainWorld('reporterApi', reporterApi);
