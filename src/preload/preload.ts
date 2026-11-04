import { contextBridge, ipcRenderer } from 'electron';
import { createReporterApi } from './api';

const reporterApi = createReporterApi(ipcRenderer);

contextBridge.exposeInMainWorld('reporterApi', reporterApi);
