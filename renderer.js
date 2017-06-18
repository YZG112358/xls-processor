// This file is required by the index.html file and will
// be executed in the renderer process for that window.
// All of the Node.js APIs are available in this process.

const ipc = require('electron').ipcRenderer;
const xlsx = require('xlsx');

document.getElementById('select-directory').addEventListener('click', function (event) {
  ipc.send('open-file-dialog');
});

document.getElementById('process').addEventListener('click', function (event) {
  ipc.send('process');
});

ipc.on('message', function (e, message) {
  var dom = document.getElementById('message');
  dom.innerHTML = '<p>' + message + '</p>' + dom.innerHTML;
});

document.getElementById('save-btn').addEventListener('click', function (event) {
  ipc.send('save-dialog')
})
