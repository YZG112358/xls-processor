const electron = require('electron')
const ipc = electron.ipcMain
const dialog = electron.dialog
// Module to control application life.
const app = electron.app
// Module to create native browser window.
const BrowserWindow = electron.BrowserWindow

const path = require('path')
const url = require('url')

const xlsx = require('xlsx');

const _ = require('lodash');

// Keep a global reference of the window object, if you don't, the window will
// be closed automatically when the JavaScript object is garbage collected.
let mainWindow

function createWindow () {
  // Create the browser window.
  mainWindow = new BrowserWindow({width: 800, height: 300})

  // and load the index.html of the app.
  mainWindow.loadURL(url.format({
    pathname: path.join(__dirname, 'index.html'),
    protocol: 'file:',
    slashes: true
  }))

  // Open the DevTools.
  // mainWindow.webContents.openDevTools()

  // Emitted when the window is closed.
  mainWindow.on('closed', function () {
    // Dereference the window object, usually you would store windows
    // in an array if your app supports multi windows, this is the time
    // when you should delete the corresponding element.
    mainWindow = null
  })
}

// This method will be called when Electron has finished
// initialization and is ready to create browser windows.
// Some APIs can only be used after this event occurs.
app.on('ready', createWindow)

// Quit when all windows are closed.
app.on('window-all-closed', function () {
  // On OS X it is common for applications and their menu bar
  // to stay active until the user quits explicitly with Cmd + Q
  if (process.platform !== 'darwin') {
    app.quit()
  }
})

app.on('activate', function () {
  // On OS X it's common to re-create a window in the app when the
  // dock icon is clicked and there are no other windows open.
  if (mainWindow === null) {
    createWindow()
  }
})

var _path = null;
var _workbook = null;

var log = function (message) {
  mainWindow.webContents.send('message', message);
};

ipc.on('open-file-dialog', function (event) {
  dialog.showOpenDialog({
    properties: ['openFile']
  }, function (files) {
    if (files) {
      try {
        _path = files;
        _workbook = xlsx.readFile(String(_path), {
          // cellDates: true
        });
        log('Opened ' + _path);
      } catch (e) {
        log('Target is not a standard xls file');
      }
    }
  })
})

var getKeywords = function (sheet, column) {
  var keywords = [];
  do {
    var c = column + (keywords.length + 1);
    var v = sheet[c] ? sheet[c].v : null;
    if (!v) return keywords;
    keywords.push(v);
  } while (true);
};

var dateFormat = function (sheet, length, column) {
  Array
    .apply(null, { length })
    .map(Number.call, Number)
    .map((n) => {
      _.each(
        _.flattenDeep(column),
        (c) => {
          sheet[c + (n + 1)].z = 'yyyy/mm/dd';
        }
      )
    });
}

ipc.on('process', function (event) {
  log('Reading ' + _path);

  // Reading data
  var sheet = _workbook.Sheets[_workbook.SheetNames[0]];
  var maxRange = xlsx.utils.decode_range(sheet['!ref']).e.r;
  var c = _.filter(
    Array
      .apply(null, { length: 26 })
      .map(Number.call, Number)
      .map((n) => { return String.fromCharCode(65 + n); }),
    (c) => !!sheet[c + 1]
  );
  var key = c.map((c) => sheet[c + 1].v || '');
  var values = [];
  do {
    let r = values.length + 2;
    let value = c.map((c) => sheet[c + r]
      ? sheet[c + r].t === 'd'
        // ? new Date(sheet[c + r].v)
        ? sheet[c + r].v
        : sheet[c + r].v
      : ''
    );
    if (!_.compact(value).length) break;
    values.push(_.zipObject(key, value));
  } while (true);
  log('Parsed ' + values.length + ' line of data');

  // Get keywords
  var incKeywords = getKeywords(_workbook.Sheets[_workbook.SheetNames[1]], 'A');
  var decKeywords = getKeywords(_workbook.Sheets[_workbook.SheetNames[1]], 'B');
  log('Include keywords ' + incKeywords.join(', '));
  log('Decclude keywords ' + decKeywords.join(', '));

  log(JSON.stringify(values[0]));

  // Filter
  _workbook.SheetNames[2] = 'A';
  _workbook.SheetNames[3] = 'B';
  _workbook.SheetNames[4] = 'C';

  var _filter = (value) => {
    return incKeywords.some((keyword) => value['检查结论'].indexOf(keyword) > -1)
      && decKeywords.every((keyword) => value['检查结论'].indexOf(keyword) === -1);
  };

  var valuesA = values.filter(_filter);
  var valuesB = values.filter((v) => !_filter(v));
  var valuesC = _.sortBy(
    _.filter(
      _.map(
        _.sortedUniqBy(
          _.sortBy(
            valuesA,
            (value) => -value['检查时间']
          ),
          (value) => value['姓名']
        ),
        (value) => _.assign(value, _.mapKeys(_.pick(_.findLast(valuesB, (valueB) => 
          value['姓名'] === valueB['姓名']
            && value['检查时间'] >= valueB['检查时间']
            && value['检查时间'] <= valueB['检查时间'] + 30
        ) || {}, ['检查时间', '检查结论']), (value, key) => '上次' + key))
      ),
      (value) => _.has(value, '上次检查时间')
    ),
    (value) => value['检查时间']
  );
  console.log(valuesC);

  _workbook.Sheets['A'] = xlsx.utils.json_to_sheet(valuesA);
  _workbook.Sheets['B'] = xlsx.utils.json_to_sheet(valuesB);
  _workbook.Sheets['C'] = xlsx.utils.json_to_sheet(valuesC);

  dateFormat(_workbook.Sheets['A'], valuesA.length, 'C');
  dateFormat(_workbook.Sheets['B'], valuesB.length, 'C');
  dateFormat(_workbook.Sheets['C'], valuesC.length, ['C', 'G']);
  dateFormat(sheet, maxRange);

  log('Done');
});

ipc.on('save-dialog', function (event) {
  const options = {
    title: 'Save',
    filters: [
      { name: 'Excel', extensions: ['xlsx'] }
    ]
  }
  dialog.showSaveDialog(options, function (filename) {
    xlsx.writeFile(_workbook, filename, {
      // cellDates: true
    });
    log('File saved');
  })
})
