const { app, BrowserWindow } = require('electron')
const XLSX = require('xlsx')

function createWindow() {
    const win = new BrowserWindow({
        width: 634,
        height: 394,
        resizable: false,
        transparent: false,
        frame: true,
        show: false,
        webPreferences: {
            nodeIntegration: true
        }
    })
    win.loadFile('./src/index.html')

    win.once("ready-to-show", () => {
        win.show()
    })
}


function convertToJSON() {
    var selectedFile = evt.target.files[0];
    var reader = new FileReader();
    reader.onload = function (event) {
        var data = event.target.result;
        var workbook = XLSX.read(data, {
            type: 'binary'
        });
        workbook.SheetNames.forEach(function (sheetName) {

            var XL_row_object = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
            var json_object = JSON.stringify(XL_row_object);
            document.getElementById("jsonObject").innerHTML = json_object;

        })
    };

    reader.onerror = function (event) {
        console.error("File could not be read! Code " + event.target.error.code);
    };

    reader.readAsBinaryString(selectedFile);

}

app.whenReady().then(createWindow)

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit()
    }
})

app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
        createWindow()
    }
})