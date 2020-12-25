const { app, BrowserWindow } = require('electron')
const XLSX = require('xlsx')
let correctedItems

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

//------------------------------------------------------------------------------------------------------------------------------------------------
// 
//------------------------------------------------------------------------------------------------------------------------------------------------

function getItems() {
    return new Promise((resolve, rejected) => {
        const selectedFile = document.getElementById('inputRightData').files[0]
        const progressBar = document.getElementById('progressBar')
        progressBar.style.width = "30%"
        if (selectedFile) {
            let fileReader = new FileReader()
            fileReader.readAsBinaryString(selectedFile)
            fileReader.onload = (event) => {
                let data = event.target.result
                let workbook = XLSX.read(data, { type: "binary" })
                workbook.SheetNames.forEach(function (sheet, i) {
                    let rowObject = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheet])
                    resolve(rowObject)
                })
            }
        } else {
            alert('Não foi possível carregar os dados da planilha mãe!')
            rejected({
                name: 'Erro',
                message: 'Não foi possível carregar os itens.'
            })
        }
    })
}

function getItemsToAnalyze() {
    return new Promise((resolve, rejected) => {
        const selectedFile = document.getElementById('inputWrongData').files[0]
        if (selectedFile) {
            let fileReader = new FileReader()
            fileReader.readAsBinaryString(selectedFile)
            fileReader.onload = (event) => {
                let data = event.target.result
                let workbook = XLSX.read(data, { type: "binary" })
                workbook.SheetNames.forEach(function (sheet, i) {
                    let rowObject = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheet])
                    resolve(rowObject)
                })
            }
        } else {
            alert('Não foi possível carregar os dados da planilha a ser analisada!')
            rejected({
                name: 'Erro',
                message: 'Não foi possível carregar os itens.'
            })
        }
    })
}

function compareData() {
    getItems().then(function (items) {
        console.log('Loaded itens count: ' + items.length)
        getItemsToAnalyze().then(function (itemsToAnalyze) {
            console.log('Loaded itens to analyze count: ' + itemsToAnalyze.length)
            itemsToAnalyze.forEach(function (item, index) {
                if (index < 1) {
                    const barCode = item.REFERENCIA
                    console.log(barCode)
                    const testeItem = items.filter(item => {
                        return item.REFERENCIA === barCode
                    })
                    console.log(testeItem[0])
                    console.log(Object.keys(testeItem[0]))
                }
            })
        })
    })
}
