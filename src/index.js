"use strict";
const EXCEL_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';

const { app, BrowserWindow } = require('electron')
const XLSX = require('xlsx');
//const saveAs = require('./filesaver');

var parameters

function createWindow() {
    const mainWindow = new BrowserWindow({
        width: 634,
        height: 394,
        resizable: false,
        transparent: false,
        frame: true,
        show: false,
        webPreferences: {
            devTools: true,
            nodeIntegrationInWorker: true,
            nodeIntegration: true
        }
    })
    mainWindow.loadFile('./src/index.html')

    mainWindow.once("ready-to-show", () => {
        mainWindow.show()
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
// Create modal window to save settings
//------------------------------------------------------------------------------------------------------------------------------------------------

function createSettingsWindow() {
    const selectedFile = document.getElementById('inputRightData').files[0]
    if (selectedFile) {
        createColumnsParameters()
        var settingsModal = document.getElementById("settingsModal")
        settingsModal.style.display = "block"

        window.onclick = function (event) {
            if (event.target == settingsModal) {
                settingsModal.style.display = "none"
            }
        }
    } else {
        alert('Insira uma planilha mãe para acessar os parâmetros!')
    }
}

function closeModal() {
    var settingsModal = document.getElementById("settingsModal")
    settingsModal.style.display = "none"
}

//------------------------------------------------------------------------------------------------------------------------------------------------
// Compare data and create new excel file
//------------------------------------------------------------------------------------------------------------------------------------------------

function createColumnsParameters() {
    const checkBoxMainDiv = document.getElementById('checkBoxDiv')
    checkBoxMainDiv.innerHTML = ''
    createLoaderDiv()
    getItems().then(function (items) {
        removeLoaderDiv()
        const headerItems = Object.keys(items[0])
        parameters = headerItems
        headerItems.forEach(function (parameter, index) {
            createParameter(parameter)
            if (index == headerItems.length - 1) {
                showParameters()
            }
        })
    })
}

//------------------------------------------------------------------------------------------------------------------------------------------------
// Create parameters to update
//------------------------------------------------------------------------------------------------------------------------------------------------

function createParameter(parameter) {
    const checkBoxMainDiv = document.getElementById('checkBoxDiv')
    const checkBoxDiv = document.createElement('div')
    checkBoxDiv.setAttribute('id', 'divBox')
    checkBoxDiv.setAttribute('class', 'divBox')

    const checkBox = document.createElement('input')
    checkBox.setAttribute('type', 'checkbox')
    checkBox.setAttribute('id', parameter)
    checkBox.setAttribute('class', 'checkBox')
    checkBox.setAttribute('checked', true)

    const labelCheckBox = document.createElement('label')
    labelCheckBox.setAttribute('for', parameter)
    labelCheckBox.setAttribute('id', 'labelCheckBox')
    labelCheckBox.setAttribute('class', 'labelCheckBox')

    labelCheckBox.innerHTML = parameter
    checkBox.value = true
    checkBoxMainDiv.appendChild(checkBoxDiv)
    checkBoxDiv.appendChild(checkBox)
    checkBoxDiv.appendChild(labelCheckBox)
}

//------------------------------------------------------------------------------------------------------------------------------------------------
// Create animation
//------------------------------------------------------------------------------------------------------------------------------------------------

function createLoaderDiv() {
    const loaderDiv = document.createElement('div')
    loaderDiv.setAttribute('class', 'loader')
    loaderDiv.setAttribute('id', 'loader')
    const checkBoxDiv = document.getElementById('checkBoxDiv')
    checkBoxDiv.appendChild(loaderDiv)
}

function removeLoaderDiv() {
    const checkBoxDiv = document.getElementById('checkBoxDiv')
    const loaderDiv = document.getElementById('loader')
    checkBoxDiv.removeChild(loaderDiv)
}

function showParameters() {
    removeLoaderDiv()
    document.getElementById("checkBoxDiv").style.display = "block"
}

//------------------------------------------------------------------------------------------------------------------------------------------------
// Compare data and create new excel file
//------------------------------------------------------------------------------------------------------------------------------------------------

function getItems() {
    showProgressBar()
    return new Promise((resolve, rejected) => {
        const selectedFile = document.getElementById('inputRightData').files[0]
        if (selectedFile) {
            let fileReader = new FileReader()
            fileReader.readAsBinaryString(selectedFile)
            fileReader.onload = (event) => {
                let data = event.target.result
                let workbook = XLSX.read(data, { type: "binary" })
                workbook.SheetNames.forEach(function (sheet, i) {
                    let rowObject = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheet])
                    const headerItems = Object.keys(rowObject[0])
                    parameters = headerItems
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
    getItems().then(function (checkedItems) {
        console.log(parameters)
        const checkedItemsCount = checkedItems.length
        console.log('Loaded items count: ' + checkedItemsCount)
        getItemsToAnalyze().then(function (itemsToAnalyze) {
            const itemsToAnalyzeCount = itemsToAnalyze.length
            console.log('Loaded items to analyze count: ' + itemsToAnalyzeCount)
            console.log(itemsToAnalyze)
            let productsChecked = []
            itemsToAnalyze.forEach(function (itemToAnalyze, index) {
                getProgressPercentage(index, itemsToAnalyzeCount - 1)
                const itemsToUpdate = checkedItems.filter(item => {
                    return item.REFERENCIA === itemToAnalyze.REFERENCIA
                })
                const itemToUpdate = itemsToUpdate[0]
                const newItem = []
                parameters.forEach(function (currentParameter, index) {
                    const data = { [currentParameter]: itemToUpdate[currentParameter] }
                    newItem.push(data)
                })
                const newItemJSON = {}
                for (var i = 0; i < newItem.length; i++) {
                    for (var propriedade in newItem[i]) {
                        newItemJSON[propriedade] = newItem[i][propriedade]
                    }
                }
                productsChecked.push(newItemJSON)
                console.log(newItemJSON)
                const products = JSON.stringify(productsChecked)
                console.log('Produtos: ' + products)
            })
            downloadAsExcel(productsChecked)
        })
    })
}

function downloadAsExcel(data) {
    var worksheet = XLSX.utils.json_to_sheet(data)
    const workbook = {
        Sheets: {
            'data': worksheet
        },
        SheetNames: ['data']
    }

    const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' })
    console.log(excelBuffer)
    saveAsExcel(excelBuffer, 'myFile')
}

function saveAsExcel(buffer, filename) {
    const data = new Blob([buffer], { type: EXCEL_TYPE })
    saveAs(data, filename + EXCEL_EXTENSION)
}

function getProgressPercentage(value, total) {
    const percent = (value / total) * 100
    updateProgressTo(percent.toPrecision(3))
}

function showProgressBar() {
    const progressBarDiv = document.getElementById('progressBarDiv')
    progressBarDiv.style.display = "block"
}

function updateProgressTo(progress) {
    const progressBar = document.getElementById('progressBar')
    progressBar.style.width = progress + "%"
    progressBar.innerHTML = progress + "%"
}