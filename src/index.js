"use strict"

const EXCEL_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8'
const EXCEL_EXTENSION = '.xlsx'

const { app, BrowserWindow } = require('electron')
const XLSX = require('xlsx')

let mainWindow
let productsChecked = []
var parameters = []

//------------------------------------------------------------------------------------------------------------------------------------------------
// Create main window
//------------------------------------------------------------------------------------------------------------------------------------------------

function createWindow() {
    mainWindow = new BrowserWindow({
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

    mainWindow.setMenu(null)

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

app.on('browser-window-created', function (e, window) {
    window.setMenu(null)
})

//------------------------------------------------------------------------------------------------------------------------------------------------
// Load data from excel tables and compare 
//------------------------------------------------------------------------------------------------------------------------------------------------

function loadDataFromExcel() {
    productsChecked = []
    showProgressBarDiv()
    Promise.all([getProductsFromExcel(), getProductsToFixFromExcel()]).then((values) => {
        const productsFromDatabase = values[0]
        const productsToFixData = values[1]
        compareDataExtractedFromExcel(productsFromDatabase, productsToFixData)
    })
}

function compareDataExtractedFromExcel(productsFromDatabase, productsToFixData) {
    productsToFixData.forEach(function (productsFound, index) {
        setTimeout(function () {
            let value = ((index + 1) * 100) / productsToFixData.length
            updateProgressStatus(value)

            const productsByBarCode = productsFromDatabase.filter(item => {
                return item.REFERENCIA === productsFound.REFERENCIA
            })
            const productToFix = productsByBarCode[0]
            const fixedProduct = fixProduct(productToFix)
            addUpdatedProductToJSON(fixedProduct)
            if (productsChecked.length === productsToFixData.length) {
                downloadAsExcel(productsChecked)
            }
        }, 1)
    })
}

//------------------------------------------------------------------------------------------------------------------------------------------------
// Find products from excel file
//------------------------------------------------------------------------------------------------------------------------------------------------

function getProductsFromExcel() {
    return new Promise((resolve, rejected) => {
        const selectedFile = document.getElementById('file').files[0]
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
            hideProgressBarDiv()
            alert('Não foi possível carregar os dados da planilha mãe!')
            rejected({
                name: 'Erro',
                message: 'Não foi possível carregar os itens.'
            })
        }
    })
}

//------------------------------------------------------------------------------------------------------------------------------------------------
// Find products to fix from excel file
//------------------------------------------------------------------------------------------------------------------------------------------------

function getProductsToFixFromExcel() {
    return new Promise((resolve, rejected) => {
        const selectedFile = document.getElementById('file2').files[0]
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
            hideProgressBarDiv()
            alert('Não foi possível carregar os dados da planilha a ser analisada!')
            rejected({
                name: 'Erro',
                message: 'Não foi possível carregar os itens.'
            })
        }
    })
}

//------------------------------------------------------------------------------------------------------------------------------------------------
// Update input file name on html
//------------------------------------------------------------------------------------------------------------------------------------------------

function updateInputProductsDataSpan(element) {
    hideProgressBarDiv()
    const inputProductsDataSpan = document.getElementById('inputProductsDataSpan')
    inputProductsDataSpan.textContent = element.value.split('\\').pop()
    if (element.value === "") {
        inputProductsDataSpan.textContent = "Insira a planilha mãe aqui..."
    }
    parameters = []
}

function updateInputProductsToFixDataSpan(element) {
    hideProgressBarDiv()
    const inputProductsToFixDataSpan = document.getElementById('inputProductsToFixDataSpan')
    inputProductsToFixDataSpan.textContent = element.value.split('\\').pop()
    if (element.value === "") {
        inputProductsToFixDataSpan.textContent = "Insira a planilha a ser analisada aqui..."
    }
}

//------------------------------------------------------------------------------------------------------------------------------------------------
// Fix product data 
//------------------------------------------------------------------------------------------------------------------------------------------------

function fixProduct(productToFix) {
    const newItem = []
    parameters.forEach(function (currentParameter, index) {
        const data = {
            [currentParameter]: productToFix[currentParameter]
        }
        newItem.push(data)
    })
    const productFixed = {}
    for (var i = 0; i < newItem.length; i++) {
        for (var propriedade in newItem[i]) {
            productFixed[propriedade] = newItem[i][propriedade]
        }
    }
    return productFixed
}

function addUpdatedProductToJSON(productFixed) {
    productsChecked.push(productFixed)
}

//------------------------------------------------------------------------------------------------------------------------------------------------
// Download new excel table
//------------------------------------------------------------------------------------------------------------------------------------------------

function downloadAsExcel(data) {
    if (data) {
        var worksheet = XLSX.utils.json_to_sheet(data)
        const workbook = {
            Sheets: {
                'data': worksheet
            },
            SheetNames: ['data']
        }
        const today = new Date()
        const fileName = 'Planilha atualizada: ' + today.toLocaleDateString("pt-Br")
        const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' })
        saveAsExcel(excelBuffer, fileName)
    }
}

function saveAsExcel(buffer, filename) {
    const data = new Blob([buffer], { type: EXCEL_TYPE })
    saveAs(data, filename + EXCEL_EXTENSION)
}

//------------------------------------------------------------------------------------------------------------------------------------------------
// Update progress status
//------------------------------------------------------------------------------------------------------------------------------------------------

function showProgressBarDiv() {
    const progressBarDiv = document.getElementById('progressBarDiv')
    progressBarDiv.style.display = "grid"
    updateProgressStatus(0)
}

function hideProgressBarDiv() {
    const progressBarDiv = document.getElementById('progressBarDiv')
    progressBarDiv.style.display = "none"
}

function updateProgressStatus(value) {
    var bar = document.getElementById('bar')
    var barLabel = document.getElementById('barLabel')
    const newValue = Number(value).toPrecision(3) | 0
    bar.style.width = newValue + '%'
    barLabel.textContent = newValue + '%'
}

//------------------------------------------------------------------------------------------------------------------------------------------------
// Create modal window to save settings
//------------------------------------------------------------------------------------------------------------------------------------------------

function createSettingsWindow() {
    const selectedFile = document.getElementById('file').files[0]
    if (selectedFile) {
        fetchParameters()
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
// Fetch parameters from excel file
//------------------------------------------------------------------------------------------------------------------------------------------------

function fetchParameters() {
    const checkBoxMainDiv = document.getElementById('checkBoxDiv')
    checkBoxMainDiv.innerHTML = ''
    createLoaderDiv()
    if (parameters !== undefined && parameters.length > 0) {
        loadLocalParameters()
    } else {
        loadParametersFromExcel()
    }
}

function loadLocalParameters() {
    parameters.forEach(function (parameter, index) {
        createParameterCheckbox(parameter)
        if (index == parameters.length - 1) {
            showParameters()
        }
    })
}

function loadParametersFromExcel() {
    getProductsFromExcel().then(function (products) {
        removeLoaderDiv()
        const headerItems = Object.keys(products[0])
        parameters = headerItems
        headerItems.forEach(function (parameter, index) {
            createParameterCheckbox(parameter)
            if (index == headerItems.length - 1) {
                showParameters()
            }
        })
    })
}

//------------------------------------------------------------------------------------------------------------------------------------------------
// Create parameters checkbox in html
//------------------------------------------------------------------------------------------------------------------------------------------------

function createParameterCheckbox(parameter) {
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
    const loaderDiv = document.getElementById('loader')
    loaderDiv.style.display = "none"
}

function showParameters() {
    removeLoaderDiv()
    document.getElementById("checkBoxDiv").style.display = "grid"
}