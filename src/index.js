"use strict"

const EXCEL_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8'
const EXCEL_EXTENSION = '.xlsx'

const { app, BrowserWindow } = require('electron')
const XLSX = require('xlsx')

let mainWindow
let newProducts = []
let localParameters = {}

//------------------------------------------------------------------------------------------------------------------------------------------------
// Create main window
//------------------------------------------------------------------------------------------------------------------------------------------------

function createWindow() {
    mainWindow = new BrowserWindow({
        icon: '../icon/icon.ico',
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
    newProducts = []
    showProgressBarDiv()
    Promise.all([getProductsFromExcel(), getProductsToFixFromExcel()]).then((values) => {
        const productsFromDatabase = values[0]
        const productsToFixData = values[1]
        compareDataExtractedFromExcel(productsFromDatabase, productsToFixData)
    })
}

function compareDataExtractedFromExcel(productsFromDatabase, productsToFixData) {
    if (checkFieldsNames(productsFromDatabase[0], "mãe.") && checkFieldsNames(productsToFixData[0], "filha.")) {
        productsToFixData.forEach(function (productToFix, index) {
            setTimeout(function () {
                let value = ((index + 1) * 100) / productsToFixData.length
                updateProgressStatus(value)

                if (productToFix.CODIGO_DE_BARRAS) {
                    console.log('Produto tem codigo de barras: ' + productToFix.CODIGO_DE_BARRAS)
                    const productFoundByBarCode = productsFromDatabase.find(productFromDatabase => formatBarCode(productFromDatabase.CODIGO_DE_BARRAS) === formatBarCode(productToFix.CODIGO_DE_BARRAS))

                    if (productFoundByBarCode) {
                        console.log('Produto encontrado: ' + productFoundByBarCode.DESCRICAO + ' - ' + productFoundByBarCode.CODIGO_DE_BARRAS)
                        productFoundByBarCode.CODIGO_DE_BARRAS = formatBarCode(productFoundByBarCode.CODIGO_DE_BARRAS)
                        const newProduct = cloneDataFrom(productFoundByBarCode, productToFix)
                        newProducts.push(newProduct)
                    } else {
                        console.log('Produto não encontrado! ' + productToFix.DESCRICAO + ' - ' + productToFix.CODIGO_DE_BARRAS)
                        const productNotFounded = formatProductNotFounded(productToFix)
                        newProducts.push(productNotFounded)
                    }
                } else {
                    console.log('Produto não tem codigo de barras: ' + productToFix.DESCRICAO + ' - ' + productToFix.CODIGO_DE_BARRAS)
                    const productNotFounded = formatProductNotFounded(productToFix)
                    newProducts.push(productNotFounded)
                }
                if (newProducts.length === productsToFixData.length) {
                    downloadAsExcel(newProducts)
                }
            }, 1)
        })
    }
}

function checkFieldsNames(product, tableName) {
    if (!product.CODIGO) {
        alert('Altere o nome da coluna com os códigos dos produtos para "CODIGO" na tabela ' + tableName)
        return false
    } else if (!product.CODIGO_DE_BARRAS) {
        alert('Altere o nome da coluna com os códigos de barras para "CODIGO_DE_BARRAS" na tabela ' + tableName)
        return false
    } else if (!product.DESCRICAO) {
        alert('Altere o nome da coluna com as descrições dos produtos para "DESCRICAO" na tabela ' + tableName)
        return false
    } else {
        return true
    }
}

function formatBarCode(barCodeNumber) {
    let onlyNumbersFromString = String(barCodeNumber).replace(/[^0-9]/g, '')
    let formattedBarCode = Number(onlyNumbersFromString)
    return String(formattedBarCode)
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
                    if (rowObject.length === 0) {
                        alert('A planilha mãe está vazia!')
                    } else {
                        createParametersObject(Object.keys(rowObject[0]))
                    }
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
                    if (rowObject.length === 0) {
                        alert('A planilha filha está vazia!')
                    }
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
    localParameters = {}
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

function cloneDataFrom(productFromDatabase, productToFix) {
    const newProduct = {}
    const parametersNames = Object.keys(localParameters)
    parametersNames.forEach(function (parameter, index) {
        if (parameter === "CODIGO") {
            newProduct["CODIGO"] = productToFix.CODIGO
        } else if (parameter === "DESCRICAO") {
            newProduct["DESCRICAO"] = productToFix.DESCRICAO
        } else if (parameter === "CODIGO_DE_BARRAS") {
            newProduct["CODIGO_DE_BARRAS"] = productToFix.CODIGO_DE_BARRAS
        } else {
            const parameterValue = localParameters[parameter]
            if (parameterValue === true) {
                newProduct[parameter] = productFromDatabase[String(parameter).toUpperCase()]
            }
        }
        if (parametersNames.length - 1 === index) {
            newProduct["STATUS"] = "PRODUTO VERIFICADO"
        }
    })
    return newProduct
}

function formatProductNotFounded(product) {
    const newProduct = {}
    const parametersNames = Object.keys(localParameters)
    parametersNames.forEach(function (parameter, index) {
        if (parameter === "CODIGO") {
            newProduct["CODIGO"] = product.CODIGO
        } else if (parameter === "DESCRICAO") {
            newProduct["DESCRICAO"] = product.DESCRICAO
        } else if (parameter === "CODIGO_DE_BARRAS") {
            newProduct["CODIGO_DE_BARRAS"] = product.CODIGO_DE_BARRAS
        } else {
            const parameterValue = localParameters[parameter]
            if (parameterValue === true) {
                newProduct[parameter] = ""
            }
        }
        if (parametersNames.length - 1 === index) {
            newProduct["STATUS"] = "PRODUTO NÃO VERIFICADO"
        }
    })
    return newProduct
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
    if (Object.keys(localParameters).length > 0) {
        loadLocalParameters()
    } else {
        loadParametersFromExcel()
    }
}

function loadLocalParameters() {
    Object.keys(localParameters).forEach(function (parameterName, index) {
        createParameterCheckbox(parameterName, localParameters[parameterName])
        if (index === Object.keys(localParameters).length - 1) {
            showParameters()
        }
    })
}

function loadParametersFromExcel() {
    getProductsFromExcel().then(function (products) {
        removeLoaderDiv()
        createParametersObject(Object.keys(products[0]))
        Object.keys(localParameters).forEach(function (parameter, index) {
            createParameterCheckbox(parameter, true)
            if (index === Object.keys(localParameters).length - 1) {
                showParameters()
            }
        })
    })
}

function createParametersObject(headerItems) {
    if (Object.keys(localParameters).length === 0) {
        headerItems.forEach(function (parameter) {
            localParameters[String(parameter).toUpperCase()] = true
        })
    }
}

//------------------------------------------------------------------------------------------------------------------------------------------------
// Create parameters checkbox in html
//------------------------------------------------------------------------------------------------------------------------------------------------

function createParameterCheckbox(parameter, state) {
    const checkBoxMainDiv = document.getElementById('checkBoxDiv')
    const checkBoxDiv = document.createElement('div')
    checkBoxDiv.setAttribute('id', 'divBox')
    checkBoxDiv.setAttribute('class', 'divBox')

    const checkBox = document.createElement('input')
    checkBox.setAttribute('type', 'checkbox')
    checkBox.setAttribute('id', parameter)
    checkBox.setAttribute('class', 'checkBox')
    if (state) { checkBox.setAttribute('checked', state) }
    checkBox.onclick = function () {
        toggleParameter(this)
    }

    const labelCheckBox = document.createElement('label')
    labelCheckBox.setAttribute('for', parameter)
    labelCheckBox.setAttribute('id', 'labelCheckBox')
    labelCheckBox.setAttribute('class', 'labelCheckBox')

    labelCheckBox.innerHTML = parameter
    checkBoxMainDiv.appendChild(checkBoxDiv)
    checkBoxDiv.appendChild(checkBox)
    checkBoxDiv.appendChild(labelCheckBox)
    if (parameter === "CODIGO") { checkBox.disabled = true }
    if (parameter === "DESCRICAO") { checkBox.disabled = true }
    if (parameter === "CODIGO_DE_BARRAS") { checkBox.disabled = true }
}

function toggleParameter(parameter) {
    const parameterValue = parameter.checked
    const parameterName = parameter.getAttribute('id')
    localParameters[parameterName] = parameterValue
    const checkBox = document.getElementById(parameterName)
    checkBox.checked = parameterValue
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