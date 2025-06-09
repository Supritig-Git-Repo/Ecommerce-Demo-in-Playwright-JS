const ExcelJS = require('exceljs')

class excelUtils {

    constructor(page) {

        this.workbook = new ExcelJS.Workbook();
        this.page=page

    }

    async readExcel(sheetName, filePathInDevice) {

        await this.workbook.xlsx.readFile(filePathInDevice)
        const workSheet = this.workbook.getWorksheet(sheetName);
        workSheet.eachRow((row, rowNumber) => {

            row.eachCell((cell, columnNumber) => {

                console.log(cell.value)

            })
        })

    }


    async traceCordinate(sheetName, filePathInDevice, verifyValue) {
        await this.workbook.xlsx.readFile(filePathInDevice)
        const workSheet = this.workbook.getWorksheet(sheetName)
        workSheet.eachRow((row, rowNumber) => {
            row.eachCell((cell, colNumber) => {
                if (cell.value === verifyValue) {
                    console.log(cell)
                }
            })

        })
    }

    async traceValue(sheetName, filePathInDevice, verifyValue) {
        await this.workbook.xlsx.readFile(filePathInDevice)
        const workSheet = this.workbook.getWorksheet(sheetName)
        workSheet.eachRow((row, rowNumber) => {
            row.eachCell((cell, colNumber) => {
                if (cell.value === verifyValue) {
                    console.log(cell.value)
                }
                
            })

        })
    }

   async writeExcel(sheetName, filePathInDevice, updateValue, oldValue, cellCordinates) {
        const workSheet = this.workbook.getWorksheet(filePathInDevice)
        workSheet.eachRow((row, rowNumber) => {
            row.eachCell((cell, colNumber) => {

                if (cell.value === oldValue) {

                    cell.value = updateValue

                }

                if (cell === cellCordinates) {
                    cell.value = updateValue
                }
            })
        })

       await this.workbook.xlsx.writeFile(filePathInDevice)


    }

    
        
    
}
module.exports = { excelUtils }