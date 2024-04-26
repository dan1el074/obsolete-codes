const XLSX = require("xlsx");
let data = {
    path: "./examples/codes.xlsx",
    codes: [],
    oldCodes: [],
    allCodes: []
}

start()
console.log("Códigos revisados: ", data.oldCodes)

function start() {
    data.codes = [];
    const workbook = XLSX.readFile(data.path);
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const resExcel = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    resExcel.forEach((array) => {
        let onlyCodes = String(array[0]).match(/\d{3,9}-?\d{1,3}/);
        if (onlyCodes) {
            data.allCodes.push(onlyCodes.input);
        }
        
        let onlyCodesWithHifen = String(array[0]).match(/\d+-\d+/);
        if (onlyCodesWithHifen) {
            data.codes.push(onlyCodes.input);
        }
    });

    console.log("Todos os códigos: ", data.codes)

    data.codes.forEach(code => {
        const codeArray = code.split('-')
        let counter = Number(String(codeArray[1])[0]) == 0 ? Number(String(codeArray[1]).slice(1)) : Number(String(codeArray[1]))

        if(counter == 1) {
            let verifiedCode = data.allCodes.find(res => res == codeArray[0])

            if(verifiedCode) {
                let verifyVerifiedCode = data.oldCodes.find(res => res == verifiedCode)
                if(!verifyVerifiedCode) {
                    data.oldCodes.push(verifiedCode)
                }
            }
        }

        let found = data.codes.find(code => {
            if(counter >= 9) {
                return code == `${codeArray[0]}-${counter + 1}`
            } else {
                return code == `${codeArray[0]}-0${counter + 1}`
            }
        })

        if(found) {
            return
        } 
        
        if(counter == 1) {
            return
        }

        for(i=counter-1; i>=0; i--) {
            let verifiedCode;

            if(i >= 10) {
                verifiedCode = data.codes.find(res => res === `${codeArray[0]}-${i}`)
            } else if(i === 0) {
                verifiedCode = data.allCodes.find(res => res === `${codeArray[0]}`)
            } else {
                verifiedCode = data.codes.find(res => res === `${codeArray[0]}-0${i}`)
            }
            
            if(verifiedCode) {
                let verifyVerifiedCode = data.oldCodes.find(res => res === verifiedCode)
                if(!verifyVerifiedCode) {
                    data.oldCodes.push(verifiedCode)
                }
            }
        }
    })

    const newWorkbook = XLSX.utils.book_new();
    const sheetData = data.oldCodes.map(item => [item]);
    const sheet = XLSX.utils.aoa_to_sheet(sheetData);

    XLSX.utils.book_append_sheet(newWorkbook, sheet, 'Result');
    XLSX.writeFile(newWorkbook, '../projetos-obsoletos.xlsx');
}
