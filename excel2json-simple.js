#!/usr/bin/env node

/**
 * Simple Excel2JSON for Mac - Converts Excel to JSON without special markers
 * Handles regular Excel tables and converts them to structured JSON
 */

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

// Configuration
const config = {
    sourceFolder: process.cwd(),
    targetFolder: 'output',
    prettyOutput: true
};

function getPrettyValue(value) {
    if (value == null || value === '') return "";
    if (typeof(value) === "number") return value;
    if (typeof(value) === "string" && !isNaN(value) && value !== "") return Number(value);
    return String(value);
}

function convertTableToObjects(sheetData) {
    if (sheetData.length < 2) return [];
    
    const headers = sheetData[0];
    const objects = [];
    
    for (let i = 1; i < sheetData.length; i++) {
        const row = sheetData[i];
        const obj = {};
        
        for (let j = 0; j < headers.length; j++) {
            const header = headers[j];
            if (header && header !== '') {
                obj[header] = getPrettyValue(row[j]);
            }
        }
        
        // Only add non-empty objects
        if (Object.values(obj).some(val => val !== "")) {
            objects.push(obj);
        }
    }
    
    return objects;
}

function convertTableToKeyValue(sheetData) {
    if (sheetData.length < 2) return {};
    
    const keyValueObj = {};
    
    for (let i = 1; i < sheetData.length; i++) {
        const row = sheetData[i];
        if (row[0] && row[1] !== undefined) {
            keyValueObj[getPrettyValue(row[0])] = getPrettyValue(row[1]);
        }
    }
    
    return keyValueObj;
}

function parseExcel(excelFile) {
    console.log(`\nLoading: ${excelFile}`);
    
    try {
        const workbook = XLSX.readFile(excelFile);
        const rootObject = {};
        
        for (const sheetName of workbook.SheetNames) {
            if (sheetName.startsWith('!')) {
                console.log(`Skipped sheet '${sheetName}', '!' prefix detected`);
                continue;
            }
            
            console.log(`Parse sheet: ${sheetName}`);
            const worksheet = workbook.Sheets[sheetName];
            const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
            
            if (sheetData.length === 0) continue;
            
            // Try to determine the best conversion method based on data structure
            const headers = sheetData[0];
            const hasIdColumn = headers.some(h => h && h.toString().toLowerCase().includes('id'));
            const hasKeyValueStructure = headers.length === 2;
            
            let convertedData;
            if (hasKeyValueStructure) {
                // Simple key-value pairs
                convertedData = convertTableToKeyValue(sheetData);
            } else if (hasIdColumn) {
                // Array of objects with ID as key
                const objects = convertTableToObjects(sheetData);
                const idColumn = headers.find(h => h && h.toString().toLowerCase().includes('id'));
                const idIndex = headers.indexOf(idColumn);
                
                convertedData = {};
                objects.forEach(obj => {
                    const id = obj[idColumn];
                    if (id !== undefined && id !== '') {
                        convertedData[id] = obj;
                    }
                });
            } else {
                // Array of objects
                convertedData = convertTableToObjects(sheetData);
            }
            
            // Use sheet name as key, clean it up
            const cleanSheetName = sheetName.replace(/[（(].*[）)]/g, '').trim();
            rootObject[cleanSheetName] = convertedData;
        }
        
        return JSON.stringify(rootObject, null, config.prettyOutput ? '\t' : '');
    } catch (error) {
        console.error(`Error parsing ${excelFile}:`, error.message);
        throw error;
    }
}

function saveJson(excelFile, jsonString) {
    const jsonFileName = path.basename(excelFile, path.extname(excelFile)) + '.json';
    const jsonPath = path.join(config.targetFolder, jsonFileName);
    
    if (!fs.existsSync(config.targetFolder)) {
        fs.mkdirSync(config.targetFolder, { recursive: true });
    }
    
    fs.writeFileSync(jsonPath, jsonString, 'utf8');
    console.log(`Output: ${jsonPath}`);
}

function getExcelFiles(dir) {
    const files = fs.readdirSync(dir);
    return files.filter(file => 
        file.endsWith('.xlsx') || file.endsWith('.xls')
    ).map(file => path.join(dir, file));
}

// Main execution
function main() {
    try {
        const args = process.argv.slice(2);
        let excels = [];
        
        if (args.length > 0) {
            // Filter Excel files from arguments
            excels = args.filter(arg => 
                arg.endsWith('.xlsx') || arg.endsWith('.xls')
            );
            
            // Check if last argument is output folder
            const lastArg = args[args.length - 1];
            if (!lastArg.endsWith('.xlsx') && !lastArg.endsWith('.xls')) {
                config.targetFolder = lastArg;
            }
        } else {
            excels = getExcelFiles(config.sourceFolder);
        }
        
        if (excels.length === 0) {
            console.log(`There is no excel files in\n${config.sourceFolder}`);
            return;
        }
        
        console.log(`Target folder: ${config.targetFolder}`);
        
        for (const excelFile of excels) {
            const jsonString = parseExcel(excelFile);
            saveJson(excelFile, jsonString);
        }
        
        console.log('\nConversion completed successfully!');
    } catch (error) {
        console.error('Error:', error.message);
        process.exit(1);
    }
}

// Run if this file is executed directly
if (require.main === module) {
    main();
}

module.exports = { parseExcel, saveJson, main };
