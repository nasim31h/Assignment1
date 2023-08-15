const openpyxl = require('openpyxl');
const webdriver = require('selenium-webdriver');
const { By, Key } = require('selenium-webdriver');
const { Options } = require('selenium-webdriver/chrome');
const { Builder, Capabilities } = require('selenium-webdriver');
const { parse } = require('node-html-parser');
const fs = require('fs');
const path = require('path');

async function searchAll(selectSheet) {
    let LastItem;
    let FirstItem;

    async function loadData(file) {
        const workbook = await openpyxl.loadWorkbook(file);
        const worksheet = workbook.getWorksheet(selectSheet);
        
        const cColumnValues = [];
        worksheet.eachRow({ minCol: 3, maxCol: 3 }, (row, rowIndex) => {
            if (rowIndex >= 2) {
                cColumnValues.push(row.getCell(1).value);
            }
        });
        
        workbook.close();
        return cColumnValues;
    }

    const filePath = 'Excel.xlsx';
    const keywords = await loadData(filePath);

    const chromeOptions = new Options();
    chromeOptions.headless = false;
    const driver = await new Builder().forBrowser('chrome').withCapabilities(Capabilities.chrome()).setChromeOptions(chromeOptions).build();

    for (const keyword of keywords) {
        const googleResults = await performGoogleSearch(keyword, driver);

        if (googleResults.length > 0) {
            FirstItem = googleResults[0];
            LastItem = googleResults[googleResults.length - 1];
        } else {
            console.log("The list is empty.");
        }

        const workbook = await openpyxl.loadWorkbook(filePath);
        const worksheet = workbook.getWorksheet(selectSheet);

        worksheet.getCell(`D${startRow}`).value = LastItem;
        worksheet.getCell(`E${startRow}`).value = FirstItem;
        startRow += 1;

        await workbook.xlsx.writeFile(filePath);
    }

    driver.quit();
}

async function performGoogleSearch(keyword, driver) {
    await driver.get('https://www.google.com/');
    await driver.wait(webdriver.until.elementLocated(By.name('q')), 10000);
    const searchBox = await driver.findElement(By.name('q'));
    await searchBox.sendKeys(keyword, Key.RETURN);

    await driver.wait(webdriver.until.elementLocated(By.css('.tF2Cxc')), 10000);
    const searchResultsHtml = await driver.getPageSource();
    const searchResults = parse(searchResultsHtml).querySelectorAll('.tF2Cxc');

    const textList = [];
    searchResults.forEach(result => {
        const spanElement = result.querySelector('span');
        if (spanElement) {
            const text = spanElement.text.trim();
            if (text) {
                textList.push(text);
            }
        }
    });

    textList.sort((a, b) => a.length - b.length);
    return textList;
}

async function readLoadData(filePath) {
    const workbook = await openpyxl.loadWorkbook(filePath);
    for (const sheetName of workbook.sheetNames) {
        const sheet = workbook.getWorksheet(sheetName);
        const currentDate = new Date();
        const todayStr = currentDate.toLocaleDateString('en-US', { weekday: 'long' });
        if (sheetName === todayStr) {
            await searchAll(sheetName);
        }
        console.log('\n');
        console.log('\n');
    }
}

let startRow = 3;
const filePath = path.join(__dirname, 'Excel.xlsx');
readLoadData(filePath);
