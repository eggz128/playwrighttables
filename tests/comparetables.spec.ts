import { test, expect } from '@playwright/test';
import excel from 'exceljs'

test('Compare webtable to excel sheet', async ({ page }) => {
    page.goto('https://www.edgewordstraining.co.uk/webdriver2/docs/forms.html')
    await page.locator('#textInput').click();
    await page.locator('#textInput').fill('Hello World');
    await page.locator('#textArea').click();
    await page.locator('#textArea').fill('Multiline');
    await page.locator('#textArea').press('Enter');
    await page.locator('#textArea').fill('Multiline\ndata');
    await page.locator('#textArea').press('Enter');
    await page.locator('#textArea').fill('Multiline\ndata\nentry');
    await page.locator('#checkbox').check();
    await page.locator('#select').selectOption('Selection Two');
    await page.locator('#two').check();
    await page.getByRole('link', { name: 'Submit' }).click();
    await page.locator('#formResults').waitFor({ state: 'visible' })
    //Data entry setup done

    // Extract the data from the web page table in to 2d array
    const webPageTable: string[][] = await page.$$eval('#formResults table tr', rows => {
        return rows.map(row => {
            const cells = row.querySelectorAll('td');
            return Array.from(cells).map(cell => cell.textContent ?? ""); //Could trim() to remove whitespace web page has around cell text 
        });
    });

    // Read Excel file using Exceljs
    const workbook = new excel.Workbook();
    await workbook.xlsx.readFile('./tests/excel/table.xlsx');
    const worksheet = workbook.getWorksheet('Sheet1');
    // Initialize 2D array
    const excelSheet: string[][] = [];
    worksheet.eachRow(row => {
        // Initialize array for current row
        const rowCellArray: string[] = [];

        // Loop through each cell in the row
        row.eachCell(cell => {
            // Push the cell value to the row array
            rowCellArray.push(cell.value?.toString() ?? "");
        });

        // Push the row array to the 2D array
        excelSheet.push(rowCellArray);
    });


    //Assert to compare -- fails largely due to whitespace in captured web content
    expect(webPageTable).toEqual(excelSheet)
})
