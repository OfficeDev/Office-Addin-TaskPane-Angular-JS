import { Component } from '@angular/core';
import { pingTestServer, sendTestResults } from "office-addin-test-helpers";
import * as testHelpers from "./test-helpers";
import * as excel from "../../src/taskpane/app/excel.app.component";
const template = require('./../../src/taskpane/app/app.component.html');
const port = 4201;
let testValues = [];

@Component({
    selector: 'app-home',
    template
})
export default class AppComponent {
    welcomeMessage = 'Welcome';
    constructor() {
        Office.onReady(async () => {
            const testServerResponse = await pingTestServer(port);
            if (testServerResponse["status"] == 200) {
                this.runTest();
            }
        });
    }

    async runTest() {
        return new Promise(async (resolve, reject) => {
            try {
                // Execute taskpane code
                const excelComponent = new excel.default();
                await excelComponent.run();

                // Get output of executed taskpane code
                await Excel.run(async context => {
                    const range = context.workbook.getSelectedRange();
                    const cellFill = range.format.fill;
                    cellFill.load('color');
                    await context.sync();

                    testHelpers.addTestResult(testValues, "fill-color", cellFill.color, "#FFFF00");
                    await sendTestResults(testValues, port);
                    testValues.pop();
                    await testHelpers.closeWorkbook();
                    resolve();
                });
            } catch (err) {
                reject(err);
            }
        });
    }
}