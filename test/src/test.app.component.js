import { Component } from '@angular/core';
import { pingTestServer } from "office-addin-test-helpers";
import * as excelComponent from "./test.excel.app.component";
import * as wordComponent from "./test.word.app.component";
const template = require('./../../src/taskpane/app/app.component.html');
const port = 4201;

@Component({
    selector: 'app-home',
    template
})
export default class AppComponent {
    welcomeMessage = 'Welcome';
    constructor() {
        Office.onReady(async (info) => {
            if (info.host === Office.HostType.Excel || info.host === Office.HostType.Word) {
                const testServerResponse = await pingTestServer(port);
                if (testServerResponse["status"] == 200) {
                    if (info.host === Office.HostType.Excel){
                        const excel = new excelComponent.default();
                        return excel.runTest();
                    } else {
                        const word = new wordComponent.default();
                        return word.runTest();
                    }
                }
            }
        });
    }
}