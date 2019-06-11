import { Component } from '@angular/core';
import { pingTestServer, sendTestResults } from "office-addin-test-helpers";
import * as testHelpers from "./test-helpers";
import * as word from "../../src/taskpane/app/word.app.component";
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

    async runTest(){
        return new Promise(async (resolve, reject) => {
            try {
                // Execute taskpane code
                const wordComponent = new word.default();
                await wordComponent.run();

                // Get output of executed taskpane code
                Word.run(async (context) => {
                    var firstParagraph = context.document.body.paragraphs.getFirst();
                    firstParagraph.load("text");
                    await context.sync();

                    testHelpers.addTestResult(testValues, "output-message", firstParagraph.text, "Hello World");
                    await sendTestResults(testValues, port);
                    testValues.pop();
                    resolve();
                });
            } catch (err) {
                reject(err);
            }
        });
    }
}