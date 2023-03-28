const fs = require('fs');
const path = require('path');

When(/Opening file "([^"]*)"/, async function (filePath) {
    console.log(filePath)

    await driver.maximizeWindow();

    const mainWindowHandle = await driver.getWindowHandle();

    // going to file
    await (await $(`//Button[@Name="File Tab"]`)).click();
    // opening new file
    await (await $(`//ListItem[@Name="Blank workbook"]`)).click();
    // clicking on Data snipper button
    await (await $(`//TabItem[@Name="DATASNIPPER"]`)).click();

    // to get source for the current window
    // fs.writeFileSync('pageSource.xml', await driver.getPageSource());

    // waiting for 5 sec to detect window
    await new Promise(r => setTimeout(r, 5000));

    const helpPopUp = (await driver.getWindowHandles()).find(x => x !== mainWindowHandle);
    if (helpPopUp) {
        await driver.switchToWindow(helpPopUp);
        try { await (await $(`//*[@AutomationId="Skip"]`)).click(); } finally { }
    }

    await driver.switchToWindow(mainWindowHandle);
    // clicking import docs
    await (await $(`//Text[@Name="Import documents"]`)).click();

    // waiting for 5 sec to detect window
    await new Promise(r => setTimeout(r, 4000));

    const fileFullPath = path.join(process.cwd(), filePath);
    console.log(fileFullPath);
    await (await $(`//Edit[@Name="File name:"]`)).sendKeys(fileFullPath.split());
    await (await $(`//Button[@Name="Open"][@AutomationId="1"]`)).click();

    // fs.writeFileSync('pageSource.xml', await driver.getPageSource());

    // waiting for 10 sec to load file
    await new Promise(r => setTimeout(r, 10000));

    //select cell A1 for text
    await (await $(`//DataItem[@AutomationId="A1"]`)).click();

    const documentViewer = await $(`//*[@AutomationId="DocumentViewer"]`);
    const rect = await documentViewer.getAttribute("BoundingRectangle");
    const viewerLocation = Object.fromEntries(rect.trim().split(" ").map(x => {
        const a = x.trim().split(":");
        return [a[0].toLowerCase(), parseInt(a[1])];
    }));

    console.log(viewerLocation);

    await driver.performActions([{
        "type": "pointer",
        "id": "finger1",
        "parameters": { "pointerType": "touch" },
        "actions": [
            { "type": "pointerMove", "duration": 0, "x": viewerLocation.left + 100, "y": viewerLocation.top + 200 },
            { "type": "pointerDown", "button": 1000 },
            { "type": "pause", "duration": 500 },
            { "type": "pointerMove", "origin": "pointer", "x": 700, "y": 300, "duration": 0 },
            { "type": "pointerUp", "button": 0 }
        ]
    }]);

    await driver.releaseActions();

    // await driver.touchAction([
    //     { action: 'longPress', x: viewerLocation.left + 100, y: viewerLocation.top + 200, element: documentViewer },
    //     { action: 'wait', ms: 3000, element: documentViewer },
    //     { action: 'moveTo', x: viewerLocation.left + 800, y: viewerLocation.top + 500, element: documentViewer },
    //     { action: 'release' }
    // ]);
});

Then(/Result should be "([^"]*)"/, async (result) => {
    const text = await (await $(`//DataItem[@AutomationId="A1"]`)).getText();

    expect(text).toEqual(result);
})