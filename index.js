const express = require('express')
var morgan = require('morgan')

const port = 8080
const app = express()

app.use(morgan());

app.post('/smartsheet', async (req, res) => {
  try {
        const body = req.body;

        // Callback could be due to validation, status change, or actual sheet change events
        if (body.challenge) {
            console.log("Received verification callback");
            // Verify we are listening by echoing challenge value
            res.status(200)
                .json({ smartsheetHookResponse: body.challenge });
        } else if (body.events) {
            console.log(`Received event callback with ${body.events.length} events at ${new Date().toLocaleString()}`);

            // Note that the callback response must be received within a few seconds.
            // If you are doing complex processing, you will need to queue up pending work.
            await processEvents(body);

            res.sendStatus(200);
        } else if (body.newWebHookStatus) {
            console.log(`Received status callback, new status: ${body.newWebHookStatus}`);
            res.sendStatus(200);
        } else {
            console.log(`Received unknown callback: ${body}`);
            res.sendStatus(200);
        }
    } catch (error) {
        console.log(error);
        res.status(500).send(`Error: ${error}`);
    }
});

/*
* Process callback events
* This sample implementation only logs to the console.
* Your implementation might make updates or send data to another system.
* Beware of infinite loops if you make modifications to the same sheet
*/
async function processEvents(callbackData) {
    if (callbackData.scope !== "sheet") {
        return;
    }

    // This sample handles each event individually.
    // Some changes (e.g. column rename) could impact a large number of cells.
    // A complete implementation should consolidate related events and/or cache intermediate data
    for (const event of callbackData.events) {
        // This sample only considers cell changes
        if (event.objectType === "cell") {
            console.log(`Cell changed, row id: ${event.rowId}, column id ${event.columnId}`);

            // Since event data is "thin", we need to read from the sheet to get updated values.
            const options = {
                id: callbackData.scopeObjectId,             // Get sheet id from callback
                queryParameters: {
                    rowIds: event.rowId.toString(),         // Just read one row
                    columnIds: event.columnId.toString()    // Just read one column
                }
            };
            const response = await smarClient.sheets.getSheet(options);
            const row = response.rows[0];
            const cell = row.cells[0];
            const column = response.columns.find(c => c.id === cell.columnId);
            console.log(`**** New cell value "${cell.displayValue}" in column "${column.title}", row number ${row.rowNumber}`);
        }
    }
}

app.listen(port, () => {
  console.log(`Example app listening on port ${port}`)
});

