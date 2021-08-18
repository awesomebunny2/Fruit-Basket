/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global console, document, Excel, Office */

Office.addin.setStartupBehavior(Office.StartupBehavior.load);

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    
    Excel.run(async context => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.onChanged.add(onChange);

    await context.sync();
    console.log("A handler has been registered for the onChanged event.");
    });
  };

/**
 * Handle the changed event from the worksheet.
 *
 * @param event The event information from Excel
 */
  async function onChange(event) {
    return Excel.run(function(context) {
      return context.sync().then(function() {
        console.log("Change type of event: " + event.changeType);
        console.log("Address of event: " + event.address);
        console.log("Source of event: " + event.source);
      });
    });
  }
});

// Office.onReady((info) => {
//   if (info.host === Office.HostType.Excel) {
//     document.getElementById("sideload-msg").style.display = "none";
//     document.getElementById("app-body").style.display = "flex";
//     // document.getElementById("run").onclick = run;
//   }
// });

// export async function run() {
//   try {
//     await Excel.run(async (context) => {
//       const range = context.workbook.getSelectedRange();
//       range.format.fill.color = "yellow";
//       range.load("address");

//       await context.sync();

//       console.log(`The range address was "${range.address}".`);
//     });
//   } catch (error) {
//     console.error(error);
//   }
// }
