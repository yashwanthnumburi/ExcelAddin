/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();

      // Read the range address.
      range.load(["values","address"]);

      // Update the fill color.
      const imageData=range.getImage();

      await context.sync();
      console.log(imageData);
      const imageElement = document.createElement("img");
      imageElement.src = "data:image/png;base64," + imageData.value; // Convert base64 to image source.
      imageElement.alt = "Preview of selected range";

      // Append the image to a container in the task pane.
      const previewContainer = document.getElementById("preview-container");
      previewContainer.innerHTML = ""; // Clear previous content.
      previewContainer.appendChild(imageElement);
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}
