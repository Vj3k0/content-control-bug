/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

(() => {
  // The initialize function must be run each time a new page is loaded
  Office.initialize = (reason) => {
    $(document).ready(() => {
      $('#run').click(run);
    });
  };

  async function run() {
    
    await Word.run(async (context) => {

      // Get selection range
      var range = context.document.getSelection();

      // Create new content control and define content and type
      var cc = range.insertContentControl();
      var ccRange = cc.insertHtml('<strong>Content text</strong>', 'replace');
      ccRange.select('end');
      context.load(cc);
      context.load(ccRange);

      await context.sync();



    });
    
  }
})();
