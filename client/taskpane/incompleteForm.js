(async () => {
    await Office.onReady();
  
    document.getElementById("close_dialog").onclick = sendStringToParentPage;
  
    function sendStringToParentPage() {
      //console.log('context')
      console.log(Office.context)
      Office.context.ui.messageParent('close', {});
  }
  })();