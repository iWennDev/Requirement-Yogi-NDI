function onOpen(e) {
    DocumentApp.getUi().createAddonMenu()
        .addItem('Start', 'handleRequirementsText')
        .addToUi();
    showSidebar()
  }
  
  function onInstall(e) {
    onOpen(e);
  }
  
  function handleRequirementsText() {
  
    var body = DocumentApp.getActiveDocument().getBody();
    
    var result = getJSON();
    result.forEach(element => {
      let finishedPage = false;
      let current = null;
      while (!finishedPage){
        let found = null;
        if (current == null){
          found = body.findText(element["key"]);
        }
        else {
          found = body.findText(element["key"], current);
        }
        
        if (found == null){
          finishedPage = true;
        }
        else{
          let startText = found.getStartOffset();
          let endText = startText + (element["key"].length - 1);
          let elemText = found.getElement().asText();

          if (elemText.getLinkUrl() == null){
            elemText.setLinkUrl(startText, endText, element["canonicalUrl"]);
          }
          else {
            DocumentApp.getUi().alert(elemText.getLinkUrl());
          }

          current = found;
        }
      }
    });
  }
  
  function getJSON() {
      var url = "https://ww1.requirementyogi.cloud/nuitdelinfo/search?offset=OFFSET_VAL"
      var result = []
      var current = 0
      do {
        var response = UrlFetchApp.fetch(url.replace("OFFSET_VAL", current.toString()), {'muteHttpExceptions': true});
        var parsed = JSON.parse(response);
        result =result.concat(parsed["results"]);
        current += parsed["limit"];
      } while (current < response["total"]);
      
      return result;
  }
