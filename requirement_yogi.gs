function onOpen(e) {
    DocumentApp.getUi().createAddonMenu()
        .addItem('Start', 'main')
        .addToUi();
  }
  
  function onInstall(e) {
    onOpen(e);
  }
  
  function handleRequirementsText(body) {

    var recap = {};
    
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

          elemText.setLinkUrl(startText, endText, element["canonicalUrl"]);
          recap[element["key"]] = element["canonicalUrl"];

          current = found;
        }
      }
    });

    return recap;
  }

  function writeRecapHTML(recap) {
    let recapHTML = "";
    recapHTML += "<h3>Requirements :</h3>"

    recapHTML += "<ul>"
    for (let element in recap) {
      Logger.log(recap[element]);
      recapHTML += `<li><a href="${recap[element]}" target="_blank">${element}</a></li>`;
    }
    recapHTML += "</ul>"

    var ui = HtmlService.createHtmlOutput(`<!DOCTYPE html><meta charset="UTF-8"><html><body>${recapHTML}</body></html>`).setTitle("Requirements").setSandboxMode(HtmlService.SandboxMode.IFRAME);
    DocumentApp.getUi().showSidebar(ui);
  }

  function writeRecapEOF(recap) {
    var doc = DocumentApp.getActiveDocument();
  
    var body = doc.getBody();

    body.appendPageBreak();

    var recapTitle = "Requirements :\n";

    var title = body.appendParagraph(recapTitle)
    var titleText = title.editAsText();
    title.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    //titleText.setFontSize(titleText.getFontSize() * 2);

    for (element in recap){
      var paragraph = body.appendParagraph(element);

      paragraph.setLinkUrl(recap[element]);
    }
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

  function main() {
    var body = DocumentApp.getActiveDocument().getBody();
    
    recap = handleRequirementsText(body)
    if (recap != {}){
      //writeRecapHTML(recap);
      writeRecapEOF(recap);
    }
  }
