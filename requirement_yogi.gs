function onOpen(e) {
  DocumentApp.getUi().createAddonMenu().addItem("Start", "openMenu").addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function openMenu() {
    try {
        let ui = HtmlService.createHtmlOutputFromFile("menu.html").setTitle("Requirement Parsing");
        DocumentApp.getUi().showSidebar(ui);
    } catch (error) {
        Logger.log("Error opening menu: " + error);
        DocumentApp.getUi().alert("Error opening the menu. Please try again.");
        throw new Error("Error opening the menu. Please try again.");

    }
}

function goButton(eof, sidebar) {
  try {
      let body = DocumentApp.getActiveDocument().getBody();
      recap = handleRequirementsText(body);
      if (recap != {}) {
          if (eof) {
              writeRecapEOF(recap);
          }
          if (sidebar) {
              writeRecapHTML(recap);
          }
      }
  } catch (error) {
      Logger.log("Error processing requirements: " + error);
      DocumentApp.getUi().alert("Error processing requirements. Please try again.");
      throw new Error("Error processing requirements. Please try again.");
  }
}

function handleRequirementsText(body) {

  let recap = {};
  
  let result = getJSON();
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
    recapHTML += `<li><a href="${encodeURI(recap[element])}" target="_blank">${element}</a></li>`;
  }
  recapHTML += "</ul>"

  let ui = HtmlService.createHtmlOutput(`<!DOCTYPE html><meta charset="UTF-8"><html><body>${recapHTML}</body></html>`).setTitle("Requirements").setSandboxMode(HtmlService.SandboxMode.IFRAME);
  DocumentApp.getUi().showSidebar(ui);
}

function writeRecapEOF(recap) {

  let body = DocumentApp.getActiveDocument().getBody();

  body.appendPageBreak();

  let recapTitle = "Requirements :";

  let title = body.appendParagraph(recapTitle)
  title.setHeading(DocumentApp.ParagraphHeading.HEADING2);


  for (let element in recap) {
    let listItem = body.appendListItem(element).setGlyphType(DocumentApp.GlyphType.BULLET);
    listItem.setLinkUrl(recap[element]);
  }
}

function getJSON() {
  try {
    let url = "https://ww1.requirementyogi.cloud/nuitdelinfo/search?offset=OFFSET_VAL"
    let result = []
    let current = 0
    do {
      var response = UrlFetchApp.fetch(url.replace("OFFSET_VAL", current.toString()), {'muteHttpExceptions': true});
      var parsed = JSON.parse(response);
      result = result.concat(parsed["results"]);
      current += parsed["limit"];
    } while (current < response["total"]);
    
    return result;
  }
  catch (error) {
    Logger.log("Error fetching data from the API: " + error);
    DocumentApp.getUi().alert("Failed to fetch data from the API. Please check your internet connection and try again.");
    throw new Error("Failed to fetch data from the API. Please check your internet connection and try again.");
  }
}
