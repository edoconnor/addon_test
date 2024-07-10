function onOpen() {
  DocumentApp.getUi()
    .createMenu("My Addon")
    .addItem("Show Alert", "showAlert")
    .addItem("Convert to HTML", "convertToHtml")
    .addToUi();
}

function showAlert() {
  DocumentApp.getUi().alert("Hello, world!");
}

function convertToHtml() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var docName = doc.getName();
  var fonts = extractFontsFromBody(body);
  var fontLink = generateGoogleFontsLink(fonts);
  var htmlOutput = convertBodyToHtml(body);
  htmlOutput = wrapInContainer(htmlOutput, fontLink);
  saveHtmlToFile(docName, htmlOutput);
  DocumentApp.getUi().alert("HTML file has been saved to your Google Drive.");
}

function extractFontsFromBody(body) {
  var fonts = new Set();
  var numChildren = body.getNumChildren();

  for (var i = 0; i < numChildren; i++) {
    var child = body.getChild(i);
    extractFontsFromElement(child, fonts);
  }

  return Array.from(fonts);
}

function extractFontsFromElement(element, fonts) {
  if (
    element.getType() == DocumentApp.ElementType.PARAGRAPH ||
    element.getType() == DocumentApp.ElementType.LIST_ITEM
  ) {
    for (var i = 0; i < element.getNumChildren(); i++) {
      var child = element.getChild(i);
      if (child.getType() == DocumentApp.ElementType.TEXT) {
        var attributes = child.getAttributes();
        if (attributes[DocumentApp.Attribute.FONT_FAMILY]) {
          fonts.add(attributes[DocumentApp.Attribute.FONT_FAMILY]);
        }
      }
    }
  } else if (element.getType() == DocumentApp.ElementType.TABLE) {
    var table = element.asTable();
    for (var r = 0; r < table.getNumRows(); r++) {
      for (var c = 0; c < table.getNumColumns(); c++) {
        extractFontsFromElement(table.getCell(r, c).getChild(0), fonts);
      }
    }
  }
}

function generateGoogleFontsLink(fonts) {
  var baseUrl = "https://fonts.googleapis.com/css2?";
  var query =
    "family=" +
    fonts
      .map(function (font) {
        return font.replace(/ /g, "+") + ":ital,wght@0,100..700;1,100..700";
      })
      .join("&family=") +
    "&display=swap";
  return baseUrl + query;
}

function convertBodyToHtml(body) {
  var numChildren = body.getNumChildren();
  var html = "";

  for (var i = 0; i < numChildren; i++) {
    var child = body.getChild(i);
    html += convertElementToHtml(child);
  }

  return html;
}

function convertElementToHtml(element) {
  var elementType = element.getType();
  switch (elementType) {
    case DocumentApp.ElementType.PARAGRAPH:
      return convertParagraphToHtml(element.asParagraph());

    case DocumentApp.ElementType.LIST_ITEM:
      return convertListItemToHtml(element.asListItem());

    case DocumentApp.ElementType.INLINE_IMAGE:
      return convertImageToHtml(element.asInlineImage());

    default:
      return "";
  }
}

function convertParagraphToHtml(paragraph) {
  var style = paragraph.getHeading();
  var text = paragraph.getText();
  var attributes = paragraph.getAttributes();
  var htmlText = applyTextStyle(text, attributes);

  if (style != DocumentApp.ParagraphHeading.NORMAL) {
    return (
      "<h" +
      getHeadingLevel(style) +
      ">" +
      htmlText +
      "</h" +
      getHeadingLevel(style) +
      ">"
    );
  } else {
    var fontSize = attributes[DocumentApp.Attribute.FONT_SIZE];
    var isUnderlined = attributes[DocumentApp.Attribute.UNDERLINE];
    var fontSizePx = fontSize ? fontSize + "pt" : "12pt";

    if (fontSize && fontSize >= 18 && isUnderlined) {
      return "<h3>" + htmlText + "</h3>";
    } else {
      return '<p style="font-size:' + fontSizePx + ';">' + htmlText + "</p>";
    }
  }
}

function getHeadingLevel(style) {
  switch (style) {
    case DocumentApp.ParagraphHeading.HEADING1:
      return 1;
    case DocumentApp.ParagraphHeading.HEADING2:
      return 2;
    case DocumentApp.ParagraphHeading.HEADING3:
      return 3;
    case DocumentApp.ParagraphHeading.HEADING4:
      return 4;
    case DocumentApp.ParagraphHeading.HEADING5:
      return 5;
    case DocumentApp.ParagraphHeading.HEADING6:
      return 6;
    default:
      return 0;
  }
}

function convertListItemToHtml(listItem) {
  var listTag =
    listItem.getNestingLevel() === 0
      ? listItem.getGlyphType() === DocumentApp.GlyphType.BULLET
        ? "ul"
        : "ol"
      : null;
  var li =
    "<li>" +
    applyTextStyle(listItem.getText(), listItem.getAttributes()) +
    "</li>";

  if (listTag) {
    return "<" + listTag + ">" + li + "</" + listTag + ">";
  } else {
    return li;
  }
}

function convertImageToHtml(inlineImage) {
  var blob = inlineImage.getBlob();
  var base64Image = Utilities.base64Encode(blob.getBytes());
  var mimeType = blob.getContentType();
  return '<img src="data:' + mimeType + ";base64," + base64Image + '">';
}

function applyTextStyle(text, attributes) {
  var htmlText = text;
  var styles = [];

  if (attributes[DocumentApp.Attribute.FONT_FAMILY]) {
    styles.push("font-family:" + attributes[DocumentApp.Attribute.FONT_FAMILY]);
  }
  if (attributes[DocumentApp.Attribute.FONT_SIZE]) {
    styles.push(
      "font-size:" + attributes[DocumentApp.Attribute.FONT_SIZE] + "pt"
    );
  }
  if (attributes[DocumentApp.Attribute.FOREGROUND_COLOR]) {
    styles.push("color:" + attributes[DocumentApp.Attribute.FOREGROUND_COLOR]);
  }
  if (attributes[DocumentApp.Attribute.BOLD]) {
    htmlText = "<b>" + htmlText + "</b>";
  }
  if (attributes[DocumentApp.Attribute.ITALIC]) {
    htmlText = "<i>" + htmlText + "</i>";
  }
  if (attributes[DocumentApp.Attribute.UNDERLINE]) {
    htmlText = "<u>" + htmlText + "</u>";
  }

  if (styles.length > 0) {
    htmlText =
      '<span style="' + styles.join("; ") + '">' + htmlText + "</span>";
  }

  return htmlText;
}

function wrapInContainer(htmlContent, fontLink) {
  return `
      <html>
      <head>
        <style>
          body {
            background-color: #d3d3d3;
            display: flex;
            justify-content: center;
            align-items: center;
            height: 100vh;
            margin: 0;
          }
          .container {
            background-color: white;
            width: 8in;
            height: 10in;
            padding: 1in;
            margin-top: 120px;
            box-shadow: 0 0 10px rgba(0,0,0,0.1);
          }
            
        </style>
        <link href="${fontLink}" rel="stylesheet">
      </head>
      <body>
        <div class="container">
          ${htmlContent}
        </div>
      </body>
      </html>
    `;
}

function saveHtmlToFile(docName, htmlContent) {
  var formattedDocName = docName.toLowerCase().replace(/ /g, "_");
  var htmlFile = DriveApp.createFile(
    formattedDocName + ".html",
    htmlContent,
    MimeType.HTML
  );
  return htmlFile.getUrl();
}
