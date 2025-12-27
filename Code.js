/**
 * Google Docs Add-on: StickyRolo
 * The script interacts with document tabs to fetch structured metadata
 * and provides context-aware suggestions based on the user's cursor position.
 * The metadata is organized using headings, which define hierarchy. Each entry can have descriptions or properties.
 * The add-on includes a sidebar UI for user interaction. Users can change how much context to include when looking up values.
 */

function onOpen() {
  DocumentApp.getUi().createMenu('StickyRolo')
    .addItem('Open Context Sidebar', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('StickyRolo')
    .setWidth(300);
  DocumentApp.getUi().showSidebar(html);
}

function getMetadataContent(tabName) {
  // Default to "Metadata" if no tab name provided
  var targetTabName = tabName || "Metadata";
  var docId = DocumentApp.getActiveDocument().getId();
  
  try {
    var doc = Docs.Documents.get(docId, {
      'includeTabsContent': true
    });

    // Root level objects (mostly for single-tab or legacy docs)
    var rootInlineObjects = doc.inlineObjects || {};
    var rootPositionedObjects = doc.positionedObjects || {};

    if (!doc.tabs) {
      // Fallback for single-tab docs
      return { items: parseBodyContent(doc.body.content, rootInlineObjects, rootPositionedObjects) };
    }

    var metadataTab = doc.tabs.find(tab => tab.tabProperties.title === targetTabName);

    if (!metadataTab) {
      return { error: "Tab named '" + targetTabName + "' not found." };
    }

    // IMPORTANT: Access objects specific to the tab if they exist, fallback to root
    // Note: The API structure for tabs usually nests these under documentTab
    var tabInlineObjects = metadataTab.documentTab.inlineObjects || rootInlineObjects;
    var tabPositionedObjects = metadataTab.documentTab.positionedObjects || rootPositionedObjects;

    var items = parseBodyContent(metadataTab.documentTab.body.content, tabInlineObjects, tabPositionedObjects);
    return { items: items };

  } catch (e) {
    return { error: "Error accessing tab: " + e.message };
  }
}

function getTabList() {
  var docId = DocumentApp.getActiveDocument().getId();
  try {
    var doc = Docs.Documents.get(docId, { 'includeTabsContent': false });
    if (!doc.tabs) return [];
    
    return doc.tabs.map(function(tab) {
      return { title: tab.tabProperties.title };
    });
  } catch (e) {
    return [];
  }
}

/**
 * Parses content into structured metadata blocks with hierarchy.
 * Supports H1-H6 as categories/path.
 * Automatically populates descriptions for parent items if they are empty.
 */
function parseBodyContent(contentArray, inlineObjects, positionedObjects) {
  var metadata = {};
  var buffer = []; 
  var headingStack = []; 
  var currentPath = "";

  // Ensure objects are defined
  inlineObjects = inlineObjects || {};
  positionedObjects = positionedObjects || {};

  if (!contentArray) return metadata;

  // 1. Extract lines with style info
  var lines = [];
  contentArray.forEach(function(element) {
    if (element.paragraph) {
      var text = "";
      var imageUrl = null;

      // Check for Inline Objects (inside text elements)
      element.paragraph.elements.forEach(function(el) {
        if (el.textRun && el.textRun.content) {
          text += el.textRun.content;
        } else if (el.inlineObjectElement && inlineObjects) {
          var objectId = el.inlineObjectElement.inlineObjectId;
          if (inlineObjects[objectId] &&
              inlineObjects[objectId].inlineObjectProperties &&
              inlineObjects[objectId].inlineObjectProperties.embeddedObject &&
              inlineObjects[objectId].inlineObjectProperties.embeddedObject.imageProperties) {

             if (!imageUrl) {
               imageUrl = inlineObjects[objectId].inlineObjectProperties.embeddedObject.imageProperties.contentUri;
             }
          }
        }
      });
      
      // Check for Positioned Objects (anchored to paragraph)
      if (!imageUrl && element.paragraph.positionedObjectIds && positionedObjects) {
        element.paragraph.positionedObjectIds.forEach(function(objectId) {
          if (!imageUrl && positionedObjects[objectId] &&
              positionedObjects[objectId].positionedObjectProperties &&
              positionedObjects[objectId].positionedObjectProperties.embeddedObject &&
              positionedObjects[objectId].positionedObjectProperties.embeddedObject.imageProperties) {

            imageUrl = positionedObjects[objectId].positionedObjectProperties.embeddedObject.imageProperties.contentUri;
          }
        });
      }

      var style = "NORMAL_TEXT";
      if (element.paragraph.paragraphStyle && element.paragraph.paragraphStyle.namedStyleType) {
        style = element.paragraph.paragraphStyle.namedStyleType;
      }

      lines.push({
        text: text.trim(),
        style: style,
        imageUrl: imageUrl
      });
    }
  });

  // 2. Helper to process buffer into a metadata item
  function processBuffer(buf, pathStr) {
    // IMPORTANT: Only shift if text is empty AND NO IMAGE.
    while(buf.length > 0 && buf[0].text === "" && !buf[0].imageUrl) {
      buf.shift();
    }
    
    if (buf.length === 0) return;

    var name = buf[0].text;
    var description = "";
    if (buf.length > 1 && !buf[1].text.includes(":") && !buf[1].text.startsWith("_")) {
       description = buf[1].text;
    }

    var properties = {};
    var foundImageUrl = null;

    // Check first line for image if present
    if (buf[0].imageUrl) foundImageUrl = buf[0].imageUrl;
    // Check second line
    if (!foundImageUrl && buf.length > 1 && buf[1].imageUrl) foundImageUrl = buf[1].imageUrl;

    var startIndex = (description !== "") ? 2 : 1;

    for (var i = startIndex; i < buf.length; i++) {
      var line = buf[i].text;

      // Also check for image in subsequent lines
      if (!foundImageUrl && buf[i].imageUrl) {
        foundImageUrl = buf[i].imageUrl;
      }

      if (line.startsWith("_")) break; 

      if (line.includes(":")) {
        var parts = line.split(":");
        var key = parts[0].trim();
        var val = parts.slice(1).join(":").trim();
        if (key && val) {
          properties[key] = val;
        }
      }
    }

    if (name) {
      metadata[name] = {
        description: description,
        properties: properties,
        category: pathStr,
        imageUrl: foundImageUrl
      };
    }
  }

  // 3. Iterate lines
  lines.forEach(function(lineObj) {
    var headingMatch = lineObj.style.match(/^HEADING_(\d)$/);
    
    if (headingMatch) {
      var level = parseInt(headingMatch[1], 10); 
      
      if (buffer.length > 0) {
        processBuffer(buffer, currentPath);
        buffer = [];
      }
      
      var parentPath = headingStack.slice(0, level - 1).join(" > ");
      currentPath = parentPath;
      
      // Treat heading as an item itself
      buffer.push(lineObj); 
      
      headingStack[level - 1] = lineObj.text;
      headingStack.length = level;
      
    } else if (lineObj.text === "" && !lineObj.imageUrl) {
      // Empty line with no image acts as separator
      if (buffer.length > 0) {
        processBuffer(buffer, currentPath);
        buffer = [];
      }
      currentPath = headingStack.join(" > ");
      
    } else {
      if (buffer.length === 0) {
        currentPath = headingStack.join(" > ");
      }
      buffer.push(lineObj);
    }
  });

  if (buffer.length > 0) {
    processBuffer(buffer, currentPath);
  }

  // 4. Post-processing: Fill empty descriptions with child entries
  Object.keys(metadata).forEach(function(parentName) {
    var item = metadata[parentName];
    // Only if description is effectively empty
    if (!item.description || item.description.trim() === "") {
      var children = [];
      
      // Calculate the expected category string for immediate children
      var expectedChildCategory = item.category ? (item.category + " > " + parentName) : parentName;

      Object.keys(metadata).forEach(function(childName) {
        if (metadata[childName].category === expectedChildCategory) {
          children.push(childName);
        }
      });
      
      if (children.length > 0) {
        item.description = "Entries: " + children.join(", ");
      }
    }
  });
  
  return metadata;
}

/**
 * Gets the text context from the user's cursor.
 * Accepts an optional 'lookaround' parameter to include preceding/following paragraphs.
 */
function getCurrentContext(lookaround) {
  var limit = lookaround || 0;
  var doc = DocumentApp.getActiveDocument();
  var selection = doc.getSelection();
  var cursor = doc.getCursor();
  
  var coreText = "";
  var startBlock = null;
  var endBlock = null;

  // Helper to find the Paragraph or ListItem container
  function getBlockParent(element) {
    while (element) {
      var type = element.getType();
      if (type === DocumentApp.ElementType.PARAGRAPH || type === DocumentApp.ElementType.LIST_ITEM) {
        return element;
      }
      element = element.getParent();
    }
    return null;
  }

  // Helper to extract text safely from blocks
  function getTextSafe(el) {
     if (!el) return "";
     var t = el.getType();
     if (t === DocumentApp.ElementType.PARAGRAPH) return el.asParagraph().getText();
     if (t === DocumentApp.ElementType.LIST_ITEM) return el.asListItem().getText();
     return "";
  }

  // 1. Capture the Core Text and identify boundaries (Start/End Block)
  if (selection) {
    var elements = selection.getRangeElements();
    if (elements.length > 0) {
      startBlock = getBlockParent(elements[0].getElement());
      endBlock = getBlockParent(elements[elements.length - 1].getElement());
    }

    // Accumulate actual selection text
    for (var i = 0; i < elements.length; i++) {
      var element = elements[i].getElement();
      var type = element.getType();
      if (type === DocumentApp.ElementType.PARAGRAPH) {
        coreText += element.asParagraph().getText() + " ";
      } else if (type === DocumentApp.ElementType.LIST_ITEM) {
        coreText += element.asListItem().getText() + " ";
      } else if (type === DocumentApp.ElementType.TEXT) {
        coreText += element.asText().getText() + " ";
      }
    }
  } 
  else if (cursor) {
    startBlock = getBlockParent(cursor.getElement());
    endBlock = startBlock; // Cursor is a single point, so start == end
    if (startBlock) {
      coreText = getTextSafe(startBlock);
    }
  }

  if (!startBlock) return "";

  // 2. Look Backwards (Prepend text)
  var prefix = "";
  var curr = startBlock;
  var count = 0;
  var safety = 0;

  while (count < limit && safety < 100) {
    curr = curr.getPreviousSibling();
    if (!curr) break;
    
    var txt = getTextSafe(curr);
    // Only count/add if the paragraph has content
    if (txt.trim().length > 0) {
      prefix = txt + " " + prefix;
      count++;
    }
    safety++;
  }

  // 3. Look Forwards (Append text)
  var suffix = "";
  curr = endBlock || startBlock; 
  count = 0;
  safety = 0;

  while (count < limit && safety < 100) {
    curr = curr.getNextSibling();
    if (!curr) break;
    
    var txt = getTextSafe(curr);
    // Only count/add if the paragraph has content
    if (txt.trim().length > 0) {
      suffix += " " + txt;
      count++;
    }
    safety++;
  }

  return prefix + " " + coreText + " " + suffix;
}