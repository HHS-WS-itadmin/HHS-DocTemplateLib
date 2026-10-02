/**************************************************************************************************
 * Library inserted in Apps Scripts as DocKit - has the following public functions:
 * 
 * --- FOLDER & CONTENT UTILITIES ---
 * getArchiveFolderHierarchy(docId)
 *    - Retrieves the folder hierarchy (Item, Series, Collection) for a given document.
 * insertBodyContentAtTop(sourceContainer, targetBody)
 *    - Copies all elements from a source container and inserts them at the top of a target body.
 * copyContainerContents(sourceContainer, targetContainer)
 *    - Copies all elements from a source container into a target container (e.g., footers).
 * hexToRgb(hex, defaultHex)
 *    - Converts standard hex color strings (e.g., "#ffffff") into decimal RGB objects for Docs API.
 * 
 * --- ZERO-WIDTH SPACE (ZWS) PLACEHOLDER UTILITIES ---
 * getPlaceholderFromHeader(placeholder, headerType)
 *    - Retrieves hidden metadata from the Document Header based on ZWS boundaries.
 * getPlaceholderFromBody(placeholder)
 *    - Retrieves hidden metadata from the Document Body based on ZWS boundaries.
 * processPlaceholderInHeader(placeholder, newContent, headerType)
 *    - Updates or inserts text in the Document Header using hidden ZWS boundaries.
 * processPlaceholderSpaceBoundaries(placeholder, newContent, url)
 *    - Updates or inserts text in the Document Body using hidden boundaries, optionally hyperlinking it.
 * processPlaceholder(placeholder, newContent) 
 *    - Legacy alias for processPlaceholderSpaceBoundaries.
 * 
 * Note: Functions ending in an underscore (e.g., generateZwsBoundaries_) are private internal helpers 
 * and are not exposed when this library is imported.
 **************************************************************************************************/

/***************************************************************************************************
 * Retrieves the folder hierarchy for the given document based on the HHS Archiving structure.
 * Structure: Shared Drive > #HHS Archive > Collection > Series > Item
 * 
 * @param {string} [docId] - Optional. If omitted, uses the currently active document.
 * @returns {Object} An object containing the names (and IDs) of the folders at each level.
 */
function getArchiveFolderHierarchy(docId) {
  if (!docId) {
    const doc = DocumentApp.getActiveDocument();
    if (!doc) throw new Error("No active document found.");
    docId = doc.getId();
  }
  
  const file = DriveApp.getFileById(docId);
  const path = [];
  
  // Climb up the folder tree (up to 5 levels to reach the Shared Drive root)
  let currentItem = file;
  for (let i = 0; i < 5; i++) {
    const parents = currentItem.getParents();
    if (parents.hasNext()) {
      const parent = parents.next();
      path.push({
        name: parent.getName(),
        id: parent.getId()
      });
      currentItem = parent;
    } else {
      break; // Reached the top or hit an access boundary
    }
  }
  
  // Map the collected path to your expected structure.
  // Because we climbed upwards, index 0 is the immediate parent (Item), index 1 is Series, etc.
  return {
    itemFolder:       path[0] ? path[0].name : "Not Found",
    seriesFolder:     path[1] ? path[1].name : "Not Found",
    collectionFolder: path[2] ? path[2].name : "Not Found",
    archiveRoot:      path[3] ? path[3].name : "Not Found",
    sharedDrive:      path[4] ? path[4].name : "Not Found",
    
    // Returning the raw array as well in case you need the folder IDs later
    rawPath: path 
  };
}

/**********************************************************************************
 * Inserts all elements from a source container into a target body at the very top.
 * Iterates backwards to preserve the original document order.
 */
function insertBodyContentAtTop(sourceContainer, targetBody) {
  const numChildren = sourceContainer.getNumChildren();
  for (let i = numChildren - 1; i >= 0; i--) {
    const child = sourceContainer.getChild(i).copy();
    const type = child.getType();
    
    // Explicitly cast to prevent silent insertion failures
    if (type === DocumentApp.ElementType.PARAGRAPH) {
      targetBody.insertParagraph(0, child.asParagraph());
    } else if (type === DocumentApp.ElementType.TABLE) {
      targetBody.insertTable(0, child.asTable());
    } else if (type === DocumentApp.ElementType.LIST_ITEM) {
      targetBody.insertListItem(0, child.asListItem());
    } else if (type === DocumentApp.ElementType.EQUATION) {
      targetBody.insertEquation(0, child.asEquation());
    }
  }
}

/*************************************************************************************
 * Copies all elements from a source container to a target container (e.g., footers).
 */
function copyContainerContents(sourceContainer, targetContainer) {
  // Capture the empty paragraph left behind by .clear()
  const initialEmptyPara = (targetContainer.getNumChildren() > 0 && targetContainer.getChild(0).getText() === "") 
                             ? targetContainer.getChild(0) : null;

  const numChildren = sourceContainer.getNumChildren();
  for (let i = 0; i < numChildren; i++) {
    const child = sourceContainer.getChild(i).copy();
    const type = child.getType();
    
    // Explicitly cast to prevent silent insertion failures
    if (type === DocumentApp.ElementType.PARAGRAPH) {
      targetContainer.appendParagraph(child.asParagraph());
    } else if (type === DocumentApp.ElementType.TABLE) {
      targetContainer.appendTable(child.asTable());
    } else if (type === DocumentApp.ElementType.LIST_ITEM) {
      targetContainer.appendListItem(child.asListItem());
    } else if (type === DocumentApp.ElementType.EQUATION) {
      targetContainer.appendEquation(child.asEquation());
    }
  }

  // Remove the original empty paragraph so it doesn't push the new footer down
  if (initialEmptyPara && targetContainer.getNumChildren() > 1) {
    initialEmptyPara.removeFromParent();
  }
}

/**************************************************************************************************
 * Retrieves metadata from the Document Header based on the Zero-Width Space Boundaries.
 * Can be restricted to a specific header type.
 * 
 * @param {string} placeholder - The raw key name (e.g., "user_profile")
 * @param {string} [headerType] - Optional. "First Page", "Even Page", or "Default".
 * @returns {string|null} The extracted text, "UNFILLED_PLACEHOLDER" if empty, or null if missing.
 */
function getPlaceholderFromHeader(placeholder, headerType) {
  const doc = DocumentApp.getActiveDocument();
  const targetType = parseHeaderType_(headerType);
  const headers = getActiveHeaders_(doc, targetType);
  
  if (!headers || headers.length === 0) return null;

  const boundaries = generateZwsBoundaries_(placeholder);
  // Strip out brackets if they exist, then format it perfectly
  const cleanPlaceholder = placeholder.replace(/^{{|}}$/g, '');
  const fullPlaceholder = "{{" + cleanPlaceholder + "}}";
  
  // 1. Look for the raw {{placeholder}} FIRST
  const escapedPlaceholder = fullPlaceholder.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  for (const header of headers) {
    const found = header.findText(escapedPlaceholder);
    if (found) {
      // Return a standard short string your main script can check for
      return "UNFILLED_PLACEHOLDER"; 
    }
  }

  // 2. If no {{placeholder}} is found, look for data that has already been populated (bounded by ZWS markers)
  const regex = boundaries.start + ".*?" + boundaries.end;
  for (const header of headers) {
    const found = header.findText(regex);
    if (found) {
      const textElement = found.getElement().asText();
      
      // Extract the exact text match from the document
      const matchText = textElement.getText().substring(
        found.getStartOffset(), 
        found.getEndOffsetInclusive() + 1
      );
      
      // Strip the invisible start and end boundaries, returning only the content
      const content = matchText.substring(boundaries.start.length, matchText.length - boundaries.end.length);
      return content;
    }
  }

  // 3. Not found anywhere in the targeted header(s)
  return null;
}

/***************************************************************************************************
 * Dynamically updates or inserts document content using invisible Zero-Width Space Boundaries.
 * Searches and updates the Document Header. Can be restricted to a specific header type.
 * 
 * @param {string} placeholder - The raw key name (e.g., "user_profile")
 * @param {string} newContent - The text to insert.
 * @param {string} [headerType] - Optional. "First Page", "Even Page", or "Default".
 */
function processPlaceholderInHeader(placeholder, newContent, headerType) {
  // Strip out brackets if they exist, then format it perfectly
  const cleanPlaceholder = placeholder.replace(/^{{|}}$/g, '');
  const fullPlaceholder = "{{" + cleanPlaceholder + "}}";
  const doc = DocumentApp.getActiveDocument();
  
  // Parse the optional target type (returns null if not specified)
  const targetType = parseHeaderType_(headerType);
  
  // Reuses the ZWS generator from the original body script
  const boundaries = generateZwsBoundaries_(placeholder);

  // 1. Try to replace a typed {{placeholder}} FIRST
  if (executeZwsPlaceholderReplacementHeader_(doc, fullPlaceholder, boundaries, newContent, targetType)) return;
  
  // 2. Try to update an existing hidden boundary
  if (executeZwsUpdateHeader_(doc, boundaries, newContent, targetType)) return;
  
  // 3. Fallback: Insert at cursor (if applicable), otherwise append to the targeted header
  executeZwsFallbackInsertionHeader_(doc, boundaries, newContent, targetType);
}

// ==============================================================================
// Header-specific private helper functions
// ==============================================================================

/**
 * Normalizes the user's input string.
 */
function parseHeaderType_(headerType) {
  if (!headerType || typeof headerType !== 'string') return null;
  
  const normalized = headerType.trim().toUpperCase().replace(/\s+/g, '_');
  if (['FIRST_PAGE', 'EVEN_PAGE', 'DEFAULT'].includes(normalized)) {
    return normalized;
  }
  return null;
}

/**
 * Retrieves the targeted header section by traversing the document's underlying structure.
 * If no target is specified, it returns all active headers.
 */
function getActiveHeaders_(doc, targetType) {
  const documentElement = doc.getBody().getParent();
  const numChildren = documentElement.getNumChildren();
  const allHeaders = [];
  
  // Extract all generic header sections from the document structure
  for (let i = 0; i < numChildren; i++) {
    const child = documentElement.getChild(i);
    if (child.getType() === DocumentApp.ElementType.HEADER_SECTION) {
      allHeaders.push(child.asHeaderSection());
    }
  }
  
  // If no specific type was requested, return all found headers
  if (!targetType) return allHeaders;
  
  // Best-effort mapping: GAS doesn't label headers, but they consistently appear in this order:
  // [0] = Default Header, [1] = First Page Header, [2] = Even Page Header
  const targetIndex = targetType === 'DEFAULT' ? 0 : 
                      targetType === 'FIRST_PAGE' ? 1 : 
                      targetType === 'EVEN_PAGE' ? 2 : -1;
                      
  if (targetIndex !== -1 && allHeaders.length > targetIndex) {
    return [allHeaders[targetIndex]];
  }
  
  // Fallback: If the requested header type hasn't been created yet, return empty
  return [];
}

function executeZwsUpdateHeader_(doc, boundaries, newContent, targetType) {
  const headers = getActiveHeaders_(doc, targetType);
  const regex = boundaries.start + ".*?" + boundaries.end;
  
  for (const header of headers) {
    const found = header.findText(regex);
    if (found) {
      const textElement = found.getElement().asText();
      const start = found.getStartOffset();
      const end = found.getEndOffsetInclusive();
      const insertStr = boundaries.start + newContent + boundaries.end;

      const attrs = textElement.getAttributes(start);
      textElement.insertText(start, insertStr);
      textElement.setAttributes(start, start + insertStr.length - 1, attrs); 
      textElement.deleteText(start + insertStr.length, end + insertStr.length);
      return true;
    }
  }
  return false;
}

function executeZwsPlaceholderReplacementHeader_(doc, fullPlaceholder, boundaries, newContent, targetType) {
  const headers = getActiveHeaders_(doc, targetType);
  const escapedPlaceholder = fullPlaceholder.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  
  for (const header of headers) {
    const found = header.findText(escapedPlaceholder);
    if (found) {
      const textElement = found.getElement().asText();
      const start = found.getStartOffset();
      const end = found.getEndOffsetInclusive();
      const insertStr = boundaries.start + newContent + boundaries.end;

      const attrs = textElement.getAttributes(start);
      textElement.insertText(start, insertStr);
      textElement.setAttributes(start, start + insertStr.length - 1, attrs); 
      textElement.deleteText(start + insertStr.length, end + insertStr.length);
      return true;
    }
  }
  return false;
}

function executeZwsFallbackInsertionHeader_(doc, boundaries, newContent, targetType) {
  const cursor = doc.getCursor();
  
  // If the user didn't force a specific header type, check if they left their cursor in ANY header
  if (!targetType && cursor) {
    let element = cursor.getElement();
    let isInHeader = false;
    
    while (element) {
      if (element.getType() === DocumentApp.ElementType.HEADER_SECTION) {
        isInHeader = true;
        break;
      }
      element = element.getParent();
    }

    if (isInHeader) {
      const textElement = cursor.insertText(boundaries.start + newContent + boundaries.end);
      if (textElement) return true;
    }
  }

  // Fallback: Append directly to the targeted header
  const headers = getActiveHeaders_(doc, targetType);
  let targetHeader = headers.length > 0 ? headers[0] : null;
  
  if (!targetHeader) {
    // We cannot programmatically create a First/Even page header via GAS if it doesn't exist.
    if (targetType && targetType !== 'DEFAULT') return false; 
    
    // We CAN create the Default header if it doesn't exist
    targetHeader = doc.getHeader() || doc.addHeader();
  }

  // Check if the last paragraph is empty so we don't create unnecessary vertical spacing
  const paragraphs = targetHeader.getParagraphs();
  if (paragraphs.length > 0 && paragraphs[paragraphs.length - 1].getText().trim() === "") {
    paragraphs[paragraphs.length - 1].setText(boundaries.start + newContent + boundaries.end);
  } else {
    targetHeader.appendParagraph(boundaries.start + newContent + boundaries.end);
  }
  
  return true;
}

/***************************************************************************************************
 * Dynamically updates or inserts document content using invisible Zero-Width Space Boundaries.
 * Only searches the body, not the headers or footers.
 * 
 * @param {string} placeholder - The raw key name (e.g., "user_profile" becomes "{{user_profile}}")
 * @param {string} newContent - The text (metadata) to insert into the document.
 * @param {string} [url] - Optional. The URL to hyperlink the inserted text.
 * **************************************************************************************************
 */
function processPlaceholderSpaceBoundaries(placeholder, newContent, url) {
  // Strip out brackets if they exist, then format it perfectly
  const cleanPlaceholder = placeholder.replace(/^{{|}}$/g, '');
  const fullPlaceholder = "{{" + cleanPlaceholder + "}}";
  
  const doc = DocumentApp.getActiveDocument();
  
  // Generate unique invisible boundaries based on the clean placeholder string
  const boundaries = generateZwsBoundaries_(cleanPlaceholder);

  // Pass the URL down to the execution helpers so they can apply the link during insertion
  // 1. Try to replace a typed {{placeholder}} FIRST
  if (executeZwsPlaceholderReplacement_(doc, fullPlaceholder, boundaries, newContent, url)) return;

  // 2. Try to update an existing hidden boundary
  if (executeZwsUpdate_(doc, boundaries, newContent, url)) return;

  // 3. Fallback to cursor insertion
  if (executeZwsCursorInsertion_(doc, boundaries, newContent, url)) return;

  DocumentApp.getUi().alert(
    "Where should this go?", 
    `We couldn't find '${fullPlaceholder}' or an existing update area.\n\nPlease click your cursor inside the document where you want this text to go, then try again.`, 
    DocumentApp.getUi().ButtonSet.OK
  );
}

/***************************************************************************************************
 * ALIAS: Ensures legacy calls to processPlaceholder automatically use the new ZWS logic.
 * **************************************************************************************************
 */
function processPlaceholder(placeholder, newContent) {
  return processPlaceholderSpaceBoundaries(placeholder, newContent);
}

// ==============================================================================
// Private helper functions (hidden from library consumers)
// ==============================================================================

/**
 * Generates unique start and end zero-width markers for a given string.
 * Converts characters to an invisible binary-like sequence to guarantee uniqueness.
 */
function generateZwsBoundaries_(placeholder) {
  const uniqueHiddenId = placeholder.split('').map(function(c) {
    return c.charCodeAt(0).toString(2).split('').map(function(b) {
      return b === '1' ? '\u200B' : '\u200C';
    }).join('');
  }).join('\u200D'); 
  
  const startMarker = '\u200B\u200D' + uniqueHiddenId + '\u200D\u200B';
  const endMarker = '\u200C\u200D' + uniqueHiddenId + '\u200D\u200C';
  
  return { start: startMarker, end: endMarker };
}

function executeZwsUpdate_(doc, boundaries, newContent, url) {
  const body = doc.getBody();
  
  // Look for the unique start marker followed by anything (non-greedy), ending with the end marker.
  const regex = boundaries.start + ".*?" + boundaries.end;
  const found = body.findText(regex);
  
  if (!found) return false;

  const textElement = found.getElement().asText();
  const start = found.getStartOffset();
  const end = found.getEndOffsetInclusive();
  const insertStr = boundaries.start + newContent + boundaries.end;

  // 1. Capture exact formatting attributes of the placeholder
  const attrs = textElement.getAttributes(start);
  
  // 2. Insert new text
  textElement.insertText(start, insertStr);
  
  // 3. Force explicit formatting onto the new text (Defeats ZWS Roboto fallback)
  textElement.setAttributes(start, start + insertStr.length - 1, attrs);
  
  // 4. Apply URL if one was provided
  if (url) {
    textElement.setLinkUrl(start, start + insertStr.length - 1, url);
  }
  
  // 5. Delete old text
  textElement.deleteText(start + insertStr.length, end + insertStr.length);
  
  return true;
}

function executeZwsPlaceholderReplacement_(doc, fullPlaceholder, boundaries, newContent, url) {
  const body = doc.getBody();
  const escapedPlaceholder = fullPlaceholder.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const found = body.findText(escapedPlaceholder);
  
  if (!found) return false;

  const textElement = found.getElement().asText();
  const start = found.getStartOffset();
  const end = found.getEndOffsetInclusive();
  const insertStr = boundaries.start + newContent + boundaries.end;

  // 1. Capture exact formatting attributes of the placeholder
  const attrs = textElement.getAttributes(start);
  
  // 2. Insert new text
  textElement.insertText(start, insertStr);
  
  // 3. Force explicit formatting onto the new text (Defeats ZWS Roboto fallback)
  textElement.setAttributes(start, start + insertStr.length - 1, attrs);
  
  // 4. Apply URL if one was provided
  if (url) {
    textElement.setLinkUrl(start, start + insertStr.length - 1, url);
  }
  
  // 5. Delete old text
  textElement.deleteText(start + insertStr.length, end + insertStr.length);
  
  return true;
}

function executeZwsCursorInsertion_(doc, boundaries, newContent, url) {
  const cursor = doc.getCursor();
  if (!cursor) return false;

  // Capture the start offset based on where the cursor currently sits
  const isText = cursor.getElement().getType() === DocumentApp.ElementType.TEXT;
  const startOffset = isText ? cursor.getSurroundingTextOffset() : 0;

  const insertStr = boundaries.start + newContent + boundaries.end;
  const textElement = cursor.insertText(insertStr);
  
  if (!textElement) return false;

  // Apply URL if one was provided
  if (url) {
    const start = isText ? startOffset : 0;
    textElement.setLinkUrl(start, start + insertStr.length - 1, url);
  }

  return true;
}

/**
 * Helper utility to convert standard hex color strings (e.g. "#fefff4")
 * into the decimal RGB object required by the Advanced Google Docs API.
 */
function hexToRgb(hex, defaultHex) {
  hex = (hex || defaultHex || "ffffff").replace(/^#/, '');
  var bigint = parseInt(hex, 16);
  return {
    red: ((bigint >> 16) & 255) / 255,
    green: ((bigint >> 8) & 255) / 255,
    blue: (bigint & 255) / 255
  };
}

/**************************************************************************************************
 * Retrieves metadata from the Document Body based on the Zero-Width Space Boundaries.
 * 
 * @param {string} placeholder - The raw key name (e.g., "ITEM_TITLE")
 * @returns {string|null} The extracted text, "UNFILLED_PLACEHOLDER" if empty, or null if missing.
 */
function getPlaceholderFromBody(placeholder) {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();

  const boundaries = generateZwsBoundaries_(placeholder);
  const cleanPlaceholder = placeholder.replace(/^{{|}}$/g, '');
  const fullPlaceholder = "{{" + cleanPlaceholder + "}}";
  
  // 1. Look for the raw {{placeholder}} FIRST
  const escapedPlaceholder = fullPlaceholder.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const foundRaw = body.findText(escapedPlaceholder);
  if (foundRaw) {
    return "UNFILLED_PLACEHOLDER"; 
  }

  // 2. Look for bounded text
  const regex = boundaries.start + ".*?" + boundaries.end;
  const found = body.findText(regex);
  if (found) {
    const textElement = found.getElement().asText();
    const matchText = textElement.getText().substring(
      found.getStartOffset(), 
      found.getEndOffsetInclusive() + 1
    );
    // Strip the invisible start and end boundaries, returning only the content
    const content = matchText.substring(boundaries.start.length, matchText.length - boundaries.end.length);
    return content;
  }

  return null;
}

/***************************************************************************************************
 * Injects a 1x1 transparent image with metadata in its Alt-Text description.
 * Actively cleans up any ghost images in all sections and forcefully injects into the BODY.
 * 
 * @param {Object} propertiesObj - The key-value pairs to store.
 **************************************************************************************************/
function injectMetadataImage(propertiesObj) {
  const doc = DocumentApp.getActiveDocument();
  
  // 1. Clean up existing metadata images EVERYWHERE (Limited to first 5 images per section for speed)
  const sections = [doc.getBody(), doc.getHeader(), doc.getFooter()];
  for (let s = 0; s < sections.length; s++) {
    if (sections[s]) {
      const images = sections[s].getImages();
      const limit = Math.min(5, images.length);
      for (let i = 0; i < limit; i++) {
        if (images[i].getAltTitle() === "DocKit_Meta") {
          images[i].removeFromParent();
        }
      }
    }
  }
  
  // 2. Programmatically generate a 1x1 transparent PNG using Base64
  const b64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAZdEVYdFNvZnR3YXJlAFBhaW50Lk5FVCB2My41LjbQg61aAAAADUlEQVQYV2P4//8/AwAI/AL+XvdLkQAAAABJRU5ErkJggg==";
  const blob = Utilities.newBlob(Utilities.base64Decode(b64), "image/png", "meta.png");
  
  // 3. Always inject at the very beginning of the BODY
  const body = doc.getBody();
  let firstChild = body.getNumChildren() > 0 ? body.getChild(0) : null;
  let inlineImg = null;

  if (firstChild && firstChild.getType() === DocumentApp.ElementType.PARAGRAPH) {
    inlineImg = firstChild.asParagraph().insertInlineImage(0, blob);
  } else {
    const paragraph = body.insertParagraph(0, "");
    inlineImg = paragraph.insertInlineImage(0, blob);
  }
  
  // 4. Set Alt-Text metadata
  inlineImg.setAltTitle("DocKit_Meta");
  inlineImg.setAltDescription(JSON.stringify(propertiesObj));
  
  console.log("Injected DocKit_Meta image safely into BODY");
}

/***************************************************************************************************
 * Reads metadata strictly from the 1x1 image in the document BODY.
 * Limited to scanning the first 5 images to preserve maximum performance.
 * 
 * @returns {Object|null} The parsed metadata object.
 **************************************************************************************************/
function readMetadataFromImage() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  
  if (!body) return null;
  
  const images = body.getImages();
  const limit = Math.min(5, images.length);
  
  for (let i = 0; i < limit; i++) {
    if (images[i].getAltTitle() === "DocKit_Meta") {
      try {
        const desc = images[i].getAltDescription();
        return JSON.parse(desc);
      } catch (e) {
        console.log("Failed to parse JSON from DocKit_Meta image: " + e.message);
        return null;
      }
    }
  }
  return null;
}