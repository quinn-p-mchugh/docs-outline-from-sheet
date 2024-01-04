/**
 * Main function to create an outline in a Google Doc from data in a Google Sheet.
 */
function createOutlineFromSheet() {
  const sheetId = '1qKFrS6_IYPAvMB_tLVTdvQU7M1A17dLVYrie7JNzyFM';  // Replace with your Google Sheet ID
  const data = readSheetData(sheetId);
  const groupedData = groupDataByOrganization(data);
  const doc = DocumentApp.create('Generate Docs Outline From Sheet');
  createOutline(doc.getBody(), groupedData);
}

/**
 * Reads data from the specified Google Sheet.
 * @param {string} sheetId - The ID of the Google Sheet.
 * @return {Array} The data from the sheet excluding the header row.
 */
function readSheetData(sheetId) {
  const sheet = SpreadsheetApp.openById(sheetId).getSheets()[0];
  // Get all data, skip header row
  return sheet.getDataRange().getValues().slice(1);
}

/**
 * Groups ticket data by organization name.
 * @param {Array} data - The array of ticket data.
 * @return {Object}P The data grouped by organization.
 */
function groupDataByOrganization(data) {
  const grouped = {};
  data.forEach(row => {
    const organizationName = row[10];  // Adjust index based on 'Organization Name' column
    if (!grouped[organizationName]) {
      grouped[organizationName] = [];
    }
    grouped[organizationName].push(row);
  });
  return grouped;
}

/**
 * Creates an outline in the document body with the provided data.
 * @param {Object} body - The body of the Google Doc.
 * @param {Object} groupedData - The ticket data grouped by organization.
 */
function createOutline(body, groupedData) {
  for (const organizationName in groupedData) {
    const orgListItem = body.appendListItem(organizationName);
    orgListItem.setBold(true);
    orgListItem.setGlyphType(DocumentApp.GlyphType.BULLET);

    groupedData[organizationName].forEach(row => {
      const ticketId = row[0];  // Assuming 'ID' is in the first column
      const ticketSummary = row[6];  // Update index based on 'Ticket Summary' column
      const ticketLink = `https://fortrobotics.zendesk.com/agent/tickets/${ticketId}`;
      const ticketText = `#${ticketId} - ${ticketSummary}`;
      
      const subListItem = body.appendListItem(ticketText).setBold(false);
      subListItem.setGlyphType(DocumentApp.GlyphType.HOLLOW_BULLET);
      formatSubBulletEntry(subListItem, ticketId, ticketLink);
    });
  }
}

/**
 * Formats a sub-bullet entry in the outline with a hyperlink.
 * @param {Object} listItem - The list item element to format.
 * @param {string} ticketId - The ID of the ticket.
 * @param {string} ticketLink - The hyperlink URL for the ticket.
 */
function formatSubBulletEntry(listItem, ticketId, ticketLink) {
  const text = listItem.getText();
  const textElement = listItem.editAsText();
  // Set hyperlink on ticket ID
  const ticketIdString = `#${ticketId}`;
  const ticketIdIndex = text.indexOf(ticketIdString);
  if (ticketIdIndex !== -1) {
    textElement.setLinkUrl(ticketIdIndex, ticketIdIndex + ticketIdString.length - 1, ticketLink);
  }
}

// Run this function to create the document
