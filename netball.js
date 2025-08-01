const DATA_START_ROW = 4; // Data starts from row 3
const GAMETYPE = 'netball'; 
const HUBSPOT_SERVERLESS_ENDPOINT =`https://www.activesgcircle.gov.sg/_hcms/api/sport-handler-api?gametype=${GAMETYPE}`;
const HUBDB_ROW_ID_COLUMN = 1; // Column A for storing HubDB row ID (hs_id)
const SYNC_STATUS_ID_COLUMN = 11; // Column M for storing sync status
const SYNC_MESSAGE_ID_COLUMN = 12; // Column N for storing sync message 

// TARGET SHEET CONFIGURATION - Change this to switch sheets
const TARGET_SHEET_NAME = 'Sheet1';
function getTargetSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(TARGET_SHEET_NAME);
  
  if (!sheet) {
    throw new Error(`Sheet "${TARGET_SHEET_NAME}" not found. Available sheets: ${spreadsheet.getSheets().map(s => s.getName()).join(', ')}`);
  }
  
  return sheet;
}

/**
 * Clear sync status and message columns for specific rows
 */
function clearSyncStatusForRows(sheet, rowNumbers) {
  try {
    if (!rowNumbers || rowNumbers.length === 0) return;
    
    console.log(`Clearing sync status for rows: ${rowNumbers.join(', ')}`);
    
    // Clear sync status and message for each row
    rowNumbers.forEach(rowNumber => {
      sheet.getRange(rowNumber, SYNC_STATUS_ID_COLUMN).setValue('');
      sheet.getRange(rowNumber, SYNC_MESSAGE_ID_COLUMN).setValue('');
    });
    
    console.log('Sync status cleared for specified rows');
  } catch (error) {
    console.error('Error clearing sync status:', error);
  }
}

/**
 * Clear sync status and message columns for all data rows
 */
function clearAllSyncStatus(sheet, lastRow) {
  try {
    if (lastRow <= DATA_START_ROW) return;
    
    console.log(`Clearing sync status for all rows from ${DATA_START_ROW} to ${lastRow}`);

    console.log('All sync status cleared');
  } catch (error) {
    console.error('Error clearing all sync status:', error);
  }
}

/**
 * Main function to get all data and detect changes
 * READ-ONLY MODE: Only reads from sheet and updates hs_id column
 */
function syncAllData() {
  try {
    const sheet = getTargetSheet(); // Use target sheet instead of active sheet
    const lastRow = sheet.getLastRow();
    
    console.log('=== STARTING GAME SYNC (READ-ONLY MODE) ===');
    console.log(`Target sheet: ${TARGET_SHEET_NAME}`);
    console.log(`Processing rows ${DATA_START_ROW} to ${lastRow}`);
    
    // Clear all sync status before starting
    clearAllSyncStatus(sheet, lastRow);
    
    // Get all data and cleared rows
    const { allData, clearedRows } = getAllSheetData(sheet, lastRow);
    
    // Handle cleared rows first (delete from HubDB)
    if (clearedRows.length > 0) {
      console.log(`Processing ${clearedRows.length} cleared rows for deletion`);
      handleClearedRows(clearedRows, sheet);
    }
    
    // Get stored data (previous state)
    const storedData = getStoredData();
    
    // Compare and find changes
    const changes = detectChanges(allData, storedData);
    
    // Check for rows that need HubDB row creation (only from changed/new rows)
    const rowsNeedingHubDBCreation = findRowsNeedingHubDBCreation(allData, sheet, changes);
    
    if (rowsNeedingHubDBCreation.length > 0) {
      console.log(`Found ${rowsNeedingHubDBCreation.length} rows needing HubDB creation`);
      createHubDBRows(rowsNeedingHubDBCreation, sheet);
      
      // Refresh data after creating HubDB rows to get the updated row IDs
      const { allData: updatedData } = getAllSheetData(sheet, lastRow);
      storeData(updatedData);
    }
    
    if (changes.length > 0) {
      console.log(`Found ${changes.length} changes`);
      
      // Filter changes to only include those with HubDB row IDs
      const changesWithHubDBId = changes.filter(change => {
        const hasHubDBId = (change.type === 'DELETED' && change.oldData?.hubdbRowId) || 
                          (change.type !== 'DELETED' && change.newData?.hubdbRowId);
        
        if (!hasHubDBId) {
          console.log(`⏭️ Skipping ${change.type} operation for ${change.uniqueId} - No HubDB row ID`);
        }
        return hasHubDBId;
      });
      
      if (changesWithHubDBId.length > 0) {
        console.log(`Sending ${changesWithHubDBId.length} changes with HubDB IDs to HubSpot`);
        sendToHubSpot(changesWithHubDBId, sheet);
      } else {
        console.log('No changes with HubDB row IDs to send to HubSpot');
      }
      
      // Store the new data state after processing changes
      if (rowsNeedingHubDBCreation.length === 0) {
        storeData(allData);
      }
      
      console.log('=== SYNC COMPLETE ===');
    } else {
      console.log('No changes detected.');
    }
    
  } catch (error) {
    console.error('Error in syncAllData:', error);
  }
}

/**
 * Get all data from the sheet including HubDB row ID
 * READ-ONLY: Only reads data from sheet
 */
function getAllSheetData(sheet, lastRow) {
  // Get HubDB row ID from column A (hs_id)
  const hubdbIdRange = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 1);
  const hubdbIds = hubdbIdRange.getValues();
  
  // Get main data from columns B to L (11 columns)
  const dataRange = sheet.getRange(DATA_START_ROW, 2, lastRow - DATA_START_ROW + 1, 11);
  const values = dataRange.getValues();
  
  const allData = {};
  const clearedRows = [];
  
  values.forEach((row, index) => {
    const actualRow = DATA_START_ROW + index;
    const hubdbRowId = hubdbIds[index][0] || '';
    
    // Check if the entire row is empty (including hs_id)
    const isCompletelyEmpty = !row[0] && !row[1] && !row[2] && !row[3] && 
                             !row[4] && !row[5] && !row[6] && !row[7] && !row[8] && 
                             !row[9] && !row[10] && !hubdbRowId;
    
    // Check if only data is cleared but hs_id exists
    const isDataClearedButHasId = (!row[0] && !row[1] && !row[2] && !row[3] && 
                                  !row[4] && !row[5] && !row[6] && !row[7] && !row[8] && 
                                  !row[9] && !row[10]) && hubdbRowId;
    
    // Track rows cleared but still have HubDB ID
    if (isDataClearedButHasId) {
      clearedRows.push({
        sheetRow: actualRow,
        hubdbRowId: hubdbRowId
      });
      return;
    }
    
    // Skip completely empty rows (including those with no hs_id)
    if (isCompletelyEmpty) return;
    
    // Skip rows that don't have core identifying fields
    if (!row[0] && !row[1] && !row[2] && !row[3] && !row[4]) return;

    const data = {
          game_no: row[0] || '',
          date_and_time: row[1] || '',
          category: row[2] || '',
          sub_category: row[3] || '',
          round: row[4] || '',
          team_1: row[5] || '',
          team_1_score: row[6] || '',
          team_2_score: row[7] || '',
          team_2: row[8] || '',
          hubdbRowId: hubdbRowId,
          sheetRow: actualRow
    };
    
    
    const uniqueId = hubdbRowId || `temp_${actualRow}`;
    allData[uniqueId] = data;
  });
  
  console.log(`Extracted ${Object.keys(allData).length} records from sheet`);
  
  return { allData, clearedRows };
}

/**
 * Find rows that need HubDB row creation
 */
function findRowsNeedingHubDBCreation(allData, sheet, changes) {
  const rowsNeedingCreation = [];
  const changedUniqueIds = new Set();
  
  if (changes && changes.length > 0) {
    changes.forEach(change => {
      if (change.type === 'NEW' || change.type === 'UPDATED') {
        changedUniqueIds.add(change.uniqueId);
      }
    });
  } else {
    Object.keys(allData).forEach(uniqueId => {
      changedUniqueIds.add(uniqueId);
    });
  }
  changedUniqueIds.forEach(uniqueId => {
    const data = allData[uniqueId];
    if (!data) return;
    const hasRequiredData = data.date_and_time && data.team_1 && data.team_2;
    const missingHubDBRowId = !data.hubdbRowId;
    if (hasRequiredData && missingHubDBRowId) {
      rowsNeedingCreation.push({
        uniqueId: uniqueId,
        data: data,
        sheetRow: data.sheetRow
      });
    }
  });
  
  return rowsNeedingCreation;
}

/**
 * Create HubDB rows for rows that need them
 */
function createHubDBRows(rowsNeedingCreation, sheet) {
  try {
    console.log('Creating HubDB rows...');
    
    // Clear sync status for rows that will be processed
    const rowNumbers = rowsNeedingCreation.map(row => row.sheetRow);
    clearSyncStatusForRows(sheet, rowNumbers);
    
    for (const rowInfo of rowsNeedingCreation) {  // Changed to for-of loop
      console.log('Creating HubDB rowInfo...', rowInfo);
      
      // Set status to "processing" before API call
      sheet.getRange(rowInfo.sheetRow, SYNC_STATUS_ID_COLUMN).setValue("syncing");
     
      const payload = {
        operation: 'CREATE_HUBDB_ROW',
        uniqueId: rowInfo.uniqueId,
        data: {
            game_no: rowInfo.data.game_no,
            date_and_time: rowInfo.data.date_and_time,
            category: rowInfo.data.category,
            sub_category: rowInfo.data.sub_category,
            round: rowInfo.data.round,
            team_1: rowInfo.data.team_1,
            team_1_score: rowInfo.data.team_1_score,
            team_2_score: rowInfo.data.team_2_score,
            team_2: rowInfo.data.team_2,
        },
        metadata: {
          timestamp: new Date().toISOString(),
          source: 'google_sheets_sync',
          sheetRow: rowInfo.sheetRow
        }
      };
      
      try{ 
        const response = UrlFetchApp.fetch(`${HUBSPOT_SERVERLESS_ENDPOINT}`, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json'
          },
          payload: JSON.stringify(payload)
        });
        const responseCode = response.getResponseCode();
        const responseText = response.getContentText();
        
        if (responseCode >= 200 && responseCode < 300) {
          const responseData = JSON.parse(responseText);
          
          if (responseData.success && responseData.hubdbRowId) {
            sheet.getRange(rowInfo.sheetRow, HUBDB_ROW_ID_COLUMN).setValue(responseData.hubdbRowId);
            sheet.getRange(rowInfo.sheetRow, SYNC_STATUS_ID_COLUMN).setValue("sync success");
            sheet.getRange(rowInfo.sheetRow, SYNC_MESSAGE_ID_COLUMN).setValue("HubDB row created successfully");
            
            console.log(`Created HubDB row ${responseData.hubdbRowId} for sheet row ${rowInfo.sheetRow}`);
          } else {
            console.warn(`HubDB row creation failed for ${rowInfo.uniqueId}`);
            sheet.getRange(rowInfo.sheetRow, SYNC_STATUS_ID_COLUMN).setValue("error");
            sheet.getRange(rowInfo.sheetRow, SYNC_MESSAGE_ID_COLUMN).setValue(`Creation failed: ${responseData.message || 'Unknown error'}`);
          }
        } else {
          sheet.getRange(rowInfo.sheetRow, SYNC_STATUS_ID_COLUMN).setValue("error");
          sheet.getRange(rowInfo.sheetRow, SYNC_MESSAGE_ID_COLUMN).setValue(`HTTP ${responseCode}: ${responseText}`);
          throw new Error(`HubDB creation failed: ${responseCode} - ${responseText}`);
        }
      }
      catch(error){
        console.log("error 247 ", error.message);
        sheet.getRange(rowInfo.sheetRow, SYNC_STATUS_ID_COLUMN).setValue("error");
        sheet.getRange(rowInfo.sheetRow, SYNC_MESSAGE_ID_COLUMN).setValue(`Exception: ${error.message}`);
      }
      
      Utilities.sleep(100);
    }
    
  } catch (error) {
    console.error('Error creating HubDB rows:', error);
    throw error;
  }
}

/**
 * Handle cleared rows - delete from HubDB
 */
function handleClearedRows(clearedRows, sheet) {
  try {
    if (clearedRows.length === 0) return;

    // Clear sync status for cleared rows before processing
    const rowNumbers = clearedRows.map(row => row.sheetRow);
    clearSyncStatusForRows(sheet, rowNumbers);

    if (clearedRows.length > 1) {
      handleBatchClearedRows(clearedRows, sheet);
    } else {
      handleSingleClearedRow(clearedRows[0], sheet);
    }
    
  } catch (error) {
    console.error('Error handling cleared rows:', error);
    throw error;
  }
}

/**
 * Handle batch delete for multiple cleared rows
 */
function handleBatchClearedRows(clearedRows, sheet) {
  try {
    const hubdbRowIds = clearedRows.map(row => row.hubdbRowId).filter(id => id);
    
    if (hubdbRowIds.length === 0) return;

    // Set processing status for all rows
    clearedRows.forEach(clearedRow => {
      sheet.getRange(clearedRow.sheetRow, SYNC_STATUS_ID_COLUMN).setValue("syncing");
    });

    const payload = {
      operation: 'BATCH_DELETE_HUBDB_ROWS',
      hubdbRowIds: hubdbRowIds,
      metadata: {
        timestamp: new Date().toISOString(),
        source: 'google_sheets_sync',
        reason: 'rows_cleared',
        totalRows: clearedRows.length
      }
    };
    
    const response = UrlFetchApp.fetch(`${HUBSPOT_SERVERLESS_ENDPOINT}`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload)
    });
    
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    console.log("402 ", responseText);
    
    if (responseCode >= 200 && responseCode < 300) {
      const responseData = JSON.parse(responseText);
      console.log("405 ", responseData);
      
      if (responseData.success) {
        const successfulDeletions = responseData.results?.filter(r => r.status === 'success') || [];
        
        // Update status for each cleared row
        clearedRows.forEach(clearedRow => {
          const wasDeleted = successfulDeletions.some(result => 
            result.hubdbId === clearedRow.hubdbRowId
          );
          
          if (wasDeleted) {
            // Clear hs_id and update sync status for successful deletions
            sheet.getRange(clearedRow.sheetRow, HUBDB_ROW_ID_COLUMN).setValue('');
            sheet.getRange(clearedRow.sheetRow, SYNC_STATUS_ID_COLUMN).setValue("deleted");
            sheet.getRange(clearedRow.sheetRow, SYNC_MESSAGE_ID_COLUMN).setValue("Successfully deleted from HubDB");
            console.log(`Batch deleted HubDB row ${clearedRow.hubdbRowId}`);
          } else {
            sheet.getRange(clearedRow.sheetRow, SYNC_STATUS_ID_COLUMN).setValue("error");
            sheet.getRange(clearedRow.sheetRow, SYNC_MESSAGE_ID_COLUMN).setValue("Batch deletion failed");
          }
        });
      } else {
        // If batch operation failed, update all rows with error status
        clearedRows.forEach(clearedRow => {
          sheet.getRange(clearedRow.sheetRow, SYNC_STATUS_ID_COLUMN).setValue("error");
          sheet.getRange(clearedRow.sheetRow, SYNC_MESSAGE_ID_COLUMN).setValue(`Batch deletion failed: ${responseData.message || 'Unknown error'}`);
        });
        
        // Fall back to individual deletions
        clearedRows.forEach(clearedRow => {
          handleSingleClearedRow(clearedRow, sheet);
        });
      }
    } else {
      // HTTP error - update all rows with error status and fall back to individual deletions
      clearedRows.forEach(clearedRow => {
        sheet.getRange(clearedRow.sheetRow, SYNC_STATUS_ID_COLUMN).setValue("error");
        sheet.getRange(clearedRow.sheetRow, SYNC_MESSAGE_ID_COLUMN).setValue(`HTTP ${responseCode}: ${responseText}`);
      });
      
      // Fall back to individual deletions
      clearedRows.forEach(clearedRow => {
        handleSingleClearedRow(clearedRow, sheet);
      });
    }
    
  } catch (error) {
    console.error('Error in batch delete, falling back to individual deletions:', error);
    
    // Update all rows with error status
    clearedRows.forEach(clearedRow => {
      sheet.getRange(clearedRow.sheetRow, SYNC_STATUS_ID_COLUMN).setValue("error");
      sheet.getRange(clearedRow.sheetRow, SYNC_MESSAGE_ID_COLUMN).setValue(`Exception: ${error.message}`);
    });
    
    // Fall back to individual deletions
    clearedRows.forEach(clearedRow => {
      try {
        handleSingleClearedRow(clearedRow, sheet);
      } catch (individualError) {
        console.error(`Failed to delete individual row ${clearedRow.hubdbRowId}:`, individualError);
      }
    });
  }
}

/**
 * Handle single cleared row deletion
 */
function handleSingleClearedRow(clearedRow, sheet) {
  try {
    // Set processing status
    sheet.getRange(clearedRow.sheetRow, SYNC_STATUS_ID_COLUMN).setValue("syncing");
    
    const payload = {
      operation: 'DELETE_HUBDB_ROW',
      hubdbRowId: clearedRow.hubdbRowId,
      sheetRow: clearedRow.sheetRow,
      metadata: {
        timestamp: new Date().toISOString(),
        source: 'google_sheets_sync',
        reason: 'row_cleared'
      }
    };
    
    const response = UrlFetchApp.fetch(`${HUBSPOT_SERVERLESS_ENDPOINT}`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload)
    });
    
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode >= 200 && responseCode < 300) {
      const responseData = JSON.parse(responseText);
      
      if (responseData.success) {
        sheet.getRange(clearedRow.sheetRow, HUBDB_ROW_ID_COLUMN).setValue('');
        sheet.getRange(clearedRow.sheetRow, SYNC_STATUS_ID_COLUMN).setValue("deleted");
        sheet.getRange(clearedRow.sheetRow, SYNC_MESSAGE_ID_COLUMN).setValue("Successfully deleted from HubDB");
        console.log(`Deleted HubDB row ${clearedRow.hubdbRowId}`);
      } else {
        sheet.getRange(clearedRow.sheetRow, SYNC_STATUS_ID_COLUMN).setValue("error");
        sheet.getRange(clearedRow.sheetRow, SYNC_MESSAGE_ID_COLUMN).setValue(`Deletion failed: ${responseData.message || 'Unknown error'}`);
      }
    } else {
      sheet.getRange(clearedRow.sheetRow, SYNC_STATUS_ID_COLUMN).setValue("error");
      sheet.getRange(clearedRow.sheetRow, SYNC_MESSAGE_ID_COLUMN).setValue(`HTTP ${responseCode}: ${responseText}`);
      console.error(`Failed to delete HubDB row ${clearedRow.hubdbRowId}: ${responseCode} - ${responseText}`);
    }
    
    Utilities.sleep(100);
    
  } catch (error) {
    console.error(`Error deleting single HubDB row ${clearedRow.hubdbRowId}:`, error);
    sheet.getRange(clearedRow.sheetRow, SYNC_STATUS_ID_COLUMN).setValue("error");
    sheet.getRange(clearedRow.sheetRow, SYNC_MESSAGE_ID_COLUMN).setValue(`Exception: ${error.message}`);
  }
}

/**
 * Detect changes between current and stored data
 * Enhanced to detect deleted rows using stored data
 */
function detectChanges(currentData, storedData) {
  const changes = [];
  
  // Check each record in current data
  Object.keys(currentData).forEach(uniqueId => {
    const current = currentData[uniqueId];
    const stored = storedData[uniqueId];
    
    if (!stored) {
      // New record
      if (current.hubdbRowId || uniqueId.startsWith('temp_')) {
        changes.push({
          type: 'NEW',
          uniqueId: uniqueId,
          row: current.sheetRow,
          newData: current,
          changedFields: ['ALL']
        });
      }
    } else {
      // Check for changes in existing record
      const changedFields = [];
      const fieldsToCheck = ["game_no" ,"date_and_time", "category", "sub_category", "round", "team_1", "team_1_score", "team_2_score", "team_2"];
      fieldsToCheck.forEach(field => {
        if (current[field] !== stored[field]) {
          changedFields.push({
            field: field,
            oldValue: stored[field],
            newValue: current[field]
          });
        }
      });
      
      if (changedFields.length > 0) {
        changes.push({
          type: 'UPDATED',
          uniqueId: uniqueId,
          row: current.sheetRow,
          newData: current,
          oldData: stored,
          changedFields: changedFields
        });
      }
    }
  });
  
  // Check for deleted records using stored data
  Object.keys(storedData).forEach(uniqueId => {
    if (!currentData[uniqueId]) {
      const storedRow = storedData[uniqueId];
      
      // Only process deletion if the stored record had a HubDB row ID
      if (storedRow.hubdbRowId) {
        changes.push({
          type: 'DELETED',
          uniqueId: uniqueId,
          oldData: storedRow,
          changedFields: ['DELETED']
        });
        console.log(`🗑️ DELETED record detected: ${uniqueId} (HubDB ID: ${storedRow.hubdbRowId})`);
      } else {
        console.log(`⏭️ Skipping deletion of ${uniqueId} - no HubDB ID in stored data`);
      }
    }
  });
  
  return changes;
}


// Create a generic HTTP request function
function makeHubSpotRequest(payload, requestDescription = '') {
  if (requestDescription) {
    console.log(`Making HubSpot request: ${requestDescription}`);
  }
  
  return UrlFetchApp.fetch(`${HUBSPOT_SERVERLESS_ENDPOINT}`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload)
  });
}

function sendToHubSpot(changes, sheet) {
  try {
    console.log('Sending to HubSpot...');
    
    clearSyncStatusForProcessedRows(changes, sheet);
    
    const { deleteChanges, otherChanges } = separateChanges(changes);
    
    if (deleteChanges.length > 1) {
      handleBatchDeletes(deleteChanges, otherChanges, sheet);
    } else if (deleteChanges.length === 1) {
      otherChanges.push(...deleteChanges);
    }
    
    if (otherChanges.length > 0) {
      return handleIndividualOperations(otherChanges, sheet);
    }
    
    return { message: 'All operations completed' };
    
  } catch (error) {
    console.error('Error sending to HubSpot:', error);
    throw error;
  }
}

function clearSyncStatusForProcessedRows(changes, sheet) {
  const rowsToProcess = changes
    .filter(change => change.row)
    .map(change => change.row);
  
  if (rowsToProcess.length > 0) {
    clearSyncStatusForRows(sheet, rowsToProcess);
  }
}

function separateChanges(changes) {
  return {
    deleteChanges: changes.filter(change => change.type === 'DELETED'),
    otherChanges: changes.filter(change => change.type !== 'DELETED')
  };
}

function handleBatchDeletes(deleteChanges, otherChanges, sheet) {
  const batchDeletePayload = createBatchDeletePayload(deleteChanges);
  
  try {
    setSyncingStatusForDeletes(deleteChanges, sheet);
    
    const batchResponse = sendBatchDeleteRequest(batchDeletePayload);
    const { responseCode, responseData } = processBatchResponse(batchResponse);
    
    if (isSuccessResponse(responseCode)) {
      updateBatchDeleteStatuses(deleteChanges, responseData, sheet);
    } else {
      handleBatchDeleteFailure(deleteChanges, otherChanges, responseCode, batchResponse.getContentText(), sheet);
    }
    
  } catch (batchError) {
    handleBatchDeleteException(deleteChanges, otherChanges, batchError, sheet);
  }
}

function createBatchDeletePayload(deleteChanges) {
  return {
    operation: 'BATCH_DELETE_HUBDB_ROWS',
    hubdbRowIds: deleteChanges.map(change => change.oldData?.hubdbRowId).filter(id => id),
    metadata: {
      timestamp: new Date().toISOString(),
      source: 'google_sheets_sync_batch',
      reason: 'data_changes',
      totalChanges: deleteChanges.length
    }
  };
}

function setSyncingStatusForDeletes(deleteChanges, sheet) {
  deleteChanges.forEach(change => {
    if (change.oldData?.sheetRow) {
      sheet.getRange(change.oldData.sheetRow, SYNC_STATUS_ID_COLUMN).setValue('syncing');
    }
  });
}

function sendBatchDeleteRequest(payload) {
  return makeHubSpotRequest(payload, 'batch delete operation');
}

function processBatchResponse(response) {
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();
  
  return {
    responseCode,
    responseData: isSuccessResponse(responseCode) ? JSON.parse(responseText) : null,
    responseText
  };
}

function isSuccessResponse(responseCode) {
  return responseCode >= 200 && responseCode < 300;
}

function updateBatchDeleteStatuses(deleteChanges, responseData, sheet) {
  console.log('Successfully batch deleted from HubSpot');
  
  if (responseData.success) {
    const successfulDeletions = responseData.results?.filter(r => r.status === 'success') || [];
    
    deleteChanges.forEach(change => {
      if (change.oldData?.sheetRow) {
        const wasDeleted = successfulDeletions.some(result => 
          result.hubdbId === change.oldData.hubdbRowId
        );
        
        updateDeleteStatus(change.oldData.sheetRow, wasDeleted, sheet);
      }
    });
  } else {
    updateAllDeletesWithError(deleteChanges, `Batch deletion failed: ${responseData.message || 'Unknown error'}`, sheet);
  }
}

function updateDeleteStatus(row, wasDeleted, sheet) {
  if (wasDeleted) {
    sheet.getRange(row, HUBDB_ROW_ID_COLUMN).setValue('');
    sheet.getRange(row, SYNC_STATUS_ID_COLUMN).setValue("deleted");
    sheet.getRange(row, SYNC_MESSAGE_ID_COLUMN).setValue("Successfully deleted from HubDB");
  } else {
    sheet.getRange(row, SYNC_STATUS_ID_COLUMN).setValue("error");
    sheet.getRange(row, SYNC_MESSAGE_ID_COLUMN).setValue("Batch deletion failed");
  }
}

function updateAllDeletesWithError(deleteChanges, errorMessage, sheet) {
  deleteChanges.forEach(change => {
    if (change.oldData?.sheetRow) {
      sheet.getRange(change.oldData.sheetRow, SYNC_STATUS_ID_COLUMN).setValue("error");
      sheet.getRange(change.oldData.sheetRow, SYNC_MESSAGE_ID_COLUMN).setValue(errorMessage);
    }
  });
}

function handleBatchDeleteFailure(deleteChanges, otherChanges, responseCode, responseText, sheet) {
  console.warn('Batch delete failed, handling individually');
  updateAllDeletesWithError(deleteChanges, `HTTP ${responseCode}: ${responseText}`, sheet);
  otherChanges.push(...deleteChanges);
}

function handleBatchDeleteException(deleteChanges, otherChanges, error, sheet) {
  console.error('Batch delete error, handling individually:', error);
  updateAllDeletesWithError(deleteChanges, `Exception: ${error.message}`, sheet);
  otherChanges.push(...deleteChanges);
}

function handleIndividualOperations(otherChanges, sheet) {
  setSyncingStatusForOperations(otherChanges, sheet);
  
  const payload = createIndividualOperationsPayload(otherChanges);
  const response = sendIndividualOperationsRequest(payload);
  
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();
  
  if (isSuccessResponse(responseCode)) {
    updateIndividualOperationStatuses(otherChanges, sheet, true);
    console.log('Successfully sent operations to HubSpot');
    return JSON.parse(responseText);
  } else {
    updateIndividualOperationStatuses(otherChanges, sheet, false, responseCode, responseText);
    throw new Error(`HubSpot API error: ${responseCode} - ${responseText}`);
  }
}

function setSyncingStatusForOperations(otherChanges, sheet) {
  otherChanges.forEach(change => {
    const targetRow = getTargetRow(change);
    if (targetRow) {
      sheet.getRange(targetRow, SYNC_STATUS_ID_COLUMN).setValue("syncing");
    }
  });
}

function getTargetRow(change) {
  return change.row || (change.oldData ? change.oldData.sheetRow : null);
}

function createIndividualOperationsPayload(otherChanges) {
  return {
    gameData: otherChanges.map(change => ({
      operation: change.type,
      uniqueId: change.uniqueId,
      data: change.newData,
      oldData: change.oldData,
      changedFields: change.changedFields,
      metadata: {
        timestamp: new Date().toISOString(),
        source: 'google_sheets_sync',
        sheetRow: getTargetRow(change)
      }
    }))
  };
}

function sendIndividualOperationsRequest(payload) {
  return makeHubSpotRequest(payload, 'individual operations');
}

function updateIndividualOperationStatuses(otherChanges, sheet, isSuccess, responseCode = null, responseText = null) {
  otherChanges.forEach(change => {
    const targetRow = getTargetRow(change);
    if (!targetRow) return;
    
    if (isSuccess) {
      updateSuccessfulOperation(change, targetRow, sheet);
    } else {
      sheet.getRange(targetRow, SYNC_STATUS_ID_COLUMN).setValue("error");
      sheet.getRange(targetRow, SYNC_MESSAGE_ID_COLUMN).setValue(`HTTP ${responseCode}: ${responseText}`);
    }
  });
}

function updateSuccessfulOperation(change, targetRow, sheet) {
  if (change.type === 'DELETED') {
    sheet.getRange(targetRow, HUBDB_ROW_ID_COLUMN).setValue('');
    sheet.getRange(targetRow, SYNC_STATUS_ID_COLUMN).setValue("deleted");
    sheet.getRange(targetRow, SYNC_MESSAGE_ID_COLUMN).setValue("Successfully deleted from HubDB");
  } else {
    sheet.getRange(targetRow, SYNC_STATUS_ID_COLUMN).setValue("sync success");
    sheet.getRange(targetRow, SYNC_MESSAGE_ID_COLUMN).setValue(`${change.type} operation completed successfully`);
  }
}


/**
 * Get stored data from Script Properties
 */
function getStoredData() {
  try {
    const stored = PropertiesService.getScriptProperties().getProperty('GAME_DATA');
    return stored ? JSON.parse(stored) : {};
  } catch (error) {
    console.log('No stored data found (first run)');
    return error;
  }
}

/**
 * Store current data state
 */
function storeData(data) {
  PropertiesService.getScriptProperties().setProperty('GAME_DATA', JSON.stringify(data));
  console.log('Data state stored');
}

/**
 * Set up triggers for automatic syncing
 * Uses onChange trigger to catch row deletions
 */
function setupTriggers() {
  try {
    // Delete existing triggers
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'syncAllData' || 
          trigger.getHandlerFunction() === 'onSheetChange') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    // Create onChange trigger (catches row deletions and other structural changes)
    ScriptApp.newTrigger('onSheetChange')
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onChange()
      .create();
    
    // Create onEdit trigger (catches cell value changes)  
    ScriptApp.newTrigger('syncAllData')
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onEdit()
      .create();
    
    // Optional: Time-based trigger for regular sync (every 5 minutes)
    ScriptApp.newTrigger('syncAllData')
      .timeBased()
      .everyMinutes(5)
      .create();
    
    console.log('Triggers set up successfully');
    console.log('   - onChange trigger: Will sync when rows are deleted/added');
    console.log('   - onEdit trigger: Will sync when cells are edited');
    console.log('   - Time-based trigger: Will sync every 5 minutes');
    
  } catch (error) {
    console.error('Error setting up triggers:', error);
  }
}

/**
 * Handle sheet structure changes (row deletions, insertions)
 * This function is called by the onChange trigger
 */
function onSheetChange(e) {
  try {
    console.log('=== SHEET CHANGE DETECTED ===');
    console.log('Change type:', e.changeType);
    
    // Always run sync on any structural change
    // This will catch deleted rows through the stored data comparison
    syncAllData();
    
  } catch (error) {
    console.error('Error in onSheetChange:', error);
  }
}

/**
 * Remove all triggers
 */
function removeTriggers() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'syncAllData' || 
          trigger.getHandlerFunction() === 'onSheetChange') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    console.log('All sync triggers removed');
    
  } catch (error) {
    console.error('Error removing triggers:', error);
  }
}

/**
 * Manual sync trigger
 */
function manualSync() {
  console.log('=== MANUAL SYNC TRIGGERED ===');
  syncAllData();
}
