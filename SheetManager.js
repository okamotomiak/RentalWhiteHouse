// SheetManager.gs - Sheet Management System
// Handles all sheet creation, formatting, and data structure

const SheetManager = {
  
  /**
   * Sheet Headers Configuration
   */
  HEADERS: {
    TENANTS: [
      'Room Number', 'Rental Price', 'Negotiated Price', 'Current Tenant Name',
      'Tenant Email', 'Tenant Phone', 'Move-In Date', 'Security Deposit Paid',
      'Room Status', 'Last Payment Date', 'Payment Status', 'Move-Out Date (Planned)',
      'Emergency Contact', 'Lease End Date', 'Notes'
    ],
    
    BUDGET: [
      'Date', 'Type', 'Description', 'Amount', 'Category', 
      'Payment Method', 'Reference Number', 'Tenant/Guest', 'Receipt'
    ],
    
    APPLICATIONS: [
      'Application ID', 'Timestamp', 'Full Name', 'Email', 'Phone',
      'Current Address', 'Desired Move-in Date', 'Preferred Room', 'Expected Stay',
      'Employment Status', 'Employer', 'Monthly Income', 'Reference 1', 'Reference 2',
      'Emergency Contact', 'Vehicle Info', 'About Yourself', 'Special Needs',
      'Application Status', 'Review Notes', 'Processed Date'
    ],
    
    MOVEOUTS: [
      'Request ID', 'Timestamp', 'Tenant Name', 'Email', 'Phone',
      'Room Number', 'Move-Out Date', 'Forwarding Address', 'Reason',
      'Additional Details', 'Inspection Availability', 'Key Return Method',
      'Satisfaction Rating', 'Appreciated Aspects', 'Improvements', 'Would Recommend',
      'Status', 'Final Inspection Date', 'Deposit Returned', 'Notes'
    ],
    
    GUEST_ROOMS: [
      'Booking ID', 'Room Number', 'Room Name', 'Room Type', 'Max Occupancy',
      'Amenities', 'Daily Rate', 'Weekly Rate', 'Monthly Rate', 'Status',
      'Last Cleaned', 'Maintenance Notes', 'Check-In Date', 'Check-Out Date',
      'Number of Nights', 'Number of Guests', 'Current Guest',
      'Purpose of Visit', 'Special Requests', 'Source', 'Total Amount',
      'Payment Status', 'Booking Status', 'Notes'
    ],
    
    
    MAINTENANCE: [
      'Request ID', 'Timestamp', 'Room/Area', 'Issue Type', 'Priority',
      'Description', 'Reported By', 'Contact Info', 'Assigned To',
      'Status', 'Estimated Cost', 'Actual Cost', 'Date Started',
      'Date Completed', 'Parts Used', 'Labor Hours', 'Photos', 'Notes'
    ],
    
    DOCUMENTS: [
      'Document ID', 'Document Type', 'Title', 'Related To', 'Entity ID',
      'File URL', 'Upload Date', 'Uploaded By', 'Expiration Date',
      'Status', 'Access Level', 'Tags', 'Notes'
    ],
    
    SETTINGS: [
      'Setting Key', 'Setting Value', 'Description', 'Category',
      'Last Modified', 'Modified By'
    ]
  },
  
  /**
   * Initialize all sheets with proper structure
   */
  initializeAllSheets: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    try {
      // Create/setup each sheet
      Object.entries(CONFIG.SHEETS).forEach(([key, sheetName]) => {
        if (key === 'GUEST_BOOKINGS') return; // responses sheet handled separately
        this.setupSheet(ss, sheetName, this.HEADERS[key]);
      });
      
      // Apply specific formatting
      this.applySheetSpecificFormatting(ss);
      
      Logger.log('All sheets initialized successfully');
      
    } catch (error) {
      Logger.log(`Error initializing sheets: ${error.toString()}`);
      throw error;
    }
  },

  /**
   * Initialize base sheets only (excluding form response sheets)
   */
  initializeCoreSheets: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const baseSheets = [
      CONFIG.SHEETS.TENANTS,
      CONFIG.SHEETS.BUDGET,
      CONFIG.SHEETS.GUEST_ROOMS,
      CONFIG.SHEETS.MAINTENANCE,
      CONFIG.SHEETS.DOCUMENTS,
      CONFIG.SHEETS.SETTINGS
    ];

    try {
      baseSheets.forEach(name => {
        const key = Object.keys(CONFIG.SHEETS).find(k => CONFIG.SHEETS[k] === name);
        this.setupSheet(ss, name, this.HEADERS[key]);
      });

      this.applySheetSpecificFormatting(ss);

      Logger.log('Core sheets initialized successfully');

    } catch (error) {
      Logger.log(`Error initializing core sheets: ${error.toString()}`);
      throw error;
    }
  },
  
  /**
   * Setup individual sheet
   */
  setupSheet: function(spreadsheet, sheetName, headers) {
    let sheet = spreadsheet.getSheetByName(sheetName);
    
    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      Logger.log(`Created sheet: ${sheetName}`);
    } else {
      // Clear existing content but preserve structure
      if (sheet.getLastRow() > 1) {
        sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
      }
      Logger.log(`Cleared sheet: ${sheetName}`);
    }
    
    // Set headers
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    
    // Apply header formatting
    headerRange
      .setFontWeight('bold')
      .setFontSize(11)
      .setBackground('#1c4587')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID);
    
    // Set frozen rows
    sheet.setFrozenRows(1);
    
    // Auto-resize columns
    for (let i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
      // Add some padding
      const currentWidth = sheet.getColumnWidth(i);
      sheet.setColumnWidth(i, Math.max(currentWidth + 20, 100));
    }
    
    // Set row height for headers
    sheet.setRowHeight(1, 35);
    
    return sheet;
  },
  
  /**
   * Apply sheet-specific formatting
   */
  applySheetSpecificFormatting: function(spreadsheet) {
    
    // Tenants sheet formatting
    const tenantsSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.TENANTS);
    if (tenantsSheet) {
      this.formatTenantsSheet(tenantsSheet);
    }
    
    // Budget sheet formatting
    const budgetSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.BUDGET);
    if (budgetSheet) {
      this.formatBudgetSheet(budgetSheet);
    }
    
    // Guest Rooms sheet formatting
    const guestRoomsSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.GUEST_ROOMS);
    if (guestRoomsSheet) {
      this.formatGuestRoomsSheet(guestRoomsSheet);
    }
    
    
    // Maintenance sheet formatting
    const maintenanceSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.MAINTENANCE);
    if (maintenanceSheet) {
      this.formatMaintenanceSheet(maintenanceSheet);
    }

    SpreadsheetApp.flush();
  },
  
  /**
   * Format Tenants Sheet
   */
  formatTenantsSheet: function(sheet) {
    if (sheet.getLastRow() < 2) return;
    
    const numRows = sheet.getMaxRows() - 1;

    // Alternate row colors using banding for efficiency
    const existingBandings = sheet.getBandings();
    existingBandings.forEach(b => b.remove());
    sheet.getRange(2, 1, sheet.getMaxRows() - 1, sheet.getLastColumn())
      .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
    
    // Conditional formatting for Room Status (column 9)
    const statusRange = sheet.getRange(2, 9, numRows, 1);

    const rules = [];
    
    // Occupied - Green
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(CONFIG.STATUS.ROOM.OCCUPIED)
      .setBackground('#d9ead3')
      .setFontColor('#137333')
      .setRanges([statusRange])
      .build());
    
    // Vacant - Yellow
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(CONFIG.STATUS.ROOM.VACANT)
      .setBackground('#fff2cc')
      .setFontColor('#b45f06')
      .setRanges([statusRange])
      .build());
    
    // Maintenance - Red
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(CONFIG.STATUS.ROOM.MAINTENANCE)
      .setBackground('#f4cccc')
      .setFontColor('#cc0000')
      .setRanges([statusRange])
      .build());
    
    // Payment Status conditional formatting (column 11)
    const paymentStatusRange = sheet.getRange(2, 11, numRows, 1);
    
    // Paid - Green
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(CONFIG.STATUS.PAYMENT.PAID)
      .setBackground('#d9ead3')
      .setFontColor('#137333')
      .setRanges([paymentStatusRange])
      .build());
    
    // Due - Yellow
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(CONFIG.STATUS.PAYMENT.DUE)
      .setBackground('#fff2cc')
      .setFontColor('#b45f06')
      .setRanges([paymentStatusRange])
      .build());
    
    // Overdue - Red
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(CONFIG.STATUS.PAYMENT.OVERDUE)
      .setBackground('#f4cccc')
      .setFontColor('#cc0000')
      .setRanges([paymentStatusRange])
      .build());

    // Data validation for Room Status dropdown
    const roomStatusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList([
        CONFIG.STATUS.ROOM.OCCUPIED,
        CONFIG.STATUS.ROOM.VACANT,
        CONFIG.STATUS.ROOM.MAINTENANCE,
        CONFIG.STATUS.ROOM.PENDING
      ], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, 9, numRows, 1).setDataValidation(roomStatusRule);

    // Data validation for Payment Status dropdown
    const paymentStatusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList([
        CONFIG.STATUS.PAYMENT.PAID,
        CONFIG.STATUS.PAYMENT.DUE,
        CONFIG.STATUS.PAYMENT.OVERDUE,
        CONFIG.STATUS.PAYMENT.PARTIAL
      ], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, 11, numRows, 1).setDataValidation(paymentStatusRule);

    sheet.setConditionalFormatRules(rules);
    
    // Format specific columns
    // Rental Price, Negotiated Price, Security Deposit (columns 2, 3, 8)
    sheet.getRange(2, 2, numRows, 1).setNumberFormat('$#,##0.00'); // Rental Price
    sheet.getRange(2, 3, numRows, 1).setNumberFormat('$#,##0.00'); // Negotiated Price
    sheet.getRange(2, 8, numRows, 1).setNumberFormat('$#,##0.00'); // Security Deposit
    
    // Date columns (7, 10, 12, 14)
    sheet.getRange(2, 7, numRows, 1).setNumberFormat('yyyy-mm-dd'); // Move-in Date
    sheet.getRange(2, 10, numRows, 1).setNumberFormat('yyyy-mm-dd'); // Last Payment
    sheet.getRange(2, 12, numRows, 1).setNumberFormat('yyyy-mm-dd'); // Move-out Date
    sheet.getRange(2, 14, numRows, 1).setNumberFormat('yyyy-mm-dd'); // Lease End Date
  },
  
  /**
   * Format Budget Sheet
   */
  formatBudgetSheet: function(sheet) {
    if (sheet.getLastRow() < 2) return;
    
    const numRows = sheet.getMaxRows() - 1;
    
    // Format Date column (column 1)
    sheet.getRange(2, 1, numRows, 1).setNumberFormat('yyyy-mm-dd');
    
    // Format Amount column (column 4)
    sheet.getRange(2, 4, numRows, 1).setNumberFormat('$#,##0.00');
    
    // Conditional formatting for positive/negative amounts
    const amountRange = sheet.getRange(2, 4, numRows, 1);
    
    const rules = [];
    
    // Positive amounts (income) - Green
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setFontColor('#137333')
      .setRanges([amountRange])
      .build());
    
    // Negative amounts (expenses) - Red
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setFontColor('#cc0000')
      .setRanges([amountRange])
      .build());

    sheet.setConditionalFormatRules(rules);

    // Data validation dropdowns based on existing values
    const typeList = this.getUniqueValues(CONFIG.SHEETS.BUDGET, 2);
    if (typeList.length > 0) {
      const typeRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(typeList.sort(), true)
        .setAllowInvalid(false)
        .build();
      sheet.getRange(2, 2, numRows, 1).setDataValidation(typeRule);
    }

    const categoryList = this.getUniqueValues(CONFIG.SHEETS.BUDGET, 5);
    if (categoryList.length > 0) {
      const catRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(categoryList.sort(), true)
        .setAllowInvalid(false)
        .build();
      sheet.getRange(2, 5, numRows, 1).setDataValidation(catRule);
    }

    const methodList = this.getUniqueValues(CONFIG.SHEETS.BUDGET, 6);
    if (methodList.length > 0) {
      const methodRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(methodList.sort(), true)
        .setAllowInvalid(false)
        .build();
      sheet.getRange(2, 6, numRows, 1).setDataValidation(methodRule);
    }
  },
  
  /**
   * Format Guest Rooms Sheet
   */
  formatGuestRoomsSheet: function(sheet) {
    if (sheet.getLastRow() < 2) return;

    const numRows = sheet.getMaxRows() - 1;

    // Remove duplicate "Maintenance Notes" column if present
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const firstIndex = headers.indexOf('Maintenance Notes');
    const lastIndex = headers.lastIndexOf('Maintenance Notes');
    if (firstIndex !== -1 && lastIndex !== -1 && firstIndex !== lastIndex) {
      sheet.deleteColumn(firstIndex + 1); // Convert to 1-based index
    }

    // Format rate columns (7, 8, 9)
    sheet.getRange(2, 7, numRows, 1).setNumberFormat('$#,##0.00'); // Daily Rate
    sheet.getRange(2, 8, numRows, 1).setNumberFormat('$#,##0.00'); // Weekly Rate
    sheet.getRange(2, 9, numRows, 1).setNumberFormat('$#,##0.00'); // Monthly Rate
    
    // Format date columns (11, 13, 14)
    sheet.getRange(2, 11, numRows, 1).setNumberFormat('yyyy-mm-dd'); // Last Cleaned
    sheet.getRange(2, 13, numRows, 1).setNumberFormat('yyyy-mm-dd'); // Check-In Date
    sheet.getRange(2, 14, numRows, 1).setNumberFormat('yyyy-mm-dd'); // Check-Out Date
    
    // Conditional formatting for Room Status (column 10)
    const statusRange = sheet.getRange(2, 10, numRows, 1);

    const rules = [];

    // Available - Green
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Available')
      .setBackground('#d9ead3')
      .setFontColor('#137333')
      .setRanges([statusRange])
      .build());

    // Occupied - Red
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Occupied')
      .setBackground('#f4cccc')
      .setFontColor('#cc0000')
      .setRanges([statusRange])
      .build());

    // Maintenance - Orange
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Maintenance')
      .setBackground('#fce5cd')
      .setFontColor('#b45f06')
      .setRanges([statusRange])
      .build());

    sheet.setConditionalFormatRules(rules);

    // Data validation dropdowns
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Available', 'Occupied', 'Maintenance'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, 10, numRows, 1).setDataValidation(statusRule);

    const paymentRule = SpreadsheetApp.newDataValidation()
      .requireValueInList([
        CONFIG.STATUS.PAYMENT.PAID,
        CONFIG.STATUS.PAYMENT.DUE,
        CONFIG.STATUS.PAYMENT.OVERDUE,
        CONFIG.STATUS.PAYMENT.PARTIAL
      ], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, 22, numRows, 1).setDataValidation(paymentRule);

    const bookingRule = SpreadsheetApp.newDataValidation()
      .requireValueInList([
        CONFIG.STATUS.BOOKING.PENDING,
        CONFIG.STATUS.BOOKING.CONFIRMED,
        CONFIG.STATUS.BOOKING.CHECKED_IN,
        CONFIG.STATUS.BOOKING.CHECKED_OUT,
        CONFIG.STATUS.BOOKING.CANCELLED
      ], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, 23, numRows, 1).setDataValidation(bookingRule);

    // Additional formatting
    sheet.getRange(2, 15, numRows, 2).setNumberFormat('0'); // Nights & Guests
    sheet.getRange(2, 21, numRows, 1).setNumberFormat('$#,##0.00'); // Total
  },
  
  /**
   * Format Maintenance Sheet
   */
  formatMaintenanceSheet: function(sheet) {
    if (sheet.getLastRow() < 2) return;
    
    const numRows = sheet.getMaxRows() - 1;
    
    // Format cost columns (11, 12)
    sheet.getRange(2, 11, numRows, 1).setNumberFormat('$#,##0.00'); // Estimated Cost
    sheet.getRange(2, 12, numRows, 1).setNumberFormat('$#,##0.00'); // Actual Cost
    
    // Format date columns (13, 14)
    sheet.getRange(2, 13, numRows, 1).setNumberFormat('yyyy-mm-dd'); // Date Started
    sheet.getRange(2, 14, numRows, 1).setNumberFormat('yyyy-mm-dd'); // Date Completed
    
    // Conditional formatting for Priority (column 5)
    const priorityRange = sheet.getRange(2, 5, numRows, 1);
    
    const priorityRules = [];
    
    // High Priority - Red
    priorityRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('High')
      .setBackground('#f4cccc')
      .setFontColor('#cc0000')
      .setRanges([priorityRange])
      .build());
    
    // Medium Priority - Yellow
    priorityRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Medium')
      .setBackground('#fff2cc')
      .setFontColor('#b45f06')
      .setRanges([priorityRange])
      .build());
    
    // Low Priority - Green
    priorityRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Low')
      .setBackground('#d9ead3')
      .setFontColor('#137333')
      .setRanges([priorityRange])
      .build());
    
    // Conditional formatting for Status (column 10)
    const statusRange = sheet.getRange(2, 10, numRows, 1);
    
    const statusRules = [];
    
    // Open - Red
    statusRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(CONFIG.STATUS.MAINTENANCE.OPEN)
      .setBackground('#f4cccc')
      .setFontColor('#cc0000')
      .setRanges([statusRange])
      .build());
    
    // In Progress - Yellow
    statusRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(CONFIG.STATUS.MAINTENANCE.IN_PROGRESS)
      .setBackground('#fff2cc')
      .setFontColor('#b45f06')
      .setRanges([statusRange])
      .build());
    
    // Completed - Green
    statusRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(CONFIG.STATUS.MAINTENANCE.COMPLETED)
      .setBackground('#d9ead3')
      .setFontColor('#137333')
      .setRanges([statusRange])
      .build());
    
    sheet.setConditionalFormatRules([...priorityRules, ...statusRules]);

    // Data validation dropdowns from existing values
    const issueTypes = this.getUniqueValues(CONFIG.SHEETS.MAINTENANCE, 4);
    if (issueTypes.length > 0) {
      const issueRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(issueTypes.sort(), true)
        .setAllowInvalid(false)
        .build();
      sheet.getRange(2, 4, numRows, 1).setDataValidation(issueRule);
    }

    const priorityList = this.getUniqueValues(CONFIG.SHEETS.MAINTENANCE, 5);
    if (priorityList.length > 0) {
      const priorityRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(priorityList.sort(), true)
        .setAllowInvalid(false)
        .build();
      sheet.getRange(2, 5, numRows, 1).setDataValidation(priorityRule);
    }

    const statusList = this.getUniqueValues(CONFIG.SHEETS.MAINTENANCE, 10);
    if (statusList.length > 0) {
      const statusRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(statusList.sort(), true)
        .setAllowInvalid(false)
        .build();
      sheet.getRange(2, 10, numRows, 1).setDataValidation(statusRule);
    }
  },
  
  /**
   * Get sheet by name with error handling
   */
  getSheet: function(sheetName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error(`Sheet '${sheetName}' not found. Please initialize the system first.`);
    }
    
    return sheet;
  },
  
  /**
   * Get next available row in a sheet
   */
  getNextRow: function(sheetName) {
    const sheet = this.getSheet(sheetName);
    return sheet.getLastRow() + 1;
  },
  
  /**
   * Add row to sheet with validation
   */
  addRow: function(sheetName, data) {
    const sheet = this.getSheet(sheetName);
    const headers = this.HEADERS[Object.keys(CONFIG.SHEETS).find(key => CONFIG.SHEETS[key] === sheetName)];
    
    if (!headers) {
      throw new Error(`Headers not defined for sheet: ${sheetName}`);
    }
    
    if (data.length > headers.length) {
      Logger.log(`Warning: Data length (${data.length}) exceeds header length (${headers.length}) for sheet ${sheetName}`);
    }
    
    sheet.appendRow(data);
    return sheet.getLastRow();
  },
  
  /**
   * Update row in sheet
   */
  updateRow: function(sheetName, rowNumber, data) {
    const sheet = this.getSheet(sheetName);
    const range = sheet.getRange(rowNumber, 1, 1, data.length);
    range.setValues([data]);
  },
  
  /**
   * Find rows by criteria
   */
  findRows: function(sheetName, columnIndex, value) {
    const sheet = this.getSheet(sheetName);
    const data = sheet.getDataRange().getValues();
    const matchingRows = [];
    
    for (let i = 1; i < data.length; i++) { // Skip header row
      if (data[i][columnIndex - 1] === value) {
        matchingRows.push({
          rowNumber: i + 1,
          data: data[i]
        });
      }
    }
    
    return matchingRows;
  },
  
  /**
   * Get all data from sheet
   */
  getAllData: function(sheetName, includeHeaders = false) {
    const sheet = this.getSheet(sheetName);
    
    if (sheet.getLastRow() < 2) {
      return [];
    }
    
    const startRow = includeHeaders ? 1 : 2;
    const numRows = sheet.getLastRow() - startRow + 1;
    
    if (numRows <= 0) {
      return [];
    }
    
    return sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();
  },

  /**
   * Get a mapping of header names to 1-based column indexes
   */
  getHeaderMap: function(sheetName) {
    const sheet = this.getSheet(sheetName);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const map = {};
    headers.forEach((h, i) => {
      if (typeof h === 'string' && h.trim() !== '') {
        map[h.trim()] = i + 1;
      }
    });
    return map;
  },

  /**
   * Get unique non-empty values from a column in a sheet
   */
  getUniqueValues: function(sheetName, columnIndex) {
    const data = this.getAllData(sheetName);
    const values = new Set();
    data.forEach(row => {
      const val = row[columnIndex - 1];
      if (val !== '' && val != null) {
        values.add(val.toString());
      }
    });
    return Array.from(values);
  },

  /**
   * Remove asterisks from header row of a sheet
   */
  cleanHeaderAsterisks: function(sheet) {
    const lastColumn = sheet.getLastColumn();
    if (lastColumn === 0) return;
    const range = sheet.getRange(1, 1, 1, lastColumn);
    const headers = range.getValues()[0];
    const cleaned = headers.map(h => (typeof h === 'string') ? h.replace(/\s*\*+$/, '').trim() : h);
    range.setValues([cleaned]);
  },
  
  /**
   * Clear sheet data (keep headers)
   */
  clearSheetData: function(sheetName) {
    const sheet = this.getSheet(sheetName);
    
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }
  },
  
  /**
   * Protect sheet with specified permissions
   */
  protectSheet: function(sheetName, description, allowedEmails = []) {
    const sheet = this.getSheet(sheetName);
    const protection = sheet.protect().setDescription(description);
    
    // Allow specific users to edit
    if (allowedEmails.length > 0) {
      protection.addEditors(allowedEmails);
    }
    
    // Remove default editors except owner
    protection.removeEditors(protection.getEditors());
    if (protection.canDomainEdit()) {
      protection.setDomainEdit(false);
    }
  }
};

Logger.log('SheetManager module loaded successfully');
