// DataManager.gs - Data Management & Sample Data Creation
// Handles data operations and creates sample data for system initialization

const DataManager = {
  
  /**
   * Create sample data for all sheets
   */
  createSampleData: function() {
    try {
      this.createSampleTenants();
      this.createSampleGuestRooms();
      this.createSampleBudgetEntries();
      this.createSampleSettings();
      
      Logger.log('Sample data created successfully');
      
    } catch (error) {
      Logger.log(`Error creating sample data: ${error.toString()}`);
      throw error;
    }
  },
  
  /**
   * Create sample tenant data
   */
  createSampleTenants: function() {
    const sampleTenants = [
      [
        '101', 800, '', '', '', '', '', 0, 
        CONFIG.STATUS.ROOM.VACANT, '', '', '', '', '', 'First floor corner room'
      ],
      [
        '102', 800, 750, 'Sarah Johnson', 'sarah.j@email.com', '(555) 123-4567', 
        new Date(2024, 8, 15), 800, CONFIG.STATUS.ROOM.OCCUPIED, 
        new Date(2024, 11, 1), CONFIG.STATUS.PAYMENT.PAID, '', 
        'Emergency: Mom (555) 987-6543', new Date(2025, 8, 15), 'Quiet tenant, no issues'
      ],
      [
        '103', 900, '', '', '', '', '', 0, 
        CONFIG.STATUS.ROOM.VACANT, '', '', '', '', '', 'Larger room with private bath'
      ],
      [
        '104', 900, 850, 'Michael Chen', 'mchen@email.com', '(555) 234-5678', 
        new Date(2024, 9, 1), 900, CONFIG.STATUS.ROOM.OCCUPIED, 
        new Date(2024, 10, 3), CONFIG.STATUS.PAYMENT.DUE, '', 
        'Emergency: Sister (555) 876-5432', new Date(2025, 9, 1), 'Graduate student, very responsible'
      ],
      [
        '201', 850, '', '', '', '', '', 0, 
        CONFIG.STATUS.ROOM.VACANT, '', '', '', '', '', 'Second floor, quiet side'
      ],
      [
        '202', 850, 825, 'Emma Rodriguez', 'emma.rod@email.com', '(555) 345-6789', 
        new Date(2024, 7, 20), 850, CONFIG.STATUS.ROOM.OCCUPIED, 
        new Date(2024, 9, 25), CONFIG.STATUS.PAYMENT.OVERDUE, '', 
        'Emergency: Father (555) 765-4321', new Date(2025, 7, 20), 'Working professional, usually travels'
      ],
      [
        '203', 950, '', '', '', '', '', 0, 
        CONFIG.STATUS.ROOM.MAINTENANCE, '', '', '', '', '', 'Under renovation - new flooring'
      ],
      [
        '204', 950, 900, 'James Wilson', 'j.wilson@email.com', '(555) 456-7890', 
        new Date(2024, 10, 10), 950, CONFIG.STATUS.ROOM.OCCUPIED, 
        new Date(2024, 11, 5), CONFIG.STATUS.PAYMENT.PAID, '', 
        'Emergency: Wife (555) 654-3210', new Date(2025, 10, 10), 'Recently married, very neat'
      ]
    ];
    
    // Clear existing data and add sample data
    SheetManager.clearSheetData(CONFIG.SHEETS.TENANTS);
    
    sampleTenants.forEach(tenant => {
      SheetManager.addRow(CONFIG.SHEETS.TENANTS, tenant);
    });
  },
  
  /**
   * Create sample guest rooms
   */
  createSampleGuestRooms: function() {
    const sampleGuestRooms = [
      [
        '', 'G1', 'Executive Suite', 'Executive Suite', 2,
        'Queen bed, Private bath, Mini fridge, Desk, City view',
        85, 510, 1800, 'Available', new Date(), 'Recently deep cleaned',
        '', '', '', '', '', '', '', '', '', '', '', ''
      ],
      [
        '', 'G2', 'Comfort Room', 'Standard Room', 2,
        'Double bed, Shared bath, WiFi, Coffee maker',
        65, 390, 1400, 'Available', new Date(), 'Standard maintenance completed',
        '', '', '', '', '', '', '', '', '', '', '', ''
      ]
    ];
    
    SheetManager.clearSheetData(CONFIG.SHEETS.GUEST_ROOMS);
    
    sampleGuestRooms.forEach(room => {
      SheetManager.addRow(CONFIG.SHEETS.GUEST_ROOMS, room);
    });
  },
  
  /**
   * Create sample budget entries
   */
  createSampleBudgetEntries: function() {
    const currentDate = new Date();
    const sampleBudgetEntries = [
      // Current month income
      [
        new Date(currentDate.getFullYear(), currentDate.getMonth(), 1),
        'Rent Income', 'Monthly rent - Sarah Johnson Room 102', 750, 'Rent',
        'Check', 'RENT-102-' + Utils.formatDate(currentDate, 'yyyyMM'), 'Sarah Johnson', ''
      ],
      [
        new Date(currentDate.getFullYear(), currentDate.getMonth(), 5),
        'Rent Income', 'Monthly rent - James Wilson Room 204', 900, 'Rent',
        'Bank Transfer', 'RENT-204-' + Utils.formatDate(currentDate, 'yyyyMM'), 'James Wilson', ''
      ],
      [
        new Date(currentDate.getFullYear(), currentDate.getMonth(), 3),
        'Guest Room Income', 'Guest room rental - John Smith Room G1', 255, 'Guest Room',
        'Credit Card', 'GUEST-G1-001', 'John Smith', ''
      ],
      [
        new Date(currentDate.getFullYear(), currentDate.getMonth(), 12),
        'Security Deposit', 'Security deposit - Michael Chen Room 104', 900, 'Deposit',
        'Cash', 'DEPOSIT-104-001', 'Michael Chen', ''
      ],
      
      // Expenses
      [
        new Date(currentDate.getFullYear(), currentDate.getMonth(), 8),
        'Utility Expense', 'Monthly electricity bill', -245.50, 'Electricity',
        'Auto Pay', 'ELEC-' + Utils.formatDate(currentDate, 'yyyyMM'), '', 'Electric Company'
      ],
      [
        new Date(currentDate.getFullYear(), currentDate.getMonth(), 10),
        'Utility Expense', 'Water and sewage bill', -89.75, 'Water',
        'Check', 'WATER-' + Utils.formatDate(currentDate, 'yyyyMM'), '', 'City Water Dept'
      ],
      [
        new Date(currentDate.getFullYear(), currentDate.getMonth(), 15),
        'Maintenance Expense', 'Plumbing repair - Room G3', -125.00, 'Maintenance',
        'Cash', 'MAINT-G3-001', '', 'ABC Plumbing'
      ],
      [
        new Date(currentDate.getFullYear(), currentDate.getMonth(), 7),
        'Supply Expense', 'Cleaning supplies and toiletries', -67.32, 'Supplies',
        'Credit Card', 'SUPPLY-' + Utils.formatDate(currentDate, 'yyyyMM'), '', 'Supply Store'
      ],
      
      // Previous month entries
      [
        new Date(currentDate.getFullYear(), currentDate.getMonth() - 1, 1),
        'Rent Income', 'Monthly rent - Sarah Johnson Room 102', 750, 'Rent',
        'Check', 'RENT-102-' + Utils.formatDate(new Date(currentDate.getFullYear(), currentDate.getMonth() - 1, 1), 'yyyyMM'), 'Sarah Johnson', ''
      ],
      [
        new Date(currentDate.getFullYear(), currentDate.getMonth() - 1, 3),
        'Rent Income', 'Monthly rent - Emma Rodriguez Room 202', 825, 'Rent',
        'Bank Transfer', 'RENT-202-' + Utils.formatDate(new Date(currentDate.getFullYear(), currentDate.getMonth() - 1, 1), 'yyyyMM'), 'Emma Rodriguez', ''
      ],
      [
        new Date(currentDate.getFullYear(), currentDate.getMonth() - 1, 18),
        'Guest Room Income', 'Guest room rental - Business traveler Room G2', 195, 'Guest Room',
        'Credit Card', 'GUEST-G2-002', 'Corporate Guest', ''
      ]
    ];
    
    SheetManager.clearSheetData(CONFIG.SHEETS.BUDGET);
    
    sampleBudgetEntries.forEach(entry => {
      SheetManager.addRow(CONFIG.SHEETS.BUDGET, entry);
    });
  },
  
  /**
   * Create sample guest bookings
   */
  createSampleGuestBookings: function() {
    const today = new Date();
    const tomorrow = new Date(today);
    tomorrow.setDate(tomorrow.getDate() + 1);
    
    const nextWeek = new Date(today);
    nextWeek.setDate(nextWeek.getDate() + 7);
    
    const sampleBookings = [
      [
        'BK001', new Date(), 'Alice Thompson', 'alice.t@email.com', '(555) 111-2222',
        'G1', tomorrow, new Date(tomorrow.getTime() + 3 * 24 * 60 * 60 * 1000), 3, 2,
        'Business', 'Late check-in requested', 255, 255, 'Paid', CONFIG.STATUS.BOOKING.CONFIRMED,
        'Direct', 'Confirmed for tomorrow arrival'
      ],
      [
        'BK002', new Date(), 'Robert Davis', 'r.davis@email.com', '(555) 333-4444',
        'G2', today, new Date(today.getTime() + 24 * 60 * 60 * 1000), 1, 1,
        'Tourism', '', 65, 65, 'Paid', CONFIG.STATUS.BOOKING.CHECKED_IN,
        'Online', 'Currently checked in'
      ],
      [
        'BK003', new Date(), 'Maria Garcia', 'maria.g@email.com', '(555) 555-6666',
        'G1', nextWeek, new Date(nextWeek.getTime() + 5 * 24 * 60 * 60 * 1000), 5, 2,
        'Visiting family', 'Quiet room preferred', 425, 200, 'Partial', CONFIG.STATUS.BOOKING.CONFIRMED,
        'Phone', 'Balance due at check-in'
      ]
    ];
    
    SheetManager.clearSheetData(CONFIG.SHEETS.GUEST_BOOKINGS);
    
    sampleBookings.forEach(booking => {
      SheetManager.addRow(CONFIG.SHEETS.GUEST_BOOKINGS, booking);
    });
  },
  
  /**
   * Create sample maintenance requests
   */
  createSampleMaintenanceRequests: function() {
    const sampleMaintenance = [
      [
        'MR001', new Date(), 'Room 102', 'Plumbing', 'Medium',
        'Faucet drips constantly in bathroom sink', 'Sarah Johnson', 'sarah.j@email.com',
        'Maintenance Team', CONFIG.STATUS.MAINTENANCE.IN_PROGRESS, 75, 0,
        new Date(), '', 'Faucet parts', 2, '', 'Parts ordered, work scheduled'
      ],
      [
        'MR002', new Date(), 'Common Kitchen', 'Electrical', 'High',
        'Light fixture flickering, possible electrical hazard', 'Michael Chen', 'mchen@email.com',
        'Electrician', CONFIG.STATUS.MAINTENANCE.OPEN, 150, 0,
        '', '', '', 0, '', 'Urgent - safety concern'
      ],
      [
        'MR003', new Date(Date.now() - 5 * 24 * 60 * 60 * 1000), 'Room 204', 'HVAC', 'Low',
        'Heating not working efficiently', 'James Wilson', 'j.wilson@email.com',
        'HVAC Tech', CONFIG.STATUS.MAINTENANCE.COMPLETED, 125, 135,
        new Date(Date.now() - 3 * 24 * 60 * 60 * 1000), new Date(Date.now() - 1 * 24 * 60 * 60 * 1000), 'Filter, labor', 3, '', 'Completed - filter replaced'
      ],
      [
        'MR004', new Date(), 'Room G3', 'Plumbing', 'High',
        'Toilet not flushing properly, water backing up', 'Property Manager', CONFIG.SYSTEM.MANAGER_EMAIL,
        'ABC Plumbing', CONFIG.STATUS.MAINTENANCE.OPEN, 200, 0,
        '', '', '', 0, '', 'Guest room out of service'
      ],
      [
        'MR005', new Date(Date.now() - 2 * 24 * 60 * 60 * 1000), 'Common Area', 'General', 'Low',
        'Light bulb replacement needed in hallway', 'Emma Rodriguez', 'emma.rod@email.com',
        'Maintenance Team', CONFIG.STATUS.MAINTENANCE.COMPLETED, 15, 12,
        new Date(Date.now() - 1 * 24 * 60 * 60 * 1000), new Date(Date.now() - 1 * 24 * 60 * 60 * 1000), 'LED bulb', 0.5, '', 'Quick fix completed'
      ]
    ];
    
    SheetManager.clearSheetData(CONFIG.SHEETS.MAINTENANCE);
    
    sampleMaintenance.forEach(request => {
      SheetManager.addRow(CONFIG.SHEETS.MAINTENANCE, request);
    });
  },
  
  /**
   * Create sample tenant applications
   */
  createSampleApplications: function() {
    const sampleApplications = [
      [
        'APP001', new Date(), 'Jennifer Williams', 'j.williams@email.com', '(555) 777-8888',
        '123 Current St, City, State 12345', new Date(Date.now() + 14 * 24 * 60 * 60 * 1000), '103',
        '6-12 months', 'Full-time employed', 'Tech Solutions Inc.', 4200,
        'Best friend: Lisa (555) 999-0000', 'Former colleague: Mark (555) 888-7777',
        'Sister: Amy Williams (555) 666-5555', 'Honda Civic, Blue, License ABC123',
        'Software developer, quiet, clean, no pets. Looking for a peaceful place to focus on work.',
        'None', 'Under Review', 'Good application, employment verified', new Date()
      ],
      [
        'APP002', new Date(Date.now() - 2 * 24 * 60 * 60 * 1000), 'David Kim', 'd.kim@email.com', '(555) 444-3333',
        '456 Another Ave, City, State 54321', new Date(Date.now() + 30 * 24 * 60 * 60 * 1000), '101',
        '12+ months', 'Graduate student', 'State University', 2800,
        'Professor: Dr. Smith (555) 111-2222', 'Roommate: Alex Chen (555) 333-4444',
        'Father: Mr. Kim (555) 555-6666', 'No vehicle',
        'PhD student in engineering. Very studious and respectful. Non-smoker, no parties.',
        'Quiet study space important', 'Approved', 'Excellent references, clean background', new Date()
      ],
      [
        'APP003', new Date(Date.now() - 1 * 24 * 60 * 60 * 1000), 'Lisa Martinez', 'lisa.m@email.com', '(555) 222-1111',
        '789 Third Rd, City, State 98765', new Date(Date.now() + 21 * 24 * 60 * 60 * 1000), '201',
        '6 months', 'Contract work', 'Various clients', 3500,
        'Client: ABC Corp (555) 123-9999', 'Friend: Sarah J (555) 987-6543',
        'Mother: Mrs. Martinez (555) 456-7890', 'Toyota Prius, Silver, License DEF456',
        'Freelance graphic designer. Clean, organized, works from home occasionally.',
        'Good internet connection needed for work', 'Pending', 'Need to verify income consistency', ''
      ]
    ];
    
    SheetManager.clearSheetData(CONFIG.SHEETS.APPLICATIONS);
    
    sampleApplications.forEach(application => {
      SheetManager.addRow(CONFIG.SHEETS.APPLICATIONS, application);
    });
  },
  
  /**
   * Create sample move-out requests
   */
  createSampleMoveOuts: function() {
    const sampleMoveOuts = [
      [
        'MO001', new Date(Date.now() - 10 * 24 * 60 * 60 * 1000), 'Former Tenant', 'former@email.com', '(555) 999-8888',
        '105', new Date(Date.now() + 20 * 24 * 60 * 60 * 1000), '321 New Address Ln, New City, State 11111',
        'Job relocation', 'Got promoted and moving to company headquarters in another city.',
        'Weekday evenings after 5 PM', 'Will leave keys with property manager',
        5, 'Loved the community feel and friendly neighbors', 'Maybe better soundproofing between rooms',
        'Yes', 'Approved', new Date(Date.now() + 21 * 24 * 60 * 60 * 1000), 750, 'Standard 30-day notice, good tenant'
      ],
      [
        'MO002', new Date(), 'Current Tenant', 'current@email.com', '(555) 888-7777',
        '203', new Date(Date.now() + 35 * 24 * 60 * 60 * 1000), '654 Future St, Another City, State 22222',
        'Buying a house', 'Finally saved enough for a down payment on our first home!',
        'Weekend mornings preferred', 'Will coordinate key return during business hours',
        4, 'Great maintenance response time and fair pricing', 'Could use more parking spaces',
        'Definitely', 'Processing', '', 0, 'Congratulations on the home purchase!'
      ]
    ];
    
    SheetManager.clearSheetData(CONFIG.SHEETS.MOVEOUTS);
    
    sampleMoveOuts.forEach(moveOut => {
      SheetManager.addRow(CONFIG.SHEETS.MOVEOUTS, moveOut);
    });
  },
  
  /**
   * Create sample document records
   */
  createSampleDocuments: function() {
    const sampleDocuments = [
      [
        'DOC001', 'Lease Agreement', 'Standard Lease Template v2.1', 'System', 'TEMPLATE',
        'https://drive.google.com/file/d/sample_lease_template', new Date(), CONFIG.SYSTEM.MANAGER_EMAIL,
        '', 'Active', 'Public', 'lease,template,legal', 'Master lease agreement template'
      ],
      [
        'DOC002', 'House Rules', 'Community Guidelines 2024', 'Property', 'RULES',
        'https://drive.google.com/file/d/sample_house_rules', new Date(), CONFIG.SYSTEM.MANAGER_EMAIL,
        '', 'Active', 'Tenants', 'rules,guidelines,policy', 'Updated house rules and community guidelines'
      ],
      [
        'DOC003', 'Tenant Handbook', 'New Resident Guide', 'Property', 'HANDBOOK',
        'https://drive.google.com/file/d/sample_handbook', new Date(), CONFIG.SYSTEM.MANAGER_EMAIL,
        '', 'Active', 'Tenants', 'handbook,guide,welcome', 'Comprehensive guide for new tenants'
      ],
      [
        'DOC004', 'Maintenance Log', 'Q4 2024 Maintenance Summary', 'Report', 'MAINTENANCE',
        'https://drive.google.com/file/d/sample_maint_log', new Date(), CONFIG.SYSTEM.MANAGER_EMAIL,
        '', 'Active', 'Management', 'maintenance,report,quarterly', 'Quarterly maintenance activity report'
      ],
      [
        'DOC005', 'Financial Report', 'November 2024 Financial Summary', 'Report', 'FINANCIAL',
        'https://drive.google.com/file/d/sample_financial', new Date(), CONFIG.SYSTEM.MANAGER_EMAIL,
        '', 'Active', 'Owner', 'financial,report,monthly', 'Monthly financial performance report'
      ]
    ];
    
    SheetManager.clearSheetData(CONFIG.SHEETS.DOCUMENTS);
    
    sampleDocuments.forEach(doc => {
      SheetManager.addRow(CONFIG.SHEETS.DOCUMENTS, doc);
    });
  },
  
  /**
   * Create system settings
   */
  createSampleSettings: function() {
    const sampleSettings = [
      ['property_name', CONFIG.SYSTEM.PROPERTY_NAME, 'Name of the property', 'General', new Date(), Session.getActiveUser().getEmail()],
      ['manager_email', CONFIG.SYSTEM.MANAGER_EMAIL, 'Primary manager email address', 'General', new Date(), Session.getActiveUser().getEmail()],
      ['late_fee_days', CONFIG.SYSTEM.LATE_FEE_DAYS, 'Number of days before late fee applies', 'Financial', new Date(), Session.getActiveUser().getEmail()],
      ['late_fee_amount', CONFIG.SYSTEM.LATE_FEE_AMOUNT, 'Late fee amount in dollars', 'Financial', new Date(), Session.getActiveUser().getEmail()],
      ['currency', CONFIG.SYSTEM.CURRENCY, 'Currency code', 'Financial', new Date(), Session.getActiveUser().getEmail()],
      ['time_zone', CONFIG.SYSTEM.TIME_ZONE, 'System timezone', 'General', new Date(), Session.getActiveUser().getEmail()],
      ['check_in_time', '15:00', 'Guest room check-in time', 'Guest Rooms', new Date(), Session.getActiveUser().getEmail()],
      ['check_out_time', '11:00', 'Guest room check-out time', 'Guest Rooms', new Date(), Session.getActiveUser().getEmail()],
      ['wifi_network', 'ParsonageWiFi', 'WiFi network name', 'Amenities', new Date(), Session.getActiveUser().getEmail()],
      ['wifi_password', 'Welcome2024!', 'WiFi password', 'Amenities', new Date(), Session.getActiveUser().getEmail()],
      ['auto_reminders', 'true', 'Enable automated email reminders', 'Automation', new Date(), Session.getActiveUser().getEmail()],
      ['backup_frequency', 'weekly', 'How often to backup data', 'System', new Date(), Session.getActiveUser().getEmail()],
      ['maintenance_email', CONFIG.SYSTEM.MANAGER_EMAIL, 'Email for maintenance notifications', 'Maintenance', new Date(), Session.getActiveUser().getEmail()],
      ['guest_confirmation_auto', 'true', 'Auto-send guest booking confirmations', 'Guest Rooms', new Date(), Session.getActiveUser().getEmail()],
      ['system_version', '2.0', 'Current system version', 'System', new Date(), Session.getActiveUser().getEmail()]
    ];
    
    SheetManager.clearSheetData(CONFIG.SHEETS.SETTINGS);
    
    sampleSettings.forEach(setting => {
      SheetManager.addRow(CONFIG.SHEETS.SETTINGS, setting);
    });
  },
  
  /**
   * Create ALL sample data including missing sheets
   */
  createAllSampleData: function() {
    try {
      this.createSampleTenants();
      this.createSampleGuestRooms();
      this.createSampleBudgetEntries();
      this.createSampleMaintenanceRequests();
      this.createSampleApplications();
      this.createSampleMoveOuts();
      this.createSampleDocuments();
      this.createSampleSettings();
      
      Logger.log('All sample data created successfully');
      
    } catch (error) {
      Logger.log(`Error creating all sample data: ${error.toString()}`);
      throw error;
    }
  },
  
  /**
   * Export system data for backup
   */
  exportSystemData: function() {
    try {
      const ui = SpreadsheetApp.getUi();
      
      const exportData = {
        timestamp: new Date().toISOString(),
        version: '2.0',
        data: {}
      };
      
      // Export all sheet data
      Object.values(CONFIG.SHEETS).forEach(sheetName => {
        try {
          exportData.data[sheetName] = SheetManager.getAllData(sheetName, true); // Include headers
        } catch (error) {
          Logger.log(`Warning: Could not export sheet ${sheetName}: ${error.message}`);
        }
      });
      
      // Create backup file
      const fileName = `ParsonageBackup_${Utils.formatDate(new Date(), 'yyyyMMdd_HHmmss')}.json`;
      const jsonString = JSON.stringify(exportData, null, 2);
      
      // Create file in Google Drive
      const blob = Utilities.newBlob(jsonString, 'application/json', fileName);
      const file = DriveApp.createFile(blob);
      
      ui.alert(
        'Data Export Complete',
        `System data has been exported to:\n\n${fileName}\n\nFile ID: ${file.getId()}\n\nYou can find this file in your Google Drive.`,
        ui.ButtonSet.OK
      );
      
      return file.getId();
      
    } catch (error) {
      handleSystemError(error, 'exportSystemData');
    }
  },
  
  /**
   * Import system data from backup
   */
  importSystemData: function() {
    try {
      const ui = SpreadsheetApp.getUi();
      
      const response = ui.prompt(
        'Import System Data',
        'Enter the Google Drive file ID of your backup file:',
        ui.ButtonSet.OK_CANCEL
      );
      
      if (response.getSelectedButton() !== ui.Button.OK) return;
      
      const fileId = response.getResponseText().trim();
      
      if (!fileId) {
        ui.alert('No file ID provided.');
        return;
      }
      
      // Confirm import
      const confirmResponse = ui.alert(
        'Confirm Data Import',
        'This will replace ALL current data with the backup data. Are you sure?',
        ui.ButtonSet.YES_NO
      );
      
      if (confirmResponse !== ui.Button.YES) return;
      
      // Get file and parse JSON
      const file = DriveApp.getFileById(fileId);
      const jsonString = file.getBlob().getDataAsString();
      const importData = JSON.parse(jsonString);
      
      // Validate backup format
      if (!importData.data || !importData.version) {
        throw new Error('Invalid backup file format');
      }
      
      // Import each sheet
      let importedSheets = 0;
      Object.entries(importData.data).forEach(([sheetName, sheetData]) => {
        try {
          if (sheetData && sheetData.length > 0) {
            // Clear existing data
            SheetManager.clearSheetData(sheetName);
            
            // Import data (skip headers as they're already set)
            const dataRows = sheetData.slice(1);
            dataRows.forEach(row => {
              SheetManager.addRow(sheetName, row);
            });
            
            importedSheets++;
          }
        } catch (error) {
          Logger.log(`Warning: Could not import sheet ${sheetName}: ${error.message}`);
        }
      });
      
      ui.alert(
        'Data Import Complete',
        `Successfully imported data for ${importedSheets} sheets from backup created on:\n${new Date(importData.timestamp).toLocaleString()}`,
        ui.ButtonSet.OK
      );
      
    } catch (error) {
      SpreadsheetApp.getUi().alert(
        'Import Error',
        `Failed to import data: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  },
  
  /**
   * Clean old data based on retention policies
   */
  cleanOldData: function() {
    try {
      const ui = SpreadsheetApp.getUi();
      
      const response = ui.alert(
        'Clean Old Data',
        'This will remove old records based on retention policies:\n\n' +
        '• Budget entries older than 2 years\n' +
        '• Completed guest bookings older than 1 year\n' +
        '• Completed maintenance requests older than 1 year\n\n' +
        'Continue?',
        ui.ButtonSet.YES_NO
      );
      
      if (response !== ui.Button.YES) return;
      
      const today = new Date();
      const twoYearsAgo = new Date(today.getFullYear() - 2, today.getMonth(), today.getDate());
      const oneYearAgo = new Date(today.getFullYear() - 1, today.getMonth(), today.getDate());
      
      let cleanedRecords = 0;
      
      // Clean budget data older than 2 years
      const budgetData = SheetManager.getAllData(CONFIG.SHEETS.BUDGET);
      const budgetSheet = SheetManager.getSheet(CONFIG.SHEETS.BUDGET);
      
      for (let i = budgetData.length - 1; i >= 0; i--) {
        const entryDate = new Date(budgetData[i][0]);
        if (entryDate < twoYearsAgo) {
          budgetSheet.deleteRow(i + 2); // +2 for header and 0-based index
          cleanedRecords++;
        }
      }
      
      // Clean old guest bookings
      const bookingData = SheetManager.getAllData(CONFIG.SHEETS.GUEST_BOOKINGS);
      const bookingSheet = SheetManager.getSheet(CONFIG.SHEETS.GUEST_BOOKINGS);
      
      for (let i = bookingData.length - 1; i >= 0; i--) {
        const checkOutDate = new Date(bookingData[i][7]); // Check-out date column
        const status = bookingData[i][15]; // Booking status column
        
        if (checkOutDate < oneYearAgo && status === CONFIG.STATUS.BOOKING.CHECKED_OUT) {
          bookingSheet.deleteRow(i + 2);
          cleanedRecords++;
        }
      }
      
      // Clean old maintenance requests
      const maintenanceData = SheetManager.getAllData(CONFIG.SHEETS.MAINTENANCE);
      const maintenanceSheet = SheetManager.getSheet(CONFIG.SHEETS.MAINTENANCE);
      
      for (let i = maintenanceData.length - 1; i >= 0; i--) {
        const completedDate = new Date(maintenanceData[i][13]); // Date completed column
        const status = maintenanceData[i][9]; // Status column
        
        if (completedDate < oneYearAgo && status === CONFIG.STATUS.MAINTENANCE.COMPLETED) {
          maintenanceSheet.deleteRow(i + 2);
          cleanedRecords++;
        }
      }
      
      ui.alert(
        'Data Cleanup Complete',
        `Removed ${cleanedRecords} old records from the system.`,
        ui.ButtonSet.OK
      );
      
    } catch (error) {
      handleSystemError(error, 'cleanOldData');
    }
  },
  
  /**
   * Validate data integrity across sheets
   */
  validateDataIntegrity: function() {
    try {
      const issues = [];
      
      // Check tenant data consistency
      const tenantData = SheetManager.getAllData(CONFIG.SHEETS.TENANTS);
      tenantData.forEach((tenant, index) => {
        const rowNumber = index + 2;
        
        // Check for duplicate room numbers
        const roomNumber = tenant[0];
        if (roomNumber) {
          const duplicates = tenantData.filter((t, i) => i !== index && t[0] === roomNumber);
          if (duplicates.length > 0) {
            issues.push(`Duplicate room number ${roomNumber} found in row ${rowNumber}`);
          }
        }
        
        // Check email format
        const email = tenant[4];
        if (email && !Utils.isValidEmail(email)) {
          issues.push(`Invalid email format in row ${rowNumber}: ${email}`);
        }
        
        // Check price consistency
        const rentalPrice = tenant[1];
        const negotiatedPrice = tenant[2];
        if (negotiatedPrice && negotiatedPrice > rentalPrice * 1.5) {
          issues.push(`Unusually high negotiated price in row ${rowNumber}: ${negotiatedPrice} vs ${rentalPrice}`);
        }
      });
      
      // Check guest room data
      const guestRoomData = SheetManager.getAllData(CONFIG.SHEETS.GUEST_ROOMS);
      guestRoomData.forEach((room, index) => {
        const rowNumber = index + 2;
        
        // Check rate consistency
        const dailyRate = room[6];
        const weeklyRate = room[7];
        const monthlyRate = room[8];
        
        if (weeklyRate && dailyRate && weeklyRate > dailyRate * 7) {
          issues.push(`Weekly rate higher than daily rate * 7 in guest room row ${rowNumber}`);
        }
        
        if (monthlyRate && dailyRate && monthlyRate > dailyRate * 30) {
          issues.push(`Monthly rate higher than daily rate * 30 in guest room row ${rowNumber}`);
        }
      });
      
      // Check budget data
      const budgetData = SheetManager.getAllData(CONFIG.SHEETS.BUDGET);
      budgetData.forEach((entry, index) => {
        const rowNumber = index + 2;
        const amount = entry[3];
        
        // Check for unusually large amounts
        if (Math.abs(amount) > 10000) {
          issues.push(`Unusually large amount in budget row ${rowNumber}: ${amount}`);
        }
        
        // Check date validity
        const entryDate = entry[0];
        if (!(entryDate instanceof Date) || isNaN(entryDate.getTime())) {
          issues.push(`Invalid date in budget row ${rowNumber}`);
        }
      });
      
      // Display results
      if (issues.length === 0) {
        SpreadsheetApp.getUi().alert(
          'Data Integrity Check',
          'All data integrity checks passed successfully!',
          SpreadsheetApp.getUi().ButtonSet.OK
        );
      } else {
        const issueList = issues.map((issue, index) => `${index + 1}. ${issue}`).join('\n');
        
        const html = HtmlService.createHtmlOutput(`
          <div style="font-family: Arial, sans-serif; padding: 20px;">
            <h3>⚠️ Data Integrity Issues Found</h3>
            <p>The following issues were detected:</p>
            <div style="background: #fff3cd; padding: 15px; border-radius: 8px; margin: 15px 0;">
              <pre style="white-space: pre-wrap; font-family: Arial, sans-serif;">${issueList}</pre>
            </div>
            <p><strong>Recommendations:</strong></p>
            <ul>
              <li>Review and correct the flagged data</li>
              <li>Run this check regularly to maintain data quality</li>
              <li>Consider implementing data validation rules</li>
            </ul>
          </div>
        `)
          .setWidth(600)
          .setHeight(400);
        
        SpreadsheetApp.getUi().showModalDialog(html, 'Data Integrity Report');
      }
      
      return issues;
      
    } catch (error) {
      handleSystemError(error, 'validateDataIntegrity');
    }
  },
  
  /**
   * Get system statistics
   */
  getSystemStats: function() {
    try {
      const stats = {
        tenants: {
          total: 0,
          occupied: 0,
          vacant: 0,
          maintenance: 0
        },
        guestRooms: {
          total: 0,
          available: 0,
          occupied: 0,
          maintenance: 0
        },
        financial: {
          monthlyIncome: 0,
          yearlyIncome: 0,
          totalTransactions: 0
        },
        maintenance: {
          open: 0,
          inProgress: 0,
          completed: 0
        }
      };
      
      // Tenant statistics
      const tenantData = SheetManager.getAllData(CONFIG.SHEETS.TENANTS);
      tenantData.forEach(tenant => {
        if (tenant[0]) { // Has room number
          stats.tenants.total++;
          
          const status = tenant[8]; // Room status column
          switch (status) {
            case CONFIG.STATUS.ROOM.OCCUPIED:
              stats.tenants.occupied++;
              break;
            case CONFIG.STATUS.ROOM.VACANT:
              stats.tenants.vacant++;
              break;
            case CONFIG.STATUS.ROOM.MAINTENANCE:
              stats.tenants.maintenance++;
              break;
          }
        }
      });
      
      // Guest room statistics
      const guestRoomData = SheetManager.getAllData(CONFIG.SHEETS.GUEST_ROOMS);
      guestRoomData.forEach(room => {
        if (room[1]) { // Has room number
          stats.guestRooms.total++;

          const status = room[9]; // Room status column
          switch (status) {
            case 'Available':
              stats.guestRooms.available++;
              break;
            case 'Occupied':
              stats.guestRooms.occupied++;
              break;
            case 'Maintenance':
              stats.guestRooms.maintenance++;
              break;
          }
        }
      });
      
      // Financial statistics
      const budgetData = SheetManager.getAllData(CONFIG.SHEETS.BUDGET);
      const currentDate = new Date();
      const monthStart = new Date(currentDate.getFullYear(), currentDate.getMonth(), 1);
      const yearStart = new Date(currentDate.getFullYear(), 0, 1);
      
      budgetData.forEach(entry => {
        const entryDate = new Date(entry[0]);
        const amount = entry[3];
        
        stats.financial.totalTransactions++;
        
        if (amount > 0) { // Income only
          if (entryDate >= yearStart) {
            stats.financial.yearlyIncome += amount;
          }
          if (entryDate >= monthStart) {
            stats.financial.monthlyIncome += amount;
          }
        }
      });
      
      // Maintenance statistics
      const maintenanceData = SheetManager.getAllData(CONFIG.SHEETS.MAINTENANCE);
      maintenanceData.forEach(request => {
        const status = request[9]; // Status column
        
        switch (status) {
          case CONFIG.STATUS.MAINTENANCE.OPEN:
            stats.maintenance.open++;
            break;
          case CONFIG.STATUS.MAINTENANCE.IN_PROGRESS:
            stats.maintenance.inProgress++;
            break;
          case CONFIG.STATUS.MAINTENANCE.COMPLETED:
            stats.maintenance.completed++;
            break;
        }
      });
      
      return stats;
      
    } catch (error) {
      Logger.log(`Error getting system stats: ${error.toString()}`);
      return null;
    }
  }
};

Logger.log('DataManager module loaded successfully');
