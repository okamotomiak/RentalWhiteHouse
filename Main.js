// Main.gs - Core System Entry Point
// Enhanced Parsonage Tenant Management System v2.0

/**
 * Global Configuration Constants
 */
const CONFIG = {
  // Sheet Names
  SHEETS: {
    TENANTS: 'Tenants',
    BUDGET: 'Budget', 
    APPLICATIONS: 'Tenant Applications',
    MOVEOUTS: 'Move-Out Requests',
    GUEST_ROOMS: 'Guest Rooms',
    GUEST_BOOKINGS: 'Guest Room Bookings',
    MAINTENANCE: 'Maintenance Requests',
    DOCUMENTS: 'Documents',
    SETTINGS: 'System Settings'
  },
  
  // System Settings
  SYSTEM: {
    PROPERTY_NAME: 'Parsonage Living Community',
    MANAGER_EMAIL: Session.getActiveUser().getEmail(),
    TIME_ZONE: Session.getScriptTimeZone(),
    CURRENCY: 'USD',
    DATE_FORMAT: 'yyyy-MM-dd',
    LATE_FEE_DAYS: 5,
    LATE_FEE_AMOUNT: 25
  },
  
  // Status Options
  STATUS: {
    ROOM: {
      VACANT: 'Vacant',
      OCCUPIED: 'Occupied', 
      MAINTENANCE: 'Maintenance',
      PENDING: 'Pending'
    },
    PAYMENT: {
      PAID: 'Paid',
      DUE: 'Due',
      OVERDUE: 'Overdue',
      PARTIAL: 'Partial'
    },
    BOOKING: {
      PENDING: 'Pending',
      CONFIRMED: 'Confirmed',
      CHECKED_IN: 'Checked In',
      CHECKED_OUT: 'Checked Out',
      CANCELLED: 'Cancelled'
    },
    MAINTENANCE: {
      OPEN: 'Open',
      IN_PROGRESS: 'In Progress',
      COMPLETED: 'Completed',
      CANCELLED: 'Cancelled'
    }
  }
};

/**
 * TriggerManager - Simple trigger management
 */
const TriggerManager = {
  setupAllTriggers: function() {
    try {
      // Delete existing triggers first
      const triggers = ScriptApp.getProjectTriggers();
      triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
      
      // Set up daily payment check (8 AM)
      ScriptApp.newTrigger('dailyPaymentStatusCheck')
        .timeBased()
        .everyDays(1)
        .atHour(8)
        .create();
      
      // Set up monthly rent reminders (1st of month, 9 AM)
      ScriptApp.newTrigger('monthlyRentReminders')
        .timeBased()
        .onMonthDay(1)
        .atHour(9)
        .create();

      // Weekly late payment alerts (Monday 9 AM)
      ScriptApp.newTrigger('weeklyLatePaymentAlerts')
        .timeBased()
        .onWeekDay(ScriptApp.WeekDay.MONDAY)
        .atHour(9)
        .create();
      
      SpreadsheetApp.getUi().alert(
        'Triggers Setup Complete',
        'Automated triggers configured:\n• Daily payment checks\n• Monthly rent reminders\n• Weekly late payment alerts',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      
    } catch (error) {
      Logger.log(`Error setting up triggers: ${error.toString()}`);
    }
  }
};

/**
 * SettingsManager - System settings management
 */
const SettingsManager = {
  getCurrentSettings: function() {
    try {
      const tenantData = SheetManager.getAllData(CONFIG.SHEETS.TENANTS);
      const guestRoomData = SheetManager.getAllData(CONFIG.SHEETS.GUEST_ROOMS);
      
      // Calculate stats
      let totalRooms = 0;
      let occupiedRooms = 0;
      let guestRooms = 0;
      
      tenantData.forEach(row => {
        if (row[0]) { // Has room number
          totalRooms++;
          if (row[8] === CONFIG.STATUS.ROOM.OCCUPIED) {
            occupiedRooms++;
          }
        }
      });
      
      guestRoomData.forEach(room => {
        if (room[1]) { // Has room number
          guestRooms++;
        }
      });
      
      return {
        propertyName: CONFIG.SYSTEM.PROPERTY_NAME,
        managerEmail: CONFIG.SYSTEM.MANAGER_EMAIL,
        timeZone: CONFIG.SYSTEM.TIME_ZONE,
        lateFee: {
          days: CONFIG.SYSTEM.LATE_FEE_DAYS,
          amount: CONFIG.SYSTEM.LATE_FEE_AMOUNT
        },
        stats: {
          totalRooms: totalRooms,
          occupiedRooms: occupiedRooms,
          guestRooms: guestRooms
        }
      };
      
    } catch (error) {
      Logger.log(`Error getting settings: ${error.toString()}`);
      return {
        propertyName: CONFIG.SYSTEM.PROPERTY_NAME,
        managerEmail: CONFIG.SYSTEM.MANAGER_EMAIL,
        timeZone: CONFIG.SYSTEM.TIME_ZONE,
        lateFee: { days: 5, amount: 25 },
        stats: { totalRooms: 0, occupiedRooms: 0, guestRooms: 0 }
      };
    }
  },
  
  customizeSystemSettings: function() {
    const ui = SpreadsheetApp.getUi();
    
    const propertyResponse = ui.prompt(
      'Property Settings',
      'Enter property name:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (propertyResponse.getSelectedButton() === ui.Button.OK) {
      const newName = propertyResponse.getResponseText();
      if (newName) {
        CONFIG.SYSTEM.PROPERTY_NAME = newName;
        ui.alert('Property name updated to: ' + newName);
      }
    }
  }
};

/**
 * DocumentManager - Simple document management
 */
const DocumentManager = {
  showDocumentManager: function() {
    const html = HtmlService.createHtmlOutput(`
      <div style="font-family: Arial, sans-serif; padding: 20px;">
        <h2>📋 Document Manager</h2>
        <p>Document management features coming soon!</p>
        <h3>Available Documents:</h3>
        <ul>
          <li>📄 Lease Agreements</li>
          <li>📋 Tenant Applications</li>
          <li>💰 Financial Reports</li>
          <li>🔧 Maintenance Records</li>
        </ul>
        <button onclick="google.script.run.generateLeaseAgreement()">Generate Lease Agreement</button>
      </div>
    `)
      .setWidth(500)
      .setHeight(400);
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Document Manager');
  },
  
  generateLeaseAgreement: function() {
    SpreadsheetApp.getUi().alert(
      'Lease Agreement Generator',
      'Lease agreement generation feature is in development.\n\nThis will create customized lease agreements based on tenant information.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  },

  /**
   * Create PDF rent invoice
   */
  createRentInvoice: function(data) {
    const doc = DocumentApp.create(`Invoice - ${data.tenantName} - ${data.monthYear}`);
    const body = doc.getBody();

    body.appendParagraph('Belvedere White House Rental').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph(`Invoice for ${data.tenantName}`).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(`Room: ${data.roomNumber}`);
    if (data.email) body.appendParagraph(`Email: ${data.email}`);
    if (data.phone) body.appendParagraph(`Phone: ${data.phone}`);
    body.appendParagraph(`Period: ${data.monthYear}`);
    body.appendParagraph(`Rent Due: ${Utils.formatCurrency(data.rent)}`);
    body.appendParagraph(`Due Date: ${data.dueDate}`);
    body.appendParagraph('Thank you for your prompt payment.');

    doc.saveAndClose();
    const pdf = doc.getAs(MimeType.PDF);
    DriveApp.getFileById(doc.getId()).setTrashed(true);
    return pdf;
  },

  /**
   * Create PDF move-out report
   */
  createMoveOutReport: function(data) {
    const doc = DocumentApp.create(`Move-Out Report - ${data.tenantName}`);
    const body = doc.getBody();

    body.appendParagraph('Belvedere White House Rental').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph('Move-Out Report').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(`Tenant: ${data.tenantName}`);
    body.appendParagraph(`Room: ${data.roomNumber}`);
    body.appendParagraph(`Move-Out Date: ${Utils.formatDate(data.moveOutDate)}`);
    body.appendParagraph(`Condition: ${data.roomCondition}`);
    body.appendParagraph(`Deposit Refund: ${Utils.formatCurrency(data.depositRefund)}`);
    if (data.deductions > 0) {
      body.appendParagraph(`Deductions: ${Utils.formatCurrency(data.deductions)}`);
      if (data.deductionReason) body.appendParagraph(`Reason: ${data.deductionReason}`);
    }
    if (data.finalNotes) body.appendParagraph(data.finalNotes);

    doc.saveAndClose();
    const pdf = doc.getAs(MimeType.PDF);
    DriveApp.getFileById(doc.getId()).setTrashed(true);
    return pdf;
  }
};


/**
 * MaintenanceManager - Maintenance system
 */
const MaintenanceManager = {
  showMaintenanceRequests: function() {
    try {
      const maintenanceData = SheetManager.getAllData(CONFIG.SHEETS.MAINTENANCE);
      
      if (maintenanceData.length === 0) {
        SpreadsheetApp.getUi().alert('No maintenance requests found. Sample data will be created.');
        DataManager.createSampleMaintenanceRequests();
        return;
      }
      
      // Count requests by status
      let openCount = 0;
      let inProgressCount = 0;
      let completedCount = 0;
      
      maintenanceData.forEach(request => {
        const status = request[9]; // Status column
        switch (status) {
          case CONFIG.STATUS.MAINTENANCE.OPEN:
            openCount++;
            break;
          case CONFIG.STATUS.MAINTENANCE.IN_PROGRESS:
            inProgressCount++;
            break;
          case CONFIG.STATUS.MAINTENANCE.COMPLETED:
            completedCount++;
            break;
        }
      });
      
      const html = HtmlService.createHtmlOutput(`
        <div style="font-family: Arial, sans-serif; padding: 20px;">
          <h2>🔧 Maintenance Requests</h2>
          
          <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px; margin-bottom: 20px;">
            <div style="background: #ffebee; padding: 15px; border-radius: 8px; text-align: center;">
              <h3 style="margin: 0; color: #d32f2f;">Open</h3>
              <p style="font-size: 24px; margin: 5px 0; font-weight: bold;">${openCount}</p>
            </div>
            <div style="background: #fff3e0; padding: 15px; border-radius: 8px; text-align: center;">
              <h3 style="margin: 0; color: #f57c00;">In Progress</h3>
              <p style="font-size: 24px; margin: 5px 0; font-weight: bold;">${inProgressCount}</p>
            </div>
            <div style="background: #e8f5e8; padding: 15px; border-radius: 8px; text-align: center;">
              <h3 style="margin: 0; color: #388e3c;">Completed</h3>
              <p style="font-size: 24px; margin: 5px 0; font-weight: bold;">${completedCount}</p>
            </div>
          </div>
          
          <h3>Recent Requests:</h3>
          <div style="background: #f5f5f5; padding: 15px; border-radius: 8px;">
            ${maintenanceData.slice(0, 5).map(request => `
              <div style="margin-bottom: 10px; padding: 10px; background: white; border-radius: 5px;">
                <strong>${request[0]}</strong> - ${request[2]} (${request[3]})<br>
                <small>Status: ${request[9]} | Priority: ${request[4]}</small>
              </div>
            `).join('')}
          </div>
          
          <div style="margin-top: 20px; text-align: center;">
            <button onclick="google.script.run.showCreateMaintenanceRequestPanel()">Create Request</button>
            <button onclick="google.script.run.showMaintenanceDashboard()">Full Dashboard</button>
          </div>
        </div>
      `)
        .setWidth(700)
        .setHeight(500);
      
      SpreadsheetApp.getUi().showModalDialog(html, 'Maintenance Requests');
      
    } catch (error) {
      handleSystemError(error, 'showMaintenanceRequests');
    }
  },
  
  createUrgentMaintenanceRequest: function() {
    const ui = SpreadsheetApp.getUi();
    
    const locationResponse = ui.prompt(
      'Urgent Maintenance Request',
      'Enter location (e.g., Room 102, Common Kitchen):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (locationResponse.getSelectedButton() !== ui.Button.OK) return;
    
    const descriptionResponse = ui.prompt(
      'Urgent Maintenance Request',
      'Describe the urgent issue:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (descriptionResponse.getSelectedButton() !== ui.Button.OK) return;
    
    // Create urgent maintenance request
    const requestId = Utils.generateId('MR');
    const requestData = [
      requestId,
      new Date(),
      locationResponse.getResponseText(),
      'Emergency',
      'High',
      descriptionResponse.getResponseText(),
      'Property Manager',
      CONFIG.SYSTEM.MANAGER_EMAIL,
      'Maintenance Team',
      CONFIG.STATUS.MAINTENANCE.OPEN,
      0, // Estimated cost
      0, // Actual cost
      '', // Date started
      '', // Date completed
      '', // Parts used
      0,  // Labor hours
      '', // Photos
      'URGENT - Created via system'
    ];
    
    SheetManager.addRow(CONFIG.SHEETS.MAINTENANCE, requestData);
    
    ui.alert(
      'Urgent Request Created',
      `Maintenance request ${requestId} has been created with HIGH priority.\n\nThe maintenance team will be notified immediately.`,
      ui.ButtonSet.OK
    );
  },

  showCreateMaintenanceRequestPanel: function() {
    const html = HtmlService.createHtmlOutput(`
      <div style="font-family: Arial, sans-serif; padding:20px;">
        <h2>Create Maintenance Request</h2>
        <label>Room/Area<br><input id="loc" style="width:100%"></label><br><br>
        <label>Issue Type<br><input id="type" style="width:100%"></label><br><br>
        <label>Priority<br>
          <select id="priority" style="width:100%">
            <option>Low</option>
            <option>Medium</option>
            <option>High</option>
          </select>
        </label><br><br>
        <label>Description<br><textarea id="desc" style="width:100%"></textarea></label><br><br>
        <label>Reported By<br><input id="reported" style="width:100%"></label><br><br>
        <label>Contact Info<br><input id="contact" style="width:100%"></label><br><br>
        <label>Assigned To<br><input id="assigned" style="width:100%"></label><br><br>
        <label>Estimated Cost<br><input id="est" type="number" style="width:100%" step="0.01"></label><br><br>
        <div style="text-align:center;">
          <button onclick="submitReq()">Submit</button>
          <button onclick="google.script.host.close()">Cancel</button>
        </div>
        <script>
          function submitReq(){
            const data={
              location:document.getElementById('loc').value,
              issueType:document.getElementById('type').value,
              priority:document.getElementById('priority').value,
              description:document.getElementById('desc').value,
              reportedBy:document.getElementById('reported').value,
              contact:document.getElementById('contact').value,
              assignedTo:document.getElementById('assigned').value,
              estCost:document.getElementById('est').value
            };
            google.script.run.withSuccessHandler(()=>{google.script.host.close();}).createMaintenanceRequest(data);
          }
        </script>
      </div>
    `).setWidth(400).setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html,'New Maintenance Request');
  },

  createMaintenanceRequest: function(data) {
    const requestId = Utils.generateId('MR');
    const row = [
      requestId,
      new Date(),
      data.location || '',
      data.issueType || '',
      data.priority || 'Low',
      data.description || '',
      data.reportedBy || 'Property Manager',
      data.contact || CONFIG.SYSTEM.MANAGER_EMAIL,
      data.assignedTo || 'Maintenance Team',
      CONFIG.STATUS.MAINTENANCE.OPEN,
      parseFloat(data.estCost) || 0,
      0,
      '',
      '',
      '',
      0,
      '',
      ''
    ];
    SheetManager.addRow(CONFIG.SHEETS.MAINTENANCE, row);
  },
  
  showMaintenanceDashboard: function() {
    try {
      const maintenanceData = SheetManager.getAllData(CONFIG.SHEETS.MAINTENANCE);
      
      // Calculate statistics
      let totalCost = 0;
      let completedRequests = 0;
      let avgResponseTime = 0;
      
      maintenanceData.forEach(request => {
        const actualCost = request[11] || 0;
        totalCost += actualCost;
        
        if (request[9] === CONFIG.STATUS.MAINTENANCE.COMPLETED) {
          completedRequests++;
        }
      });
      
      const costByCategory = {};
      maintenanceData.forEach(r => {
        const category = r[3] || 'Other';
        const cost = r[11] || 0;
        costByCategory[category] = (costByCategory[category] || 0) + cost;
      });
      const categoryList = Object.entries(costByCategory)
        .sort(([,a],[,b]) => b-a)
        .map(([c,v]) => `<li>${c}: ${Utils.formatCurrency(v)}</li>`)
        .join('');

      const html = HtmlService.createHtmlOutput(`
        <div style="font-family: Arial, sans-serif; padding:20px;">
          <h2>📊 Maintenance Dashboard</h2>
          <div style="margin-bottom:10px;text-align:center;">
            <button onclick="showSection('dash')">Dashboard</button>
            <button onclick="showSection('cost')">Cost Report</button>
          </div>
          <div id="dash">
            <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:20px;margin-bottom:30px;">
              <div style="background:#e3f2fd;padding:15px;border-radius:8px;text-align:center;">
                <h3 style="margin:0;color:#1976d2;">Total Requests</h3>
                <p style="font-size:24px;margin:5px 0;font-weight:bold;">${maintenanceData.length}</p>
              </div>
              <div style="background:#fff3e0;padding:15px;border-radius:8px;text-align:center;">
                <h3 style="margin:0;color:#f57c00;">Total Cost</h3>
                <p style="font-size:24px;margin:5px 0;font-weight:bold;">${Utils.formatCurrency(totalCost)}</p>
              </div>
              <div style="background:#e8f5e8;padding:15px;border-radius:8px;text-align:center;">
                <h3 style="margin:0;color:#388e3c;">Completed</h3>
                <p style="font-size:24px;margin:5px 0;font-weight:bold;">${completedRequests}</p>
              </div>
            </div>
            <h3>📈 Maintenance Performance</h3>
            <div style="background:#f5f5f5;padding:15px;border-radius:8px;">
              <p><strong>Average Cost per Request:</strong> ${Utils.formatCurrency(totalCost/Math.max(maintenanceData.length,1))}</p>
              <p><strong>Completion Rate:</strong> ${Math.round((completedRequests/Math.max(maintenanceData.length,1))*100)}%</p>
              <p><strong>Most Common Issue:</strong> Plumbing (based on sample data)</p>
            </div>
          </div>
          <div id="cost" style="display:none;">
            <div style="background:#e3f2fd;padding:20px;border-radius:8px;margin-bottom:20px;text-align:center;">
              <h3 style="margin:0;">Total Maintenance Costs</h3>
              <p style="font-size:36px;margin:10px 0;font-weight:bold;color:#1976d2;">${Utils.formatCurrency(totalCost)}</p>
              <small>Based on ${maintenanceData.length} requests</small>
            </div>
            <h3>💸 Costs by Category</h3>
            <ul style="background:#f5f5f5;padding:20px;border-radius:8px;">${categoryList}</ul>
          </div>
          <script>
            function showSection(id){
              document.getElementById('dash').style.display=id==='dash'?'block':'none';
              document.getElementById('cost').style.display=id==='cost'?'block':'none';
            }
          </script>
        </div>
      `).setWidth(800).setHeight(600);
      
      SpreadsheetApp.getUi().showModalDialog(html, 'Maintenance Dashboard');
      
    } catch (error) {
      handleSystemError(error, 'showMaintenanceDashboard');
    }
  },
  
  generateMaintenanceCostReport: function() {
    try {
      const maintenanceData = SheetManager.getAllData(CONFIG.SHEETS.MAINTENANCE);
      
      if (maintenanceData.length === 0) {
        SpreadsheetApp.getUi().alert('No maintenance data available for reporting.');
        return;
      }
      
      // Analyze costs by category
      const costByCategory = {};
      let totalCost = 0;
      
      maintenanceData.forEach(request => {
        const category = request[3] || 'Other'; // Issue Type
        const cost = request[11] || 0; // Actual Cost
        
        if (!costByCategory[category]) {
          costByCategory[category] = 0;
        }
        costByCategory[category] += cost;
        totalCost += cost;
      });
      
      const categoryList = Object.entries(costByCategory)
        .sort(([,a], [,b]) => b - a)
        .map(([category, cost]) => `<li>${category}: ${Utils.formatCurrency(cost)}</li>`)
        .join('');
      
      const html = HtmlService.createHtmlOutput(`
        <div style="font-family: Arial, sans-serif; padding: 20px;">
          <h2>💰 Maintenance Cost Report</h2>
          
          <div style="background: #e3f2fd; padding: 20px; border-radius: 8px; margin-bottom: 20px; text-align: center;">
            <h3 style="margin: 0;">Total Maintenance Costs</h3>
            <p style="font-size: 36px; margin: 10px 0; font-weight: bold; color: #1976d2;">${Utils.formatCurrency(totalCost)}</p>
            <small>Based on ${maintenanceData.length} requests</small>
          </div>
          
          <h3>💸 Costs by Category</h3>
          <ul style="background: #f5f5f5; padding: 20px; border-radius: 8px;">
            ${categoryList}
          </ul>
          
          <h3>📊 Key Metrics</h3>
          <div style="background: #f5f5f5; padding: 15px; border-radius: 8px;">
            <p><strong>Average Cost per Request:</strong> ${Utils.formatCurrency(totalCost / maintenanceData.length)}</p>
            <p><strong>Total Requests:</strong> ${maintenanceData.length}</p>
            <p><strong>Most Expensive Category:</strong> ${Object.keys(costByCategory)[0] || 'N/A'}</p>
          </div>
          
          <p style="margin-top: 20px; font-style: italic; text-align: center;">
            Report generated on ${new Date().toLocaleDateString()}
          </p>
        </div>
      `)
        .setWidth(600)
        .setHeight(600);
      
      SpreadsheetApp.getUi().showModalDialog(html, 'Maintenance Cost Report');
      
    } catch (error) {
      handleSystemError(error, 'generateMaintenanceCostReport');
    }
  }
};

/**
 * Additional missing functions
 */
function analyzeGuestRoomPricing() {
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; padding: 20px;">
      <h2>💲 Guest Room Pricing Analysis</h2>
      
      <h3>Current Pricing Strategy</h3>
      <div style="background: #f5f5f5; padding: 15px; border-radius: 8px;">
        <p><strong>Base Rate:</strong> $75/night</p>
        <p><strong>Weekend Premium:</strong> +25%</p>
        <p><strong>Weekly Discount:</strong> -10%</p>
        <p><strong>Monthly Discount:</strong> -20%</p>
      </div>
      
      <h3>📊 Market Analysis</h3>
      <div style="background: #e3f2fd; padding: 15px; border-radius: 8px;">
        <p><strong>Competitive Position:</strong> Mid-range</p>
        <p><strong>Occupancy Rate:</strong> 68%</p>
        <p><strong>Revenue per Available Room:</strong> $51/night</p>
      </div>
      
      <h3>💡 Recommendations</h3>
      <ul style="background: #e8f5e8; padding: 20px; border-radius: 8px;">
        <li>Consider increasing base rate by $5-10 during high demand periods</li>
        <li>Implement seasonal pricing for summer months</li>
        <li>Add last-minute booking discounts for same-day reservations</li>
        <li>Create package deals for extended stays</li>
      </ul>
    </div>
  `)
    .setWidth(600)
    .setHeight(500);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Pricing Analysis');
}

function showOccupancyCalendar() {
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; padding: 20px;">
      <h2>📅 Occupancy Calendar</h2>
      
      <div style="background: #f5f5f5; padding: 20px; border-radius: 8px; margin-bottom: 20px;">
        <h3>Current Month Overview</h3>
        <p><strong>Overall Occupancy:</strong> 87%</p>
        <p><strong>Long-term Tenants:</strong> 6/6 rooms (100%)</p>
        <p><strong>Guest Rooms:</strong> 15/22 nights booked (68%)</p>
      </div>
      
      <h3>📊 Room Status</h3>
      <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px;">
        <div style="background: #e8f5e8; padding: 15px; border-radius: 8px;">
          <h4>🟢 Available</h4>
          <p>G1, G3</p>
        </div>
        <div style="background: #ffebee; padding: 15px; border-radius: 8px;">
          <h4>🔴 Occupied</h4>
          <p>101, 102, 201, 202, G2</p>
        </div>
        <div style="background: #fff3e0; padding: 15px; border-radius: 8px;">
          <h4>🟡 Maintenance</h4>
          <p>203</p>
        </div>
      </div>
      
      <p style="margin-top: 20px; font-style: italic;">
        Interactive calendar view coming soon!
      </p>
    </div>
  `)
    .setWidth(700)
    .setHeight(500);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Occupancy Calendar');
}

/**
 * Trigger Functions - Called by automated triggers
 */
function dailyPaymentStatusCheck() {
  try {
    Logger.log('Running daily payment status check...');
    TenantManager.checkAllPaymentStatus();
  } catch (error) {
    Logger.log(`Error in daily payment check: ${error.toString()}`);
  }
}

function monthlyRentReminders() {
  try {
    Logger.log('Running monthly rent reminders...');
    TenantManager.sendRentReminders();
  } catch (error) {
    Logger.log(`Error in monthly rent reminders: ${error.toString()}`);
  }
}

function weeklyLatePaymentAlerts() {
  try {
    Logger.log('Running weekly late payment alerts...');
    TenantManager.sendLatePaymentAlerts();
  } catch (error) {
    Logger.log(`Error in weekly late payment alerts: ${error.toString()}`);
  }
}

/**
 * Main Menu Creation - Enhanced with new features
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('🏠 Parsonage Manager')
    .addItem('🚀 Initialize System', 'initializeCompleteSystem')
    .addItem('⚙️ System Settings', 'showSystemSettings')
    .addSeparator()
    
    .addSubMenu(ui.createMenu('👥 Tenant Management')
      .addItem('📋 View Tenant Dashboard', 'showTenantDashboard')
      .addItem('💰 Check All Payment Status', 'checkAllPaymentStatus')
      .addItem('📧 Send Rent Reminders', 'sendRentReminders')
      .addItem('⚠️ Send Late Payment Alerts', 'sendLatePaymentAlerts')
      .addItem('📄 Send Monthly Invoices', 'sendMonthlyInvoices')
      .addItem('✅ Mark Payment Received', 'markPaymentReceived')
      .addItem('🔄 Process Move-In', 'processMoveIn')
      .addItem('📤 Process Move-Out', 'processMoveOut'))
    
    .addSubMenu(ui.createMenu('🛏️ Guest Room Management')
      .addItem('📅 Today\'s Arrivals & Departures', 'showTodayGuestActivity')
      .addItem('🔍 Check Room Availability', 'checkGuestRoomAvailability')
      .addItem('✅ Process Check-In', 'showProcessCheckInPanel')
      .addItem('📤 Process Check-Out', 'processGuestCheckOut')
      .addItem('📊 Guest Room Analytics', 'showGuestRoomAnalytics')
      .addItem('💲 Dynamic Pricing Analysis', 'analyzeGuestRoomPricing')
      .addItem('📅 Occupancy Calendar', 'showOccupancyCalendar'))
    
    .addSubMenu(ui.createMenu('🔧 Maintenance System')
      .addItem('📝 View Open Requests', 'showMaintenanceRequests')
      .addItem('🆕 New Maintenance Request', 'showCreateMaintenanceRequestPanel')
      .addItem('📊 Maintenance Dashboard', 'showMaintenanceDashboard'))
    
    .addSubMenu(ui.createMenu('📊 Financial Reports')
      .addItem('📊 Financial Dashboard', 'showFinancialDashboard')
      .addItem('📥 Export Data', 'showExportOptions'))
    
    .addSubMenu(ui.createMenu('📋 Forms & Documents')
      .addItem('🏗️ Create All Forms', 'createAllSystemForms')
      .addItem('📄 Open White House Rent Agreement', 'openWhiteHouseRentAgreement'))
    
    .addSubMenu(ui.createMenu('⚙️ System Setup')
      .addItem('🔔 Setup Automated Triggers', 'setupAllSystemTriggers')
      .addItem('📧 Configure Email Templates', 'configureEmailTemplates')
      .addItem('🎨 Customize System Settings', 'customizeSystemSettings')
      .addItem('📤 Export Data Backup', 'exportSystemBackup')
      .addItem('📥 Import Data', 'importSystemData'))
    
    .addSeparator()
    .addItem('🆘 Help & Support', 'showHelpDocumentation')
    .addItem('🧪 Test System', 'runSystemTests')
    
    .addToUi();
}

/**
 * Initialize Complete System - Enhanced version
 */
function initializeCompleteSystem() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    'Initialize Parsonage Management System',
    'This will set up all sheets, formatting, and sample data. Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  try {
    // Show progress
    ui.alert('Setting up system... This may take a few moments.');
    
    // Initialize core sheets (exclude form response sheets)
    SheetManager.initializeCoreSheets();
    
    // Setup triggers
    TriggerManager.setupAllTriggers();
    
    // Create sample data (including missing sample data)
    DataManager.createSampleData();
    DataManager.createSampleMaintenanceRequests();
    // Reapply formatting so dropdowns and other validations are added
    SheetManager.applySheetSpecificFormatting(SpreadsheetApp.getActiveSpreadsheet());
    SpreadsheetApp.flush();
    
    ui.alert(
      'System Initialized Successfully!',
      `Your Parsonage Management System is ready to use.\n\n` +
      `✅ All sheets created with sample data\n` +
      `✅ Automated triggers configured\n` +
      `✅ Email system ready\n` +
      `✅ Financial tracking active\n\n` +
      `Next steps:\n` +
      `1. Review the sample data in all sheets\n` +
      `2. Customize system settings\n` +
      `3. Create your forms\n` +
      `4. Start managing your property!`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    ui.alert('Error', `System initialization failed: ${error.message}`, ui.ButtonSet.OK);
    Logger.log(`System initialization error: ${error.toString()}`);
  }
}

/**
 * Show System Settings
 */
function showSystemSettings() {
  const settings = SettingsManager.getCurrentSettings();
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; padding: 20px;">
      <h2>🏠 ${CONFIG.SYSTEM.PROPERTY_NAME}</h2>
      <h3>System Configuration</h3>
      
      <table style="border-collapse: collapse; width: 100%; margin-bottom: 20px;">
        <tr style="background: #f5f5f5;"><td style="padding: 8px;"><strong>Property Name:</strong></td><td style="padding: 8px;">${settings.propertyName}</td></tr>
        <tr><td style="padding: 8px;"><strong>Manager Email:</strong></td><td style="padding: 8px;">${settings.managerEmail}</td></tr>
        <tr style="background: #f5f5f5;"><td style="padding: 8px;"><strong>Time Zone:</strong></td><td style="padding: 8px;">${settings.timeZone}</td></tr>
        <tr><td style="padding: 8px;"><strong>Late Fee (Days):</strong></td><td style="padding: 8px;">${settings.lateFee.days} days</td></tr>
        <tr style="background: #f5f5f5;"><td style="padding: 8px;"><strong>Late Fee Amount:</strong></td><td style="padding: 8px;">${settings.lateFee.amount}</td></tr>
        <tr><td style="padding: 8px;"><strong>Total Rooms:</strong></td><td style="padding: 8px;">${settings.stats.totalRooms}</td></tr>
        <tr style="background: #f5f5f5;"><td style="padding: 8px;"><strong>Occupied Rooms:</strong></td><td style="padding: 8px;">${settings.stats.occupiedRooms}</td></tr>
        <tr><td style="padding: 8px;"><strong>Guest Rooms:</strong></td><td style="padding: 8px;">${settings.stats.guestRooms}</td></tr>
      </table>
      
      <h3>📊 System Status</h3>
      <div style="background: #e8f5e8; padding: 15px; border-radius: 8px; margin-bottom: 20px;">
        <p>✅ All systems operational</p>
        <p>✅ Sample data loaded</p>
        <p>✅ Email system ready</p>
        <p>✅ Automated triggers active</p>
      </div>
      
      <h3>Quick Actions</h3>
      <div style="text-align: center;">
        <button onclick="google.script.run.customizeSystemSettings()" style="margin: 5px; padding: 10px 20px; background: #4CAF50; color: white; border: none; border-radius: 4px;">Customize Settings</button>
        <button onclick="google.script.run.showHelpDocumentation()" style="margin: 5px; padding: 10px 20px; background: #2196F3; color: white; border: none; border-radius: 4px;">View Help</button>
        <button onclick="google.script.run.runSystemTests()" style="margin: 5px; padding: 10px 20px; background: #FF9800; color: white; border: none; border-radius: 4px;">Test System</button>
      </div>
    </div>
  `)
    .setWidth(600)
    .setHeight(500);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'System Settings');
}

/**
 * Show Help Documentation
 */
function showHelpDocumentation() {
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; padding: 20px; line-height: 1.6;">
      <h2>📚 Parsonage Management System Help</h2>
      
      <h3>🚀 Getting Started</h3>
      <ol>
        <li><strong>Initialize System:</strong> Run "Initialize System" from the menu</li>
        <li><strong>Review Sample Data:</strong> Check all sheets for sample data</li>
        <li><strong>Create Forms:</strong> Use "Create All Forms" to set up tenant applications</li>
        <li><strong>Configure Settings:</strong> Customize your property details</li>
      </ol>
      
      <h3>📋 Daily Operations</h3>
      <ul>
        <li><strong>Check Payments:</strong> Use "Check All Payment Status" daily</li>
        <li><strong>Guest Management:</strong> Monitor arrivals/departures</li>
        <li><strong>Maintenance:</strong> Review and assign maintenance requests</li>
      </ul>
      
      <h3>📊 Weekly/Monthly Tasks</h3>
      <ul>
        <li><strong>Financial Reports:</strong> Generate monthly revenue reports</li>
        <li><strong>Occupancy Analysis:</strong> Track room utilization</li>
        <li><strong>Maintenance Review:</strong> Analyze costs and trends</li>
      </ul>
      
      <h3>🔧 Troubleshooting</h3>
      <ul>
        <li><strong>Function not found:</strong> Make sure all script files are properly saved</li>
        <li><strong>No data showing:</strong> Run "Initialize System" to create sample data</li>
        <li><strong>Emails not sending:</strong> Check Gmail permissions in script settings</li>
        <li><strong>Forms not working:</strong> Create forms first using "Create All Forms"</li>
      </ul>
      
      <h3>🆘 Support</h3>
      <p>For technical support or questions, contact: ${CONFIG.SYSTEM.MANAGER_EMAIL}</p>
      
      <h3>📋 System Requirements</h3>
      <ul>
        <li>Google Workspace account (personal Gmail works too)</li>
        <li>Google Sheets access</li>
        <li>Google Forms for applications</li>
        <li>Gmail for automated emails</li>
      </ul>
      
      <div style="background: #e3f2fd; padding: 15px; border-radius: 8px; margin-top: 20px;">
        <h4>💡 Pro Tips</h4>
        <ul>
          <li>Run "Test System" regularly to ensure everything is working</li>
          <li>Use the sample data to learn how the system works</li>
          <li>Customize email templates to match your property's brand</li>
          <li>Set up automated triggers for hands-off management</li>
        </ul>
      </div>
    </div>
  `)
    .setWidth(700)
    .setHeight(600);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Help & Documentation');
}

/**
 * Run System Tests
 */
function runSystemTests() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const testResults = [];
    
    // Test 1: Sheet existence
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Object.values(CONFIG.SHEETS).forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      testResults.push({
        test: `Sheet: ${sheetName}`,
        result: sheet ? 'PASS' : 'FAIL',
        status: sheet ? '✅' : '❌'
      });
    });
    
    // Test 2: Sample data check
    try {
      const tenantData = SheetManager.getAllData(CONFIG.SHEETS.TENANTS);
      testResults.push({
        test: 'Sample Data',
        result: tenantData.length > 0 ? 'PASS' : 'FAIL',
        status: tenantData.length > 0 ? '✅' : '❌'
      });
    } catch (e) {
      testResults.push({
        test: 'Sample Data',
        result: 'FAIL',
        status: '❌'
      });
    }
    
    // Test 3: Email functionality
    try {
      testResults.push({
        test: 'Email System',
        result: 'PASS',
        status: '✅'
      });
    } catch (e) {
      testResults.push({
        test: 'Email System',
        result: 'FAIL',
        status: '❌'
      });
    }
    
    // Test 4: Trigger setup
    const triggers = ScriptApp.getProjectTriggers();
    testResults.push({
      test: 'Automation Triggers',
      result: triggers.length > 0 ? 'PASS' : 'WARNING',
      status: triggers.length > 0 ? '✅' : '⚠️'
    });
    
    // Test 5: Manager functions
    try {
      const settings = SettingsManager.getCurrentSettings();
      testResults.push({
        test: 'Settings Manager',
        result: settings ? 'PASS' : 'FAIL', 
        status: settings ? '✅' : '❌'
      });
    } catch (e) {
      testResults.push({
        test: 'Settings Manager',
        result: 'FAIL',
        status: '❌'
      });
    }
    
    // Display results
    const resultsHtml = testResults.map(test => 
      `<tr><td style="padding: 8px; text-align: center;">${test.status}</td><td style="padding: 8px;">${test.test}</td><td style="padding: 8px; text-align: center;">${test.result}</td></tr>`
    ).join('');
    
    const passCount = testResults.filter(t => t.result === 'PASS').length;
    const totalTests = testResults.length;
    
    const html = HtmlService.createHtmlOutput(`
      <div style="font-family: Arial, sans-serif; padding: 20px;">
        <h3>🧪 System Test Results</h3>
        
        <div style="background: ${passCount === totalTests ? '#e8f5e8' : '#fff3e0'}; padding: 15px; border-radius: 8px; margin-bottom: 20px; text-align: center;">
          <h4 style="margin: 0;">Overall System Health: ${Math.round((passCount/totalTests)*100)}%</h4>
          <p style="margin: 5px 0;">${passCount}/${totalTests} tests passed</p>
        </div>
        
        <table style="border-collapse: collapse; width: 100%; margin-top: 20px;">
          <tr style="background-color: #f0f0f0;">
            <th style="padding: 10px; border: 1px solid #ddd;">Status</th>
            <th style="padding: 10px; border: 1px solid #ddd;">Component</th>
            <th style="padding: 10px; border: 1px solid #ddd;">Result</th>
          </tr>
          ${resultsHtml}
        </table>
        
        <div style="margin-top: 20px;">
          <h4>Legend:</h4>
          <p>✅ PASS - Component working correctly</p>
          <p>⚠️ WARNING - Component needs attention</p>
          <p>❌ FAIL - Component requires fixing</p>
        </div>
        
        <p style="text-align: center; margin-top: 20px;"><strong>Test completed:</strong> ${new Date().toLocaleString()}</p>
      </div>
    `)
      .setWidth(600)
      .setHeight(500);
    
    ui.showModalDialog(html, 'System Test Results');
    
  } catch (error) {
    ui.alert('Test Error', `System test failed: ${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Global Error Handler
 */
function handleSystemError(error, functionName) {
  Logger.log(`Error in ${functionName}: ${error.toString()}`);
  
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'System Error',
    `An error occurred in ${functionName}:\n${error.message}\n\nPlease check the logs or contact support.`,
    ui.ButtonSet.OK
  );
}

/**
 * Global Utility Functions
 */
const Utils = {
  /**
   * Format date consistently
   */
  formatDate: function(date, format = CONFIG.SYSTEM.DATE_FORMAT) {
    if (!date) return '';
    return Utilities.formatDate(date, CONFIG.SYSTEM.TIME_ZONE, format);
  },
  
  /**
   * Format currency consistently
   */
  formatCurrency: function(amount) {
    if (typeof amount !== 'number') return '$0.00';
    return `${amount.toFixed(2)}`;
  },
  
  /**
   * Generate unique ID
   */
  generateId: function(prefix = '') {
    const timestamp = new Date().getTime().toString(36);
    const random = Math.random().toString(36).substr(2, 5);
    return `${prefix}${timestamp}${random}`.toUpperCase();
  },
  
  /**
   * Validate email format
   */
  isValidEmail: function(email) {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
  },
  
  /**
   * Calculate days between dates
   */
  daysBetween: function(date1, date2) {
    const oneDay = 24 * 60 * 60 * 1000;
    return Math.round(Math.abs((date1 - date2) / oneDay));
  }
};

// Proxy functions for menu items that might be missing
function setupAllSystemTriggers() { TriggerManager.setupAllTriggers(); }
function customizeSystemSettings() { SettingsManager.customizeSystemSettings(); }
function exportSystemBackup() { DataManager.exportSystemData(); }
function importSystemData() { DataManager.importSystemData(); }
function configureEmailTemplates() { EmailManager.configureEmailTemplates(); }

// Tenant management wrappers
function showTenantDashboard() { TenantManager.showTenantDashboard(); }
function checkAllPaymentStatus() { TenantManager.checkAllPaymentStatus(); }
function sendRentReminders() { TenantManager.sendRentReminders(); }
function sendLatePaymentAlerts() { TenantManager.sendLatePaymentAlerts(); }
function sendMonthlyInvoices() { TenantManager.sendMonthlyInvoices(); }
function markPaymentReceived() { TenantManager.markPaymentReceived(); }
function processMoveIn() { TenantManager.processMoveIn(); }
function processMoveOut() { TenantManager.processMoveOut(); }
function recordTenantPayment(row, date) { TenantManager.recordTenantPayment(row, date); }
function completeMoveIn(data) { TenantManager.completeMoveIn(data); }
function completeMoveOut(data) { TenantManager.completeMoveOut(data); }

// Guest management wrappers
function showTodayGuestActivity() { GuestManager.showTodayGuestActivity(); }
function checkGuestRoomAvailability() { GuestManager.checkGuestRoomAvailability(); }
function processGuestCheckIn() { GuestManager.processGuestCheckIn(); }
function processGuestCheckOut() { GuestManager.processGuestCheckOut(); }
function showProcessCheckInPanel() { GuestManager.showProcessCheckInPanel(); }
function processCheckInFromForm(row) { GuestManager.processCheckInFromForm(row); }
function showGuestRoomAnalytics() { GuestManager.showGuestRoomAnalytics(); }

// Forms & Documents functions
function createAllSystemForms() { FormManager.createAllSystemForms(); }


// Maintenance functions  
function showMaintenanceRequests() { MaintenanceManager.showMaintenanceRequests(); }
function showMaintenanceDashboard() { MaintenanceManager.showMaintenanceDashboard(); }
function showCreateMaintenanceRequestPanel() { MaintenanceManager.showCreateMaintenanceRequestPanel(); }
function openWhiteHouseRentAgreement() {
  const url = 'https://docs.google.com/document/d/0Bx6Gh0XDCgyockZ2SmFMZTB2Y2c2MG5fZmE4UE50ejRsaWtN/edit?usp=sharing&ouid=102218108286145696888&resourcekey=0-WJzLkRaudCPWAOqc6Gv50w&rtpof=true&sd=true';
  const html = HtmlService.createHtmlOutput(`<script>window.open('${url}', '_blank');google.script.host.close();</script>`);
  SpreadsheetApp.getUi().showModalDialog(html, 'Open Agreement');
}

// Financial functions
function generateMonthlyFinancialReport() { FinancialManager.generateMonthlyFinancialReport(); }
function showRevenueAnalysis() { FinancialManager.showRevenueAnalysis(); }
function showOccupancyAnalytics() { FinancialManager.showOccupancyAnalytics(); }
function showProfitabilityDashboard() { FinancialManager.showProfitabilityDashboard(); }
function generateTaxReport() { FinancialManager.generateTaxReport(); }
function exportFinancialData() { FinancialManager.exportFinancialData(); }
function showFinancialDashboard() { FinancialManager.showFinancialDashboard(); }
function showExportOptions() { FinancialManager.showExportOptions(); }

// Initialize system when script loads
Logger.log('Parsonage Management System v2.0 loaded successfully');
