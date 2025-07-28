// TenantManager.gs - Tenant Management System
// Handles all tenant-related operations

const TenantManager = {
  
  /**
   * Column indexes for Tenants sheet (1-based)
   */
  COL: {
    ROOM_NUMBER: 1,
    RENTAL_PRICE: 2,
    NEGOTIATED_PRICE: 3,
    TENANT_NAME: 4,
    TENANT_EMAIL: 5,
    TENANT_PHONE: 6,
    MOVE_IN_DATE: 7,
    SECURITY_DEPOSIT: 8,
    ROOM_STATUS: 9,
    LAST_PAYMENT: 10,
    PAYMENT_STATUS: 11,
    MOVE_OUT_PLANNED: 12,
    EMERGENCY_CONTACT: 13,
    LEASE_END_DATE: 14,
    NOTES: 15
  },
  
  /**
   * Check and update payment status for all tenants
   */
  checkAllPaymentStatus: function() {
    try {
      const sheet = SheetManager.getSheet(CONFIG.SHEETS.TENANTS);
      const data = SheetManager.getAllData(CONFIG.SHEETS.TENANTS);
      
      if (data.length === 0) {
        SpreadsheetApp.getUi().alert('No tenant data found.');
        return;
      }
      
      const today = new Date();
      const firstThisMonth = new Date(today.getFullYear(), today.getMonth(), 1);
      const firstLastMonth = new Date(today.getFullYear(), today.getMonth() - 1, 1);
      
      let updatedCount = 0;
      
      data.forEach((row, index) => {
        const rowNumber = index + 2; // Account for header row
        const roomStatus = row[this.COL.ROOM_STATUS - 1];
        
        if (roomStatus !== CONFIG.STATUS.ROOM.OCCUPIED) {
          // Clear payment status for non-occupied rooms
          sheet.getRange(rowNumber, this.COL.PAYMENT_STATUS).setValue('');
          return;
        }
        
        const lastPayment = row[this.COL.LAST_PAYMENT - 1];
        let status = CONFIG.STATUS.PAYMENT.OVERDUE;
        
        if (lastPayment instanceof Date) {
          if (lastPayment >= firstThisMonth) {
            status = CONFIG.STATUS.PAYMENT.PAID;
          } else if (lastPayment >= firstLastMonth) {
            status = CONFIG.STATUS.PAYMENT.DUE;
          }
        }
        
        // Update payment status
        sheet.getRange(rowNumber, this.COL.PAYMENT_STATUS).setValue(status);
        updatedCount++;
      });
      
      SpreadsheetApp.getUi().alert(
        'Payment Status Updated',
        `Updated payment status for ${updatedCount} tenants.`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      
    } catch (error) {
      handleSystemError(error, 'checkAllPaymentStatus');
    }
  },
  
  /**
   * Send rent reminders to tenants with due or overdue payments
   */
  sendRentReminders: function() {
    try {
      const data = SheetManager.getAllData(CONFIG.SHEETS.TENANTS);
      
      if (data.length === 0) {
        SpreadsheetApp.getUi().alert('No tenant data found.');
        return;
      }
      
      const monthYear = Utils.formatDate(new Date(), 'MMMM yyyy');
      let sentCount = 0;
      let failedCount = 0;
      
      data.forEach(row => {
        const paymentStatus = row[this.COL.PAYMENT_STATUS - 1];
        const email = row[this.COL.TENANT_EMAIL - 1];
        
        if ((paymentStatus === CONFIG.STATUS.PAYMENT.DUE || paymentStatus === CONFIG.STATUS.PAYMENT.OVERDUE) && email) {
          const tenantName = row[this.COL.TENANT_NAME - 1];
          const roomNumber = row[this.COL.ROOM_NUMBER - 1];
          const rent = row[this.COL.NEGOTIATED_PRICE - 1] || row[this.COL.RENTAL_PRICE - 1];
          
          const emailData = {
            tenantName: tenantName,
            roomNumber: roomNumber,
            rent: rent,
            status: paymentStatus,
            monthYear: monthYear,
            dueDate: this.calculateRentDueDate(),
            lateFee: paymentStatus === CONFIG.STATUS.PAYMENT.OVERDUE ? CONFIG.SYSTEM.LATE_FEE_AMOUNT : 0
          };
          
          try {
            EmailManager.sendRentReminder(email, emailData);
            sentCount++;
          } catch (emailError) {
            Logger.log(`Failed to send reminder to ${email}: ${emailError.message}`);
            failedCount++;
          }
        }
      });
      
      const message = `Rent reminders sent: ${sentCount}` + 
                     (failedCount > 0 ? `\nFailed to send: ${failedCount}` : '');
      
      SpreadsheetApp.getUi().alert('Rent Reminders', message, SpreadsheetApp.getUi().ButtonSet.OK);
      
    } catch (error) {
      handleSystemError(error, 'sendRentReminders');
    }
  },
  
  /**
   * Send late payment alerts to manager
   */
  sendLatePaymentAlerts: function() {
    try {
      const data = SheetManager.getAllData(CONFIG.SHEETS.TENANTS);
      
      if (data.length === 0) {
        return;
      }
      
      const overdueList = [];
      
      data.forEach(row => {
        if (row[this.COL.PAYMENT_STATUS - 1] === CONFIG.STATUS.PAYMENT.OVERDUE) {
          const tenant = row[this.COL.TENANT_NAME - 1];
          const email = row[this.COL.TENANT_EMAIL - 1];
          const room = row[this.COL.ROOM_NUMBER - 1];
          const lastPayment = row[this.COL.LAST_PAYMENT - 1];
          const lastPaymentStr = lastPayment ? Utils.formatDate(lastPayment) : 'Never';
          
          overdueList.push({
            tenant: tenant,
            email: email,
            room: room,
            lastPayment: lastPaymentStr
          });
        }
      });
      
      if (overdueList.length === 0) {
        SpreadsheetApp.getUi().alert('No overdue tenants found.');
        return;
      }
      
      EmailManager.sendLatePaymentAlert(CONFIG.SYSTEM.MANAGER_EMAIL, {
        overdueList: overdueList,
        count: overdueList.length
      });
      
      SpreadsheetApp.getUi().alert(
        'Late Payment Alert Sent',
        `Alert sent to manager about ${overdueList.length} overdue tenant(s).`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      
    } catch (error) {
      handleSystemError(error, 'sendLatePaymentAlerts');
    }
  },
  
  /**
   * Send monthly invoices to all occupied rooms
   */
  sendMonthlyInvoices: function() {
    try {
      const data = SheetManager.getAllData(CONFIG.SHEETS.TENANTS);
      
      if (data.length === 0) {
        SpreadsheetApp.getUi().alert('No tenant data found.');
        return;
      }
      
      const monthYear = Utils.formatDate(new Date(), 'MMMM yyyy');
      let sentCount = 0;
      let failedCount = 0;
      
      data.forEach(row => {
        const roomStatus = row[this.COL.ROOM_STATUS - 1];
        const email = row[this.COL.TENANT_EMAIL - 1];
        
        if (roomStatus === CONFIG.STATUS.ROOM.OCCUPIED && email) {
          const tenant = row[this.COL.TENANT_NAME - 1];
          const room = row[this.COL.ROOM_NUMBER - 1];
          const rent = row[this.COL.NEGOTIATED_PRICE - 1] || row[this.COL.RENTAL_PRICE - 1];
          const dueDate = this.calculateRentDueDate();
          
          try {
            const pdf = DocumentManager.createRentInvoice({
              tenantName: tenant,
              roomNumber: room,
              rent: rent,
              monthYear: monthYear,
              dueDate: dueDate,
              propertyName: CONFIG.SYSTEM.PROPERTY_NAME
            });
            
            EmailManager.sendMonthlyInvoice(email, {
              tenantName: tenant,
              monthYear: monthYear
            }, pdf);
            
            sentCount++;
          } catch (invoiceError) {
            Logger.log(`Failed to send invoice to ${email}: ${invoiceError.message}`);
            failedCount++;
          }
        }
      });
      
      const message = `Monthly invoices sent: ${sentCount}` + 
                     (failedCount > 0 ? `\nFailed to send: ${failedCount}` : '');
      
      SpreadsheetApp.getUi().alert('Monthly Invoices', message, SpreadsheetApp.getUi().ButtonSet.OK);
      
    } catch (error) {
      handleSystemError(error, 'sendMonthlyInvoices');
    }
  },
  
  /**
   * Mark payment as received for selected tenant
   */
  markPaymentReceived: function() {
    try {
      const ui = SpreadsheetApp.getUi();
      const sheet = SpreadsheetApp.getActiveSheet();
      
      if (sheet.getName() !== CONFIG.SHEETS.TENANTS) {
        ui.alert('Please select a row in the Tenants sheet.');
        return;
      }
      
      const row = sheet.getActiveRange().getRow();
      if (row <= 1) {
        ui.alert('Please select a tenant row.');
        return;
      }
      
      const tenantName = sheet.getRange(row, this.COL.TENANT_NAME).getValue();
      const roomNumber = sheet.getRange(row, this.COL.ROOM_NUMBER).getValue();
      const rent = sheet.getRange(row, this.COL.NEGOTIATED_PRICE).getValue() || 
                   sheet.getRange(row, this.COL.RENTAL_PRICE).getValue();
      
      if (!tenantName || !roomNumber) {
        ui.alert('Invalid tenant row selected.');
        return;
      }
      
      // Get payment amount from user
      const response = ui.prompt(
        'Record Payment',
        `Recording payment for ${tenantName} (Room ${roomNumber})\nEnter payment amount:`,
        ui.ButtonSet.OK_CANCEL
      );
      
      if (response.getSelectedButton() !== ui.Button.OK) {
        return;
      }
      
      const paymentAmount = parseFloat(response.getResponseText()) || rent;
      
      // Update tenant record
      sheet.getRange(row, this.COL.LAST_PAYMENT).setValue(new Date());
      
      if (paymentAmount >= rent) {
        sheet.getRange(row, this.COL.PAYMENT_STATUS).setValue(CONFIG.STATUS.PAYMENT.PAID);
      } else {
        sheet.getRange(row, this.COL.PAYMENT_STATUS).setValue(CONFIG.STATUS.PAYMENT.PARTIAL);
      }
      
      // Log payment in budget
      FinancialManager.logPayment({
        date: new Date(),
        type: 'Rent Income',
        description: `Rent payment from ${tenantName} - Room ${roomNumber}`,
        amount: paymentAmount,
        category: 'Rent',
        tenant: tenantName,
        reference: `RENT-${roomNumber}-${Utils.formatDate(new Date(), 'yyyyMM')}`
      });
      
      ui.alert(
        'Payment Recorded',
        `Payment of ${Utils.formatCurrency(paymentAmount)} from ${tenantName} has been recorded.`,
        ui.ButtonSet.OK
      );
      
    } catch (error) {
      handleSystemError(error, 'markPaymentReceived');
    }
  },
  
  /**
   * Process tenant move-in
   */
  processMoveIn: function() {
    try {
      const ui = SpreadsheetApp.getUi();
      
      // Get available vacant rooms
      const vacantRooms = this.getVacantRooms();
      
      if (vacantRooms.length === 0) {
        ui.alert('No vacant rooms available.');
        return;
      }
      
      // Show move-in form
      const html = HtmlService.createHtmlOutput(`
        <div style="font-family: Arial, sans-serif; padding: 20px;">
          <h3>Process Tenant Move-In</h3>
          <form id="moveInForm">
            <table>
              <tr>
                <td><label>Room Number:</label></td>
                <td>
                  <select name="roomNumber" required>
                    <option value="">Select Room</option>
                    ${vacantRooms.map(room => 
                      `<option value="${room.number}">Room ${room.number} - ${Utils.formatCurrency(room.rent)}/month</option>`
                    ).join('')}
                  </select>
                </td>
              </tr>
              <tr>
                <td><label>Tenant Name:</label></td>
                <td><input type="text" name="tenantName" required style="width: 200px;"></td>
              </tr>
              <tr>
                <td><label>Email:</label></td>
                <td><input type="email" name="email" required style="width: 200px;"></td>
              </tr>
              <tr>
                <td><label>Phone:</label></td>
                <td><input type="tel" name="phone" required style="width: 200px;"></td>
              </tr>
              <tr>
                <td><label>Move-in Date:</label></td>
                <td><input type="date" name="moveInDate" required></td>
              </tr>
              <tr>
                <td><label>Security Deposit:</label></td>
                <td><input type="number" name="securityDeposit" step="0.01" required style="width: 100px;"></td>
              </tr>
              <tr>
                <td><label>Negotiated Rent:</label></td>
                <td><input type="number" name="negotiatedRent" step="0.01" style="width: 100px;"> (optional)</td>
              </tr>
              <tr>
                <td><label>Emergency Contact:</label></td>
                <td><input type="text" name="emergencyContact" style="width: 300px;"></td>
              </tr>
              <tr>
                <td><label>Lease End Date:</label></td>
                <td><input type="date" name="leaseEndDate"></td>
              </tr>
              <tr>
                <td><label>Notes:</label></td>
                <td><textarea name="notes" rows="3" style="width: 300px;"></textarea></td>
              </tr>
            </table>
            <br>
            <button type="button" onclick="processMoveIn()">Process Move-In</button>
            <button type="button" onclick="google.script.host.close()">Cancel</button>
          </form>
        </div>
        
        <script>
          function processMoveIn() {
            const form = document.getElementById('moveInForm');
            const formData = new FormData(form);
            const data = Object.fromEntries(formData);
            
            google.script.run
              .withSuccessHandler(function(result) {
                alert('Move-in processed successfully!');
                google.script.host.close();
              })
              .withFailureHandler(function(error) {
                alert('Error: ' + error.message);
              })
              .completeMoveIn(data);
          }
        </script>
      `)
        .setWidth(500)
        .setHeight(600);
      
      ui.showModalDialog(html, 'Process Move-In');
      
    } catch (error) {
      handleSystemError(error, 'processMoveIn');
    }
  },
  
  /**
   * Complete move-in process (called from HTML form)
   */
  completeMoveIn: function(data) {
    try {
      const sheet = SheetManager.getSheet(CONFIG.SHEETS.TENANTS);
      
      // Find the room row
      const roomRows = SheetManager.findRows(CONFIG.SHEETS.TENANTS, this.COL.ROOM_NUMBER, data.roomNumber);
      
      if (roomRows.length === 0) {
        throw new Error('Room not found');
      }
      
      const roomRow = roomRows[0];
      const rowNumber = roomRow.rowNumber;
      
      // Update tenant information
      const moveInDate = new Date(data.moveInDate);
      const leaseEndDate = data.leaseEndDate ? new Date(data.leaseEndDate) : '';
      
      sheet.getRange(rowNumber, this.COL.NEGOTIATED_PRICE).setValue(data.negotiatedRent || '');
      sheet.getRange(rowNumber, this.COL.TENANT_NAME).setValue(data.tenantName);
      sheet.getRange(rowNumber, this.COL.TENANT_EMAIL).setValue(data.email);
      sheet.getRange(rowNumber, this.COL.TENANT_PHONE).setValue(data.phone);
      sheet.getRange(rowNumber, this.COL.MOVE_IN_DATE).setValue(moveInDate);
      sheet.getRange(rowNumber, this.COL.SECURITY_DEPOSIT).setValue(parseFloat(data.securityDeposit));
      sheet.getRange(rowNumber, this.COL.ROOM_STATUS).setValue(CONFIG.STATUS.ROOM.OCCUPIED);
      sheet.getRange(rowNumber, this.COL.PAYMENT_STATUS).setValue(CONFIG.STATUS.PAYMENT.DUE);
      sheet.getRange(rowNumber, this.COL.EMERGENCY_CONTACT).setValue(data.emergencyContact || '');
      sheet.getRange(rowNumber, this.COL.LEASE_END_DATE).setValue(leaseEndDate);
      sheet.getRange(rowNumber, this.COL.NOTES).setValue(data.notes || '');
      
      // Log security deposit payment
      if (data.securityDeposit > 0) {
        FinancialManager.logPayment({
          date: moveInDate,
          type: 'Security Deposit',
          description: `Security deposit from ${data.tenantName} - Room ${data.roomNumber}`,
          amount: parseFloat(data.securityDeposit),
          category: 'Deposit',
          tenant: data.tenantName,
          reference: `DEPOSIT-${data.roomNumber}-${Utils.formatDate(moveInDate, 'yyyyMMdd')}`
        });
      }
      
      // Send welcome email
      EmailManager.sendWelcomeEmail(data.email, {
        tenantName: data.tenantName,
        roomNumber: data.roomNumber,
        moveInDate: Utils.formatDate(moveInDate, 'MMMM dd, yyyy'),
        rent: data.negotiatedRent || roomRow.data[this.COL.RENTAL_PRICE - 1],
        propertyName: CONFIG.SYSTEM.PROPERTY_NAME
      });
      
      return { success: true, message: 'Move-in processed successfully' };
      
    } catch (error) {
      Logger.log(`Error in completeMoveIn: ${error.toString()}`);
      throw error;
    }
  },
  
  /**
   * Process tenant move-out
   */
  processMoveOut: function() {
    try {
      const ui = SpreadsheetApp.getUi();
      const sheet = SpreadsheetApp.getActiveSheet();
      
      if (sheet.getName() !== CONFIG.SHEETS.TENANTS) {
        ui.alert('Please select a tenant row in the Tenants sheet.');
        return;
      }
      
      const row = sheet.getActiveRange().getRow();
      if (row <= 1) {
        ui.alert('Please select a tenant row.');
        return;
      }
      
      const tenantName = sheet.getRange(row, this.COL.TENANT_NAME).getValue();
      const roomNumber = sheet.getRange(row, this.COL.ROOM_NUMBER).getValue();
      const securityDeposit = sheet.getRange(row, this.COL.SECURITY_DEPOSIT).getValue() || 0;
      
      if (!tenantName) {
        ui.alert('No tenant selected or invalid row.');
        return;
      }
      
      const response = ui.alert(
        'Process Move-Out',
        `Process move-out for ${tenantName} (Room ${roomNumber})?`,
        ui.ButtonSet.YES_NO
      );
      
      if (response !== ui.Button.YES) return;
      
      // Show move-out form
      const html = HtmlService.createHtmlOutput(`
        <div style="font-family: Arial, sans-serif; padding: 20px;">
          <h3>Process Move-Out: ${tenantName}</h3>
          <p><strong>Room:</strong> ${roomNumber}</p>
          <p><strong>Security Deposit:</strong> ${Utils.formatCurrency(securityDeposit)}</p>
          
          <form id="moveOutForm">
            <input type="hidden" name="rowNumber" value="${row}">
            <input type="hidden" name="tenantName" value="${tenantName}">
            <input type="hidden" name="roomNumber" value="${roomNumber}">
            <input type="hidden" name="securityDeposit" value="${securityDeposit}">
            
            <table>
              <tr>
                <td><label>Move-Out Date:</label></td>
                <td><input type="date" name="moveOutDate" required value="${Utils.formatDate(new Date(), 'yyyy-MM-dd')}"></td>
              </tr>
              <tr>
                <td><label>Forwarding Address:</label></td>
                <td><textarea name="forwardingAddress" rows="3" style="width: 300px;" required></textarea></td>
              </tr>
              <tr>
                <td><label>Room Condition:</label></td>
                <td>
                  <select name="roomCondition" required>
                    <option value="">Select Condition</option>
                    <option value="Excellent">Excellent - No deductions</option>
                    <option value="Good">Good - Minor cleaning needed</option>
                    <option value="Fair">Fair - Some repairs/deep cleaning</option>
                    <option value="Poor">Poor - Significant repairs needed</option>
                  </select>
                </td>
              </tr>
              <tr>
                <td><label>Deductions Amount:</label></td>
                <td><input type="number" name="deductions" step="0.01" min="0" max="${securityDeposit}" value="0"></td>
              </tr>
              <tr>
                <td><label>Deduction Reason:</label></td>
                <td><textarea name="deductionReason" rows="2" style="width: 300px;"></textarea></td>
              </tr>
              <tr>
                <td><label>Keys Returned:</label></td>
                <td>
                  <input type="checkbox" name="keysReturned" value="yes" required>
                  <label>All keys returned</label>
                </td>
              </tr>
              <tr>
                <td><label>Final Notes:</label></td>
                <td><textarea name="finalNotes" rows="3" style="width: 300px;"></textarea></td>
              </tr>
            </table>
            <br>
            <button type="button" onclick="completeMoveOut()">Complete Move-Out</button>
            <button type="button" onclick="google.script.host.close()">Cancel</button>
          </form>
        </div>
        
        <script>
          function completeMoveOut() {
            const form = document.getElementById('moveOutForm');
            const formData = new FormData(form);
            const data = Object.fromEntries(formData);
            
            if (!data.keysReturned) {
              alert('Please confirm that all keys have been returned.');
              return;
            }
            
            google.script.run
              .withSuccessHandler(function(result) {
                alert('Move-out processed successfully!');
                google.script.host.close();
              })
              .withFailureHandler(function(error) {
                alert('Error: ' + error.message);
              })
              .completeMoveOut(data);
          }
        </script>
      `)
        .setWidth(500)
        .setHeight(600);
      
      ui.showModalDialog(html, 'Process Move-Out');
      
    } catch (error) {
      handleSystemError(error, 'processMoveOut');
    }
  },
  
  /**
   * Complete move-out process (called from HTML form)
   */
  completeMoveOut: function(data) {
    try {
      const sheet = SheetManager.getSheet(CONFIG.SHEETS.TENANTS);
      const rowNumber = parseInt(data.rowNumber);
      const moveOutDate = new Date(data.moveOutDate);
      const deductions = parseFloat(data.deductions) || 0;
      const depositRefund = parseFloat(data.securityDeposit) - deductions;
      
      // Update tenant record
      sheet.getRange(rowNumber, this.COL.ROOM_STATUS).setValue(CONFIG.STATUS.ROOM.VACANT);
      sheet.getRange(rowNumber, this.COL.MOVE_OUT_PLANNED).setValue(moveOutDate);
      sheet.getRange(rowNumber, this.COL.PAYMENT_STATUS).setValue('');
      sheet.getRange(rowNumber, this.COL.NOTES).setValue(
        `MOVED OUT: ${Utils.formatDate(moveOutDate)} - ${data.finalNotes || ''}`
      );
      
      // Clear tenant information but keep historical data in notes
      const historicalNote = `Former tenant: ${data.tenantName} (${Utils.formatDate(moveOutDate)})`;
      sheet.getRange(rowNumber, this.COL.TENANT_NAME).setValue('');
      sheet.getRange(rowNumber, this.COL.TENANT_EMAIL).setValue('');
      sheet.getRange(rowNumber, this.COL.TENANT_PHONE).setValue('');
      sheet.getRange(rowNumber, this.COL.LAST_PAYMENT).setValue('');
      sheet.getRange(rowNumber, this.COL.EMERGENCY_CONTACT).setValue('');
      sheet.getRange(rowNumber, this.COL.LEASE_END_DATE).setValue('');
      
      // Log security deposit refund
      if (depositRefund > 0) {
        FinancialManager.logPayment({
          date: moveOutDate,
          type: 'Security Deposit Refund',
          description: `Security deposit refund to ${data.tenantName} - Room ${data.roomNumber}`,
          amount: -depositRefund, // Negative because it's money going out
          category: 'Deposit Refund',
          tenant: data.tenantName,
          reference: `REFUND-${data.roomNumber}-${Utils.formatDate(moveOutDate, 'yyyyMMdd')}`
        });
      }
      
      // Log any deductions
      if (deductions > 0) {
        FinancialManager.logPayment({
          date: moveOutDate,
          type: 'Deposit Deduction',
          description: `Deposit deduction - ${data.tenantName} - ${data.deductionReason}`,
          amount: deductions,
          category: 'Maintenance',
          tenant: data.tenantName,
          reference: `DEDUCTION-${data.roomNumber}-${Utils.formatDate(moveOutDate, 'yyyyMMdd')}`
        });
      }
      
      // Send move-out confirmation email
      EmailManager.sendMoveOutConfirmation(sheet.getRange(rowNumber, this.COL.TENANT_EMAIL).getValue(), {
        tenantName: data.tenantName,
        roomNumber: data.roomNumber,
        moveOutDate: Utils.formatDate(moveOutDate, 'MMMM dd, yyyy'),
        depositRefund: depositRefund,
        deductions: deductions,
        forwardingAddress: data.forwardingAddress
      });
      
      // Create move-out document
      DocumentManager.createMoveOutReport({
        tenantName: data.tenantName,
        roomNumber: data.roomNumber,
        moveOutDate: moveOutDate,
        roomCondition: data.roomCondition,
        deductions: deductions,
        depositRefund: depositRefund,
        deductionReason: data.deductionReason,
        finalNotes: data.finalNotes,
        forwardingAddress: data.forwardingAddress
      });
      
      return { success: true, message: 'Move-out processed successfully' };
      
    } catch (error) {
      Logger.log(`Error in completeMoveOut: ${error.toString()}`);
      throw error;
    }
  },
  
  /**
   * Show tenant dashboard
   */
  showTenantDashboard: function() {
    try {
      const data = SheetManager.getAllData(CONFIG.SHEETS.TENANTS);
      const stats = this.calculateTenantStats(data);
      
      const html = HtmlService.createHtmlOutput(`
        <div style="font-family: Arial, sans-serif; padding: 20px;">
          <h2>🏠 Tenant Dashboard</h2>
          
          <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; margin-bottom: 30px;">
            <div style="background: #e3f2fd; padding: 15px; border-radius: 8px; text-align: center;">
              <h3 style="margin: 0; color: #1976d2;">Total Rooms</h3>
              <p style="font-size: 24px; margin: 5px 0; font-weight: bold;">${stats.totalRooms}</p>
            </div>
            <div style="background: #e8f5e8; padding: 15px; border-radius: 8px; text-align: center;">
              <h3 style="margin: 0; color: #388e3c;">Occupied</h3>
              <p style="font-size: 24px; margin: 5px 0; font-weight: bold;">${stats.occupiedRooms}</p>
            </div>
            <div style="background: #fff3e0; padding: 15px; border-radius: 8px; text-align: center;">
              <h3 style="margin: 0; color: #f57c00;">Vacant</h3>
              <p style="font-size: 24px; margin: 5px 0; font-weight: bold;">${stats.vacantRooms}</p>
            </div>
          </div>
          
          <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; margin-bottom: 30px;">
            <div style="background: #f1f8e9; padding: 15px; border-radius: 8px; text-align: center;">
              <h3 style="margin: 0; color: #689f38;">Payments Paid</h3>
              <p style="font-size: 24px; margin: 5px 0; font-weight: bold;">${stats.paidCount}</p>
            </div>
            <div style="background: #fff8e1; padding: 15px; border-radius: 8px; text-align: center;">
              <h3 style="margin: 0; color: #ffa000;">Payments Due</h3>
              <p style="font-size: 24px; margin: 5px 0; font-weight: bold;">${stats.dueCount}</p>
            </div>
            <div style="background: #ffebee; padding: 15px; border-radius: 8px; text-align: center;">
              <h3 style="margin: 0; color: #d32f2f;">Overdue</h3>
              <p style="font-size: 24px; margin: 5px 0; font-weight: bold;">${stats.overdueCount}</p>
            </div>
          </div>
          
          <h3>📊 Occupancy Rate: ${stats.occupancyRate}%</h3>
          <div style="background: #f5f5f5; height: 20px; border-radius: 10px; overflow: hidden;">
            <div style="background: #4caf50; height: 100%; width: ${stats.occupancyRate}%; transition: width 0.3s;"></div>
          </div>
          
          <h3>💰 Monthly Revenue Potential</h3>
          <p><strong>Current Monthly Income:</strong> ${Utils.formatCurrency(stats.currentRevenue)}</p>
          <p><strong>Potential Monthly Income:</strong> ${Utils.formatCurrency(stats.potentialRevenue)}</p>
          <p><strong>Revenue Loss from Vacancies:</strong> ${Utils.formatCurrency(stats.potentialRevenue - stats.currentRevenue)}</p>
          
          ${stats.overdueList.length > 0 ? `
            <h3>⚠️ Overdue Tenants</h3>
            <div style="background: #ffebee; padding: 15px; border-radius: 8px;">
              ${stats.overdueList.map(tenant => 
                `<p>• <strong>${tenant.name}</strong> (Room ${tenant.room}) - Last payment: ${tenant.lastPayment}</p>`
              ).join('')}
            </div>
          ` : ''}
          
          <div style="margin-top: 30px; text-align: center;">
            <button onclick="google.script.run.checkAllPaymentStatus()" style="margin: 5px; padding: 10px 20px;">Update Payment Status</button>
            <button onclick="google.script.run.sendRentReminders()" style="margin: 5px; padding: 10px 20px;">Send Rent Reminders</button>
            ${stats.overdueCount > 0 ? '<button onclick="google.script.run.sendLatePaymentAlerts()" style="margin: 5px; padding: 10px 20px; background: #f44336; color: white;">Send Overdue Alerts</button>' : ''}
          </div>
        </div>
      `)
        .setWidth(800)
        .setHeight(600);
      
      SpreadsheetApp.getUi().showModalDialog(html, 'Tenant Dashboard');
      
    } catch (error) {
      handleSystemError(error, 'showTenantDashboard');
    }
  },
  
  /**
   * Calculate tenant statistics
   */
  calculateTenantStats: function(data) {
    const stats = {
      totalRooms: 0,
      occupiedRooms: 0,
      vacantRooms: 0,
      maintenanceRooms: 0,
      paidCount: 0,
      dueCount: 0,
      overdueCount: 0,
      currentRevenue: 0,
      potentialRevenue: 0,
      overdueList: []
    };
    
    data.forEach(row => {
      if (!row[this.COL.ROOM_NUMBER - 1]) return; // Skip empty rows
      
      stats.totalRooms++;
      
      const roomStatus = row[this.COL.ROOM_STATUS - 1];
      const paymentStatus = row[this.COL.PAYMENT_STATUS - 1];
      const rent = row[this.COL.NEGOTIATED_PRICE - 1] || row[this.COL.RENTAL_PRICE - 1] || 0;
      
      // Room status counts
      switch (roomStatus) {
        case CONFIG.STATUS.ROOM.OCCUPIED:
          stats.occupiedRooms++;
          stats.currentRevenue += rent;
          break;
        case CONFIG.STATUS.ROOM.VACANT:
          stats.vacantRooms++;
          break;
        case CONFIG.STATUS.ROOM.MAINTENANCE:
          stats.maintenanceRooms++;
          break;
      }
      
      // Potential revenue (all rooms)
      stats.potentialRevenue += rent;
      
      // Payment status counts
      switch (paymentStatus) {
        case CONFIG.STATUS.PAYMENT.PAID:
          stats.paidCount++;
          break;
        case CONFIG.STATUS.PAYMENT.DUE:
          stats.dueCount++;
          break;
        case CONFIG.STATUS.PAYMENT.OVERDUE:
          stats.overdueCount++;
          const lastPayment = row[this.COL.LAST_PAYMENT - 1];
          stats.overdueList.push({
            name: row[this.COL.TENANT_NAME - 1],
            room: row[this.COL.ROOM_NUMBER - 1],
            lastPayment: lastPayment ? Utils.formatDate(lastPayment) : 'Never'
          });
          break;
      }
    });
    
    stats.occupancyRate = stats.totalRooms > 0 ? 
      Math.round((stats.occupiedRooms / stats.totalRooms) * 100) : 0;
    
    return stats;
  },
  
  /**
   * Get vacant rooms
   */
  getVacantRooms: function() {
    const data = SheetManager.getAllData(CONFIG.SHEETS.TENANTS);
    const vacantRooms = [];
    
    data.forEach(row => {
      if (row[this.COL.ROOM_STATUS - 1] === CONFIG.STATUS.ROOM.VACANT) {
        vacantRooms.push({
          number: row[this.COL.ROOM_NUMBER - 1],
          rent: row[this.COL.RENTAL_PRICE - 1]
        });
      }
    });
    
    return vacantRooms;
  },
  
  /**
   * Calculate rent due date
   */
  calculateRentDueDate: function() {
    const today = new Date();
    const dueDate = new Date(today.getFullYear(), today.getMonth(), 5); // 5th of current month
    
    if (today.getDate() > 5) {
      // If past the 5th, return next month's due date
      dueDate.setMonth(dueDate.getMonth() + 1);
    }
    
    return Utils.formatDate(dueDate, 'MMMM dd, yyyy');
  },
  
  /**
   * Get tenant by room number
   */
  getTenantByRoom: function(roomNumber) {
    const tenantRows = SheetManager.findRows(CONFIG.SHEETS.TENANTS, this.COL.ROOM_NUMBER, roomNumber);
    return tenantRows.length > 0 ? tenantRows[0] : null;
  },
  
  /**
   * Get tenant by email
   */
  getTenantByEmail: function(email) {
    const tenantRows = SheetManager.findRows(CONFIG.SHEETS.TENANTS, this.COL.TENANT_EMAIL, email);
    return tenantRows.length > 0 ? tenantRows[0] : null;
  }
};

Logger.log('TenantManager module loaded successfully');
