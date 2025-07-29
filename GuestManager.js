// GuestManager.gs - Guest Room Management System
// Handles all guest room operations and short-term bookings

const GuestManager = {
  
  /**
   * Column indexes for Guest Rooms sheet (1-based)
   */
  ROOM_COL: {
    BOOKING_ID: 1,
    ROOM_NUMBER: 2,
    ROOM_NAME: 3,
    ROOM_TYPE: 4,
    MAX_OCCUPANCY: 5,
    AMENITIES: 6,
    DAILY_RATE: 7,
    WEEKLY_RATE: 8,
    MONTHLY_RATE: 9,
    STATUS: 10,
    LAST_CLEANED: 11,
    MAINTENANCE_NOTES: 12,
    CHECK_IN_DATE: 13,
    CHECK_OUT_DATE: 14,
    NUMBER_OF_NIGHTS: 15,
    NUMBER_OF_GUESTS: 16,
    CURRENT_GUEST: 17,
    PURPOSE_OF_VISIT: 18,
    SPECIAL_REQUESTS: 19,
    SOURCE: 20,
    TOTAL_AMOUNT: 21,
    PAYMENT_STATUS: 22,
    BOOKING_STATUS: 23,
    NOTES: 24
  },
  
  /**
   * Column indexes for Guest Room Bookings sheet (1-based)
   */
  BOOKING_COL: {
    BOOKING_ID: 1,
    TIMESTAMP: 2,
    GUEST_NAME: 3,
    EMAIL: 4,
    PHONE: 5,
    ROOM_NUMBER: 6,
    CHECK_IN_DATE: 7,
    CHECK_OUT_DATE: 8,
    NUMBER_OF_NIGHTS: 9,
    NUMBER_OF_GUESTS: 10,
    PURPOSE_OF_VISIT: 11,
    SPECIAL_REQUESTS: 12,
    TOTAL_AMOUNT: 13,
    AMOUNT_PAID: 14,
    PAYMENT_STATUS: 15,
    BOOKING_STATUS: 16,
    SOURCE: 17,
    NOTES: 18
  },
  
  /**
   * Show today's guest activity (arrivals and departures)
   */
  showTodayGuestActivity: function() {
    try {
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      
      const bookingData = SheetManager.getAllData(CONFIG.SHEETS.GUEST_BOOKINGS);
      const arrivals = [];
      const departures = [];
      
      bookingData.forEach(row => {
        const checkIn = new Date(row[this.BOOKING_COL.CHECK_IN_DATE - 1]);
        const checkOut = new Date(row[this.BOOKING_COL.CHECK_OUT_DATE - 1]);
        checkIn.setHours(0, 0, 0, 0);
        checkOut.setHours(0, 0, 0, 0);
        
        const bookingStatus = row[this.BOOKING_COL.BOOKING_STATUS - 1];
        const guestName = row[this.BOOKING_COL.GUEST_NAME - 1];
        const roomNumber = row[this.BOOKING_COL.ROOM_NUMBER - 1];
        const numberOfGuests = row[this.BOOKING_COL.NUMBER_OF_GUESTS - 1];
        
        // Today's arrivals
        if (checkIn.getTime() === today.getTime() && bookingStatus === CONFIG.STATUS.BOOKING.CONFIRMED) {
          arrivals.push({
            guest: guestName,
            room: roomNumber,
            guests: numberOfGuests,
            bookingId: row[this.BOOKING_COL.BOOKING_ID - 1]
          });
        }
        
        // Today's departures
        if (checkOut.getTime() === today.getTime() && bookingStatus === CONFIG.STATUS.BOOKING.CHECKED_IN) {
          const totalAmount = row[this.BOOKING_COL.TOTAL_AMOUNT - 1] || 0;
          const amountPaid = row[this.BOOKING_COL.AMOUNT_PAID - 1] || 0;
          const balance = totalAmount - amountPaid;
          
          departures.push({
            guest: guestName,
            room: roomNumber,
            balance: balance,
            bookingId: row[this.BOOKING_COL.BOOKING_ID - 1]
          });
        }
      });
      
      const html = HtmlService.createHtmlOutput(`
        <div style="font-family: Arial, sans-serif; padding: 20px;">
          <h2>📅 Today's Guest Activity - ${Utils.formatDate(today, 'MMMM dd, yyyy')}</h2>
          
          <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 30px;">
            <div>
              <h3 style="color: #4caf50;">✈️ Arrivals (${arrivals.length})</h3>
              ${arrivals.length > 0 ? `
                <div style="background: #e8f5e8; padding: 15px; border-radius: 8px;">
                  ${arrivals.map(arrival => `
                    <div style="margin-bottom: 10px; padding: 10px; background: white; border-radius: 5px;">
                      <strong>${arrival.guest}</strong><br>
                      Room ${arrival.room} • ${arrival.guests} guest${arrival.guests > 1 ? 's' : ''}<br>
                      <small>Booking: ${arrival.bookingId}</small>
                    </div>
                  `).join('')}
                  <button onclick="google.script.run.sendCheckInReminders()" 
                          style="margin-top: 10px; padding: 8px 16px; background: #4caf50; color: white; border: none; border-radius: 4px;">
                    Send Check-in Reminders
                  </button>
                </div>
              ` : `
                <div style="background: #f5f5f5; padding: 15px; border-radius: 8px; text-align: center; color: #666;">
                  No arrivals scheduled for today
                </div>
              `}
            </div>
            
            <div>
              <h3 style="color: #ff9800;">🧳 Departures (${departures.length})</h3>
              ${departures.length > 0 ? `
                <div style="background: #fff3e0; padding: 15px; border-radius: 8px;">
                  ${departures.map(departure => `
                    <div style="margin-bottom: 10px; padding: 10px; background: white; border-radius: 5px;">
                      <strong>${departure.guest}</strong><br>
                      Room ${departure.room}<br>
                      ${departure.balance > 0 ? 
                        `<span style="color: #f44336;">Balance Due: ${Utils.formatCurrency(departure.balance)}</span>` : 
                        '<span style="color: #4caf50;">Paid in Full</span>'
                      }<br>
                      <small>Booking: ${departure.bookingId}</small>
                    </div>
                  `).join('')}
                  <button onclick="google.script.run.processAllCheckOuts()" 
                          style="margin-top: 10px; padding: 8px 16px; background: #ff9800; color: white; border: none; border-radius: 4px;">
                    Process Check-outs
                  </button>
                </div>
              ` : `
                <div style="background: #f5f5f5; padding: 15px; border-radius: 8px; text-align: center; color: #666;">
                  No departures scheduled for today
                </div>
              `}
            </div>
          </div>
          
          <div style="margin-top: 30px; text-align: center;">
            <button onclick="google.script.run.checkGuestRoomAvailability()" style="margin: 5px; padding: 10px 20px;">Check Availability</button>
            <button onclick="google.script.run.showGuestRoomAnalytics()" style="margin: 5px; padding: 10px 20px;">View Analytics</button>
            <button onclick="google.script.run.showOccupancyCalendar()" style="margin: 5px; padding: 10px 20px;">Occupancy Calendar</button>
          </div>
        </div>
      `)
        .setWidth(800)
        .setHeight(600);
      
      SpreadsheetApp.getUi().showModalDialog(html, 'Today\'s Guest Activity');
      
    } catch (error) {
      handleSystemError(error, 'showTodayGuestActivity');
    }
  },
  
  /**
   * Check guest room availability for a date range
   */
  checkGuestRoomAvailability: function() {
    try {
      const ui = SpreadsheetApp.getUi();
      
      const checkInResponse = ui.prompt(
        'Check Guest Room Availability',
        'Enter check-in date (MM/DD/YYYY):',
        ui.ButtonSet.OK_CANCEL
      );
      
      if (checkInResponse.getSelectedButton() !== ui.Button.OK) return;
      
      const checkOutResponse = ui.prompt(
        'Check Guest Room Availability',
        'Enter check-out date (MM/DD/YYYY):',
        ui.ButtonSet.OK_CANCEL
      );
      
      if (checkOutResponse.getSelectedButton() !== ui.Button.OK) return;
      
      const checkIn = new Date(checkInResponse.getResponseText());
      const checkOut = new Date(checkOutResponse.getResponseText());
      
      if (isNaN(checkIn.getTime()) || isNaN(checkOut.getTime()) || checkOut <= checkIn) {
        ui.alert('Invalid dates entered. Please ensure check-out is after check-in.');
        return;
      }
      
      const availableRooms = this.getAvailableGuestRooms(checkIn, checkOut);
      const numberOfNights = Utils.daysBetween(checkIn, checkOut);
      
      if (availableRooms.length === 0) {
        ui.alert('No Rooms Available', 'No guest rooms are available for the selected dates.', ui.ButtonSet.OK);
      } else {
        const roomList = availableRooms.map(room => {
          const totalCost = this.calculateDynamicPrice(room, checkIn, checkOut, numberOfNights);
          return `${room.name} (${room.number}) - ${Utils.formatCurrency(room.dailyRate)}/night - Total: ${Utils.formatCurrency(totalCost)}`;
        }).join('\n');
        
        ui.alert(
          'Available Rooms',
          `Available rooms for ${Utils.formatDate(checkIn, 'MM/dd/yyyy')} to ${Utils.formatDate(checkOut, 'MM/dd/yyyy')} (${numberOfNights} nights):\n\n${roomList}`,
          ui.ButtonSet.OK
        );
      }
      
    } catch (error) {
      handleSystemError(error, 'checkGuestRoomAvailability');
    }
  },

  /**
   * Show panel to process guest check-in from form responses
   */
  showProcessCheckInPanel: function() {
    try {
      const data = SheetManager.getAllData(CONFIG.SHEETS.GUEST_BOOKINGS);
      const headerMap = SheetManager.getHeaderMap(CONFIG.SHEETS.GUEST_BOOKINGS);
      const gCol = {
        GUEST_NAME: headerMap['Guest Name'] || this.BOOKING_COL.GUEST_NAME,
        EMAIL: headerMap['Email'] || this.BOOKING_COL.EMAIL,
        PHONE: headerMap['Phone'] || this.BOOKING_COL.PHONE,
        ROOM_NUMBER: headerMap['Room Number'] || this.BOOKING_COL.ROOM_NUMBER,
        CHECK_IN_DATE: headerMap['Check-In Date'] || this.BOOKING_COL.CHECK_IN_DATE,
        CHECK_OUT_DATE: headerMap['Check-Out Date'] || this.BOOKING_COL.CHECK_OUT_DATE,
        NUMBER_OF_NIGHTS: headerMap['Number of Nights'] || this.BOOKING_COL.NUMBER_OF_NIGHTS
      };
      const options = data.map((row, i) => `<option value="${i}">${row[gCol.GUEST_NAME - 1]}</option>`).join('');

      const html = HtmlService.createHtmlOutput(`
        <div style="font-family: Arial, sans-serif; padding:20px;">
          <h3>Process Check-In</h3>
          <label>Guest:</label>
          <select id="guestSelect" onchange="fillFields()">
            <option value="">Select...</option>
            ${options}
          </select>
          <div style="margin-top:10px;">
            <label>Name:</label><input type="text" id="guestName" style="width:100%" readonly><br>
            <label>Email:</label><input type="text" id="email" style="width:100%" readonly><br>
            <label>Phone:</label><input type="text" id="phone" style="width:100%" readonly><br>
            <label>Room:</label><input type="text" id="room" style="width:100%" readonly><br>
            <label>Check-In:</label><input type="text" id="checkin" style="width:100%" readonly><br>
            <label>Check-Out:</label><input type="text" id="checkout" style="width:100%" readonly><br>
          </div>
          <button onclick="processCheckIn()" style="margin-top:10px;">Process Check-In</button>
          <script>
            const data = ${JSON.stringify(data)};
            const gCol = ${JSON.stringify(gCol)};
            function fillFields(){
              const idx = document.getElementById('guestSelect').value;
              if(idx===''){return;}
              const d = data[idx];
              document.getElementById('guestName').value = d[gCol.GUEST_NAME - 1];
              document.getElementById('email').value = d[gCol.EMAIL - 1];
              document.getElementById('phone').value = d[gCol.PHONE - 1];
              document.getElementById('room').value = d[gCol.ROOM_NUMBER - 1];
              document.getElementById('checkin').value = d[gCol.CHECK_IN_DATE - 1];
              let co = d[gCol.CHECK_OUT_DATE - 1];
              const nights = parseInt(d[gCol.NUMBER_OF_NIGHTS - 1]) || 0;
              if(!co && d[gCol.CHECK_IN_DATE - 1]){
                const start = new Date(d[gCol.CHECK_IN_DATE - 1]);
                if(!isNaN(start.getTime()) && nights){
                  start.setDate(start.getDate() + nights);
                  co = start.toISOString().slice(0,10);
                }
              }
              document.getElementById('checkout').value = co || '';
            }
            function processCheckIn(){
              const idx = document.getElementById('guestSelect').value;
              if(idx==='') return;
              google.script.run.processCheckInFromForm(Number(idx)+2);
              google.script.host.close();
            }
          </script>
        </div>
      `).setWidth(400).setHeight(500);

      SpreadsheetApp.getUi().showModalDialog(html, 'Process Check-In');

    } catch (error) {
      handleSystemError(error, 'showProcessCheckInPanel');
    }
  },
  
  /**
   * Get available guest rooms for a date range
   */
  getAvailableGuestRooms: function(checkIn, checkOut) {
    const roomsData = SheetManager.getAllData(CONFIG.SHEETS.GUEST_ROOMS);
    const bookingsData = SheetManager.getAllData(CONFIG.SHEETS.GUEST_BOOKINGS);
    const availableRooms = [];
    
    roomsData.forEach(room => {
      if (!room[this.ROOM_COL.ROOM_NUMBER - 1]) return;
      
      let isAvailable = true;
      const roomStatus = room[this.ROOM_COL.STATUS - 1];
      
      // Skip rooms under maintenance
      if (roomStatus === 'Maintenance') {
        isAvailable = false;
      } else {
        // Check for overlapping bookings
        bookingsData.forEach(booking => {
          if (booking[this.BOOKING_COL.ROOM_NUMBER - 1] === room[this.ROOM_COL.ROOM_NUMBER - 1]) {
            const bookingStatus = booking[this.BOOKING_COL.BOOKING_STATUS - 1];
            
            if (bookingStatus === CONFIG.STATUS.BOOKING.CONFIRMED || 
                bookingStatus === CONFIG.STATUS.BOOKING.CHECKED_IN) {
              
              const bookingCheckIn = new Date(booking[this.BOOKING_COL.CHECK_IN_DATE - 1]);
              const bookingCheckOut = new Date(booking[this.BOOKING_COL.CHECK_OUT_DATE - 1]);
              
              // Check for date overlap
              if (!(checkOut <= bookingCheckIn || checkIn >= bookingCheckOut)) {
                isAvailable = false;
              }
            }
          }
        });
      }
      
      if (isAvailable) {
        availableRooms.push({
          number: room[this.ROOM_COL.ROOM_NUMBER - 1],
          name: room[this.ROOM_COL.ROOM_NAME - 1],
          dailyRate: room[this.ROOM_COL.DAILY_RATE - 1],
          weeklyRate: room[this.ROOM_COL.WEEKLY_RATE - 1],
          monthlyRate: room[this.ROOM_COL.MONTHLY_RATE - 1],
          maxOccupancy: room[this.ROOM_COL.MAX_OCCUPANCY - 1],
          amenities: room[this.ROOM_COL.AMENITIES - 1]
        });
      }
    });
    
    return availableRooms;
  },
  
  /**
   * Calculate dynamic pricing based on various factors
   */
  calculateDynamicPrice: function(room, checkIn, checkOut, numberOfNights) {
    let baseRate = room.dailyRate;
    let totalCost = 0;
    
    // Apply weekly/monthly discounts if applicable
    if (numberOfNights >= 28) {
      totalCost = room.monthlyRate || (baseRate * numberOfNights * 0.8); // 20% monthly discount
    } else if (numberOfNights >= 7) {
      totalCost = room.weeklyRate || (baseRate * numberOfNights * 0.9); // 10% weekly discount
    } else {
      // Day-by-day pricing with weekend premiums
      const currentDate = new Date(checkIn);
      
      for (let i = 0; i < numberOfNights; i++) {
        let dayRate = baseRate;
        
        // Weekend premium (Friday and Saturday nights)
        const dayOfWeek = currentDate.getDay();
        if (dayOfWeek === 5 || dayOfWeek === 6) { // Friday or Saturday
          dayRate = baseRate * 1.25; // 25% weekend premium
        }
        
        // Seasonal adjustments
        const month = currentDate.getMonth();
        if (month >= 5 && month <= 7) { // June-August (summer)
          dayRate = dayRate * 1.15; // 15% summer premium
        } else if (month === 0 || month === 1) { // January-February (winter)
          dayRate = dayRate * 0.9; // 10% winter discount
        }
        
        totalCost += dayRate;
        currentDate.setDate(currentDate.getDate() + 1);
      }
    }
    
    return Math.round(totalCost * 100) / 100; // Round to 2 decimal places
  },
  
  /**
   * Process guest check-in
   */
  processGuestCheckIn: function() {
    try {
      const ui = SpreadsheetApp.getUi();
      const sheet = SpreadsheetApp.getActiveSheet();
      
      if (sheet.getName() !== CONFIG.SHEETS.GUEST_BOOKINGS) {
        ui.alert('Please select a booking in the Guest Room Bookings sheet.');
        return;
      }
      
      const row = sheet.getActiveRange().getRow();
      if (row <= 1) {
        ui.alert('Please select a booking row.');
        return;
      }
      
      const bookingStatus = sheet.getRange(row, this.BOOKING_COL.BOOKING_STATUS).getValue();
      if (bookingStatus !== CONFIG.STATUS.BOOKING.CONFIRMED) {
        ui.alert('Only confirmed bookings can be checked in.');
        return;
      }
      
      const guestName = sheet.getRange(row, this.BOOKING_COL.GUEST_NAME).getValue();
      const roomNumber = sheet.getRange(row, this.BOOKING_COL.ROOM_NUMBER).getValue();
      const checkOutDate = sheet.getRange(row, this.BOOKING_COL.CHECK_OUT_DATE).getValue();
      const totalAmount = sheet.getRange(row, this.BOOKING_COL.TOTAL_AMOUNT).getValue() || 0;
      const amountPaid = sheet.getRange(row, this.BOOKING_COL.AMOUNT_PAID).getValue() || 0;
      
      // Check if payment is required
      if (amountPaid < totalAmount) {
        const response = ui.alert(
          'Outstanding Balance',
          `Guest has an outstanding balance of ${Utils.formatCurrency(totalAmount - amountPaid)}.\n\nProceed with check-in?`,
          ui.ButtonSet.YES_NO
        );
        
        if (response !== ui.Button.YES) return;
      }
      
      // Update booking status
      sheet.getRange(row, this.BOOKING_COL.BOOKING_STATUS).setValue(CONFIG.STATUS.BOOKING.CHECKED_IN);
      
      // Update guest room status
      this.updateGuestRoomStatus(roomNumber, 'Occupied', guestName, new Date(), checkOutDate);
      
      // Send welcome email with check-in information
      const email = sheet.getRange(row, this.BOOKING_COL.EMAIL).getValue();
      if (email) {
        EmailManager.sendGuestWelcome(email, {
          guestName: guestName,
          roomNumber: roomNumber,
          checkOutDate: Utils.formatDate(checkOutDate, 'MMMM dd, yyyy'),
          propertyName: CONFIG.SYSTEM.PROPERTY_NAME
        });
      }
      
      ui.alert('Check-In Complete', `${guestName} has been checked into room ${roomNumber}.`, ui.ButtonSet.OK);
      
    } catch (error) {
      handleSystemError(error, 'processGuestCheckIn');
    }
  },
  
  /**
   * Process guest check-out
   */
  processGuestCheckOut: function() {
    try {
      const ui = SpreadsheetApp.getUi();
      const sheet = SpreadsheetApp.getActiveSheet();
      
      if (sheet.getName() !== CONFIG.SHEETS.GUEST_BOOKINGS) {
        ui.alert('Please select a booking in the Guest Room Bookings sheet.');
        return;
      }
      
      const row = sheet.getActiveRange().getRow();
      if (row <= 1) {
        ui.alert('Please select a booking row.');
        return;
      }
      
      const bookingStatus = sheet.getRange(row, this.BOOKING_COL.BOOKING_STATUS).getValue();
      if (bookingStatus !== CONFIG.STATUS.BOOKING.CHECKED_IN) {
        ui.alert('Only checked-in bookings can be checked out.');
        return;
      }
      
      const guestName = sheet.getRange(row, this.BOOKING_COL.GUEST_NAME).getValue();
      const roomNumber = sheet.getRange(row, this.BOOKING_COL.ROOM_NUMBER).getValue();
      const totalAmount = sheet.getRange(row, this.BOOKING_COL.TOTAL_AMOUNT).getValue() || 0;
      const amountPaid = sheet.getRange(row, this.BOOKING_COL.AMOUNT_PAID).getValue() || 0;
      const balance = totalAmount - amountPaid;
      
      // Check for outstanding balance
      if (balance > 0) {
        const response = ui.prompt(
          'Outstanding Balance',
          `There is an outstanding balance of ${Utils.formatCurrency(balance)}.\nEnter additional payment amount (or 0 to proceed):`,
          ui.ButtonSet.OK_CANCEL
        );
        
        if (response.getSelectedButton() !== ui.Button.OK) return;
        
        const additionalPayment = parseFloat(response.getResponseText()) || 0;
        if (additionalPayment > 0) {
          sheet.getRange(row, this.BOOKING_COL.AMOUNT_PAID).setValue(amountPaid + additionalPayment);
        }
      }
      
      // Update booking status
      sheet.getRange(row, this.BOOKING_COL.BOOKING_STATUS).setValue(CONFIG.STATUS.BOOKING.CHECKED_OUT);
      
      // Update guest room status
      this.updateGuestRoomStatus(roomNumber, 'Available', '', '', '');
      
      // Log revenue in budget
      const finalAmountPaid = sheet.getRange(row, this.BOOKING_COL.AMOUNT_PAID).getValue() || 0;
      FinancialManager.logPayment({
        date: new Date(),
        type: 'Guest Room Income',
        description: `Guest room rental - ${guestName} (Room ${roomNumber})`,
        amount: finalAmountPaid,
        category: 'Guest Room',
        tenant: guestName,
        reference: sheet.getRange(row, this.BOOKING_COL.BOOKING_ID).getValue()
      });
      
      // Send checkout confirmation
      const email = sheet.getRange(row, this.BOOKING_COL.EMAIL).getValue();
      if (email) {
        EmailManager.sendGuestCheckoutConfirmation(email, {
          guestName: guestName,
          roomNumber: roomNumber,
          totalAmount: totalAmount,
          amountPaid: finalAmountPaid,
          propertyName: CONFIG.SYSTEM.PROPERTY_NAME
        });
      }
      
      ui.alert('Check-Out Complete', `Check-out completed for ${guestName} from room ${roomNumber}.`, ui.ButtonSet.OK);
      
    } catch (error) {
      handleSystemError(error, 'processGuestCheckOut');
    }
  },
  
  /**
   * Update guest room status
   */
  updateGuestRoomStatus: function(roomNumber, status, guestName, checkInDate, checkOutDate) {
    const sheet = SheetManager.getSheet(CONFIG.SHEETS.GUEST_ROOMS);
    const roomRows = SheetManager.findRows(CONFIG.SHEETS.GUEST_ROOMS, this.ROOM_COL.ROOM_NUMBER, roomNumber);
    
    if (roomRows.length > 0) {
      const rowNumber = roomRows[0].rowNumber;

      sheet.getRange(rowNumber, this.ROOM_COL.STATUS).setValue(status);
      sheet.getRange(rowNumber, this.ROOM_COL.CURRENT_GUEST).setValue(guestName || '');
      sheet.getRange(rowNumber, this.ROOM_COL.CHECK_IN_DATE).setValue(checkInDate || '');
      sheet.getRange(rowNumber, this.ROOM_COL.CHECK_OUT_DATE).setValue(checkOutDate || '');

      // Update last cleaned date when room becomes available
      if (status === 'Available') {
        sheet.getRange(rowNumber, this.ROOM_COL.LAST_CLEANED).setValue(new Date());
      }
    }
  },

  /**
   * Process check-in from form row
   */
  processCheckInFromForm: function(rowNumber) {
    try {
      const formSheet = SheetManager.getSheet(CONFIG.SHEETS.GUEST_BOOKINGS);
      const values = formSheet.getRange(rowNumber, 1, 1, formSheet.getLastColumn()).getValues()[0];

      const headerMap = SheetManager.getHeaderMap(CONFIG.SHEETS.GUEST_BOOKINGS);
      const gCol = {
        GUEST_NAME: headerMap['Guest Name'] || this.BOOKING_COL.GUEST_NAME,
        ROOM_NUMBER: headerMap['Room Number'] || this.BOOKING_COL.ROOM_NUMBER,
        CHECK_IN_DATE: headerMap['Check-In Date'] || this.BOOKING_COL.CHECK_IN_DATE,
        CHECK_OUT_DATE: headerMap['Check-Out Date'] || this.BOOKING_COL.CHECK_OUT_DATE,
        NUMBER_OF_NIGHTS: headerMap['Number of Nights'] || this.BOOKING_COL.NUMBER_OF_NIGHTS,
        NUMBER_OF_GUESTS: headerMap['Number of Guests'] || this.BOOKING_COL.NUMBER_OF_GUESTS,
        PURPOSE_OF_VISIT: headerMap['Purpose of Visit'] || this.BOOKING_COL.PURPOSE_OF_VISIT,
        SPECIAL_REQUESTS: headerMap['Special Requests'] || this.BOOKING_COL.SPECIAL_REQUESTS
      };

      const guestName = values[gCol.GUEST_NAME - 1];
      const roomNumber = values[gCol.ROOM_NUMBER - 1];
      const checkInDate = values[gCol.CHECK_IN_DATE - 1];
      let checkOutDate = values[gCol.CHECK_OUT_DATE - 1];
      const nightsVal = parseInt(values[gCol.NUMBER_OF_NIGHTS - 1], 10) || 0;
      if (!checkOutDate && checkInDate && nightsVal) {
        const tmp = new Date(checkInDate);
        if (!isNaN(tmp.getTime())) {
          tmp.setDate(tmp.getDate() + nightsVal);
          checkOutDate = tmp;
        }
      }
      const nights = nightsVal || Utils.daysBetween(new Date(checkInDate), new Date(checkOutDate));
      const guests = values[gCol.NUMBER_OF_GUESTS - 1];
      const purpose = values[gCol.PURPOSE_OF_VISIT - 1];
      const requests = values[gCol.SPECIAL_REQUESTS - 1];

      const roomRows = SheetManager.findRows(CONFIG.SHEETS.GUEST_ROOMS, this.ROOM_COL.ROOM_NUMBER, roomNumber);
      if (roomRows.length === 0) return;

      const row = roomRows[0].rowNumber;
      const roomSheet = SheetManager.getSheet(CONFIG.SHEETS.GUEST_ROOMS);

      const dailyRate = roomSheet.getRange(row, this.ROOM_COL.DAILY_RATE).getValue() || 0;
      const weeklyRate = roomSheet.getRange(row, this.ROOM_COL.WEEKLY_RATE).getValue() || 0;
      const monthlyRate = roomSheet.getRange(row, this.ROOM_COL.MONTHLY_RATE).getValue() || 0;

      let total = dailyRate * nights;
      if (nights >= 28 && monthlyRate) total = monthlyRate;
      else if (nights >= 7 && weeklyRate) total = weeklyRate;

      const bookingId = Utils.generateId('BK');

      roomSheet.getRange(row, this.ROOM_COL.BOOKING_ID).setValue(bookingId);
      roomSheet.getRange(row, this.ROOM_COL.CURRENT_GUEST).setValue(guestName);
      roomSheet.getRange(row, this.ROOM_COL.STATUS).setValue('Occupied');
      roomSheet.getRange(row, this.ROOM_COL.CHECK_IN_DATE).setValue(checkInDate);
      roomSheet.getRange(row, this.ROOM_COL.CHECK_OUT_DATE).setValue(checkOutDate);
      roomSheet.getRange(row, this.ROOM_COL.NUMBER_OF_NIGHTS).setValue(nights);
      roomSheet.getRange(row, this.ROOM_COL.NUMBER_OF_GUESTS).setValue(guests);
      roomSheet.getRange(row, this.ROOM_COL.PURPOSE_OF_VISIT).setValue(purpose);
      roomSheet.getRange(row, this.ROOM_COL.SPECIAL_REQUESTS).setValue(requests);
      roomSheet.getRange(row, this.ROOM_COL.TOTAL_AMOUNT).setValue(total);
      roomSheet.getRange(row, this.ROOM_COL.PAYMENT_STATUS).setValue(CONFIG.STATUS.PAYMENT.DUE);
      roomSheet.getRange(row, this.ROOM_COL.BOOKING_STATUS).setValue(CONFIG.STATUS.BOOKING.CHECKED_IN);
      roomSheet.getRange(row, this.ROOM_COL.SOURCE).setValue('Form');

    } catch (error) {
      handleSystemError(error, 'processCheckInFromForm');
    }
  },

  /**
   * Show panel to process guest check-out from occupied rooms
   */
  showProcessCheckOutPanel: function() {
    try {
      const sheet = SheetManager.getSheet(CONFIG.SHEETS.GUEST_ROOMS);
      const data = sheet.getDataRange().getValues();
      const occupied = [];

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row[this.ROOM_COL.STATUS - 1] === 'Occupied') {
          occupied.push({
            row: i + 1,
            guest: row[this.ROOM_COL.CURRENT_GUEST - 1],
            room: row[this.ROOM_COL.ROOM_NUMBER - 1],
            checkIn: row[this.ROOM_COL.CHECK_IN_DATE - 1],
            checkOut: row[this.ROOM_COL.CHECK_OUT_DATE - 1]
          });
        }
      }

      if (occupied.length === 0) {
        SpreadsheetApp.getUi().alert('No guests currently checked in.');
        return;
      }

      const options = occupied.map(o => `<option value="${o.row}">${o.guest} - Room ${o.room}</option>`).join('');

      const html = HtmlService.createHtmlOutput(`
        <div style="font-family: Arial, sans-serif; padding:20px;">
          <h3>Process Check-Out</h3>
          <label>Guest:</label>
          <select id="guestSelect" onchange="fillFields()">
            <option value="">Select...</option>
            ${options}
          </select>
          <div style="margin-top:10px;">
            <label>Room:</label><input type="text" id="room" style="width:100%" readonly><br>
            <label>Check-In:</label><input type="text" id="checkin" style="width:100%" readonly><br>
            <label>Check-Out:</label><input type="text" id="checkout" style="width:100%" readonly><br>
          </div>
          <button onclick="processCO()" style="margin-top:10px;">Process Check-Out</button>
          <script>
            const data = ${JSON.stringify(occupied)};
            function fillFields(){
              const val = document.getElementById('guestSelect').value;
              const rec = data.find(r => r.row == val);
              if(!rec){return;}
              document.getElementById('room').value = rec.room;
              document.getElementById('checkin').value = rec.checkIn;
              document.getElementById('checkout').value = rec.checkOut;
            }
            function processCO(){
              const val = document.getElementById('guestSelect').value;
              if(val=='') return;
              google.script.run.processRoomCheckOut(Number(val));
              google.script.host.close();
            }
          </script>
        </div>
      `).setWidth(400).setHeight(400);

      SpreadsheetApp.getUi().showModalDialog(html, 'Process Check-Out');

    } catch (error) {
      handleSystemError(error, 'showProcessCheckOutPanel');
    }
  },

  /**
   * Process check-out for a room row
   */
  processRoomCheckOut: function(rowNumber) {
    try {
      const sheet = SheetManager.getSheet(CONFIG.SHEETS.GUEST_ROOMS);
      const guestName = sheet.getRange(rowNumber, this.ROOM_COL.CURRENT_GUEST).getValue();
      const roomNum = sheet.getRange(rowNumber, this.ROOM_COL.ROOM_NUMBER).getValue();

      sheet.getRange(rowNumber, this.ROOM_COL.STATUS).setValue('Available');
      sheet.getRange(rowNumber, this.ROOM_COL.CURRENT_GUEST).clearContent();
      sheet.getRange(rowNumber, this.ROOM_COL.CHECK_IN_DATE).clearContent();
      sheet.getRange(rowNumber, this.ROOM_COL.CHECK_OUT_DATE).clearContent();
      sheet.getRange(rowNumber, this.ROOM_COL.BOOKING_STATUS).setValue(CONFIG.STATUS.BOOKING.CHECKED_OUT);
      sheet.getRange(rowNumber, this.ROOM_COL.LAST_CLEANED).setValue(new Date());

      SpreadsheetApp.getUi().alert('Check-Out Complete', `Check-out completed for ${guestName} from room ${roomNum}.`, SpreadsheetApp.getUi().ButtonSet.OK);

    } catch (error) {
      handleSystemError(error, 'processRoomCheckOut');
    }
  },
  
  /**
   * Show guest room analytics
   */
  showGuestRoomAnalytics: function() {
    try {
      const bookingData = SheetManager.getAllData(CONFIG.SHEETS.GUEST_BOOKINGS);
      const roomData = SheetManager.getAllData(CONFIG.SHEETS.GUEST_ROOMS);
      
      const analytics = this.calculateGuestAnalytics(bookingData, roomData);
      
      const html = HtmlService.createHtmlOutput(`
        <div style="font-family: Arial, sans-serif; padding: 20px;">
          <h2>📊 Guest Room Analytics</h2>
          
          <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; margin-bottom: 30px;">
            <div style="background: #e3f2fd; padding: 15px; border-radius: 8px; text-align: center;">
              <h3 style="margin: 0; color: #1976d2;">Total Bookings</h3>
              <p style="font-size: 24px; margin: 5px 0; font-weight: bold;">${analytics.totalBookings}</p>
              <small>This month</small>
            </div>
            <div style="background: #e8f5e8; padding: 15px; border-radius: 8px; text-align: center;">
              <h3 style="margin: 0; color: #388e3c;">Occupancy Rate</h3>
              <p style="font-size: 24px; margin: 5px 0; font-weight: bold;">${analytics.occupancyRate}%</p>
              <small>This month</small>
            </div>
            <div style="background: #fff3e0; padding: 15px; border-radius: 8px; text-align: center;">
              <h3 style="margin: 0; color: #f57c00;">Avg Daily Rate</h3>
              <p style="font-size: 24px; margin: 5px 0; font-weight: bold;">${Utils.formatCurrency(analytics.avgDailyRate)}</p>
              <small>This month</small>
            </div>
          </div>
          
          <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 30px;">
            <div>
              <h3>💰 Revenue Analysis</h3>
              <div style="background: #f5f5f5; padding: 15px; border-radius: 8px;">
                <p><strong>This Month:</strong> ${Utils.formatCurrency(analytics.monthlyRevenue)}</p>
                <p><strong>YTD Revenue:</strong> ${Utils.formatCurrency(analytics.ytdRevenue)}</p>
                <p><strong>Average Booking Value:</strong> ${Utils.formatCurrency(analytics.avgBookingValue)}</p>
                <p><strong>Revenue per Available Room:</strong> ${Utils.formatCurrency(analytics.revPAR)}</p>
              </div>
            </div>
            
            <div>
              <h3>📈 Booking Patterns</h3>
              <div style="background: #f5f5f5; padding: 15px; border-radius: 8px;">
                <p><strong>Average Stay:</strong> ${analytics.avgStayLength} nights</p>
                <p><strong>Weekend vs Weekday:</strong></p>
                <ul style="margin: 5px 0;">
                  <li>Weekend Bookings: ${analytics.weekendBookings}%</li>
                  <li>Weekday Bookings: ${analytics.weekdayBookings}%</li>
                </ul>
                <p><strong>Top Purpose:</strong> ${analytics.topPurpose}</p>
              </div>
            </div>
          </div>
          
          <h3>🏠 Room Performance</h3>
          <div style="background: #f5f5f5; padding: 15px; border-radius: 8px;">
            ${analytics.roomPerformance.map(room => `
              <div style="margin-bottom: 10px; padding: 10px; background: white; border-radius: 5px;">
                <strong>${room.name}</strong> (${room.number})<br>
                Bookings: ${room.bookings} | Revenue: ${Utils.formatCurrency(room.revenue)} | Occupancy: ${room.occupancy}%
              </div>
            `).join('')}
          </div>
          
          <div style="margin-top: 30px; text-align: center;">
            <button onclick="google.script.run.analyzeGuestRoomPricing()" style="margin: 5px; padding: 10px 20px;">Pricing Analysis</button>
            <button onclick="google.script.run.generateGuestRoomReport()" style="margin: 5px; padding: 10px 20px;">Generate Report</button>
          </div>
        </div>
      `)
        .setWidth(800)
        .setHeight(700);
      
      SpreadsheetApp.getUi().showModalDialog(html, 'Guest Room Analytics');
      
    } catch (error) {
      handleSystemError(error, 'showGuestRoomAnalytics');
    }
  },
  
  /**
   * Calculate guest room analytics
   */
  calculateGuestAnalytics: function(bookingData, roomData) {
    const now = new Date();
    const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
    const yearStart = new Date(now.getFullYear(), 0, 1);
    
    const analytics = {
      totalBookings: 0,
      monthlyRevenue: 0,
      ytdRevenue: 0,
      totalNights: 0,
      totalRevenue: 0,
      weekendBookings: 0,
      weekdayBookings: 0,
      purposeStats: {},
      roomStats: {}
    };
    
    // Initialize room stats
    roomData.forEach(room => {
      if (room[this.ROOM_COL.ROOM_NUMBER - 1]) {
        analytics.roomStats[room[this.ROOM_COL.ROOM_NUMBER - 1]] = {
          name: room[this.ROOM_COL.ROOM_NAME - 1],
          number: room[this.ROOM_COL.ROOM_NUMBER - 1],
          bookings: 0,
          revenue: 0,
          nights: 0
        };
      }
    });
    
    // Process booking data
    bookingData.forEach(booking => {
      const checkIn = new Date(booking[this.BOOKING_COL.CHECK_IN_DATE - 1]);
      const amount = booking[this.BOOKING_COL.AMOUNT_PAID - 1] || 0;
      const nights = booking[this.BOOKING_COL.NUMBER_OF_NIGHTS - 1] || 0;
      const status = booking[this.BOOKING_COL.BOOKING_STATUS - 1];
      const purpose = booking[this.BOOKING_COL.PURPOSE_OF_VISIT - 1] || 'Other';
      const roomNumber = booking[this.BOOKING_COL.ROOM_NUMBER - 1];
      
      if (status === CONFIG.STATUS.BOOKING.CHECKED_OUT || status === CONFIG.STATUS.BOOKING.CHECKED_IN) {
        // This month bookings
        if (checkIn >= monthStart) {
          analytics.totalBookings++;
          analytics.monthlyRevenue += amount;
        }
        
        // YTD bookings
        if (checkIn >= yearStart) {
          analytics.ytdRevenue += amount;
        }
        
        analytics.totalRevenue += amount;
        analytics.totalNights += nights;
        
        // Weekend vs weekday analysis
        const dayOfWeek = checkIn.getDay();
        if (dayOfWeek === 5 || dayOfWeek === 6) {
          analytics.weekendBookings++;
        } else {
          analytics.weekdayBookings++;
        }
        
        // Purpose tracking
        analytics.purposeStats[purpose] = (analytics.purposeStats[purpose] || 0) + 1;
        
        // Room performance
        if (analytics.roomStats[roomNumber]) {
          analytics.roomStats[roomNumber].bookings++;
          analytics.roomStats[roomNumber].revenue += amount;
          analytics.roomStats[roomNumber].nights += nights;
        }
      }
    });
    
    // Calculate derived metrics
    const totalBookingsAll = analytics.weekendBookings + analytics.weekdayBookings;
    analytics.avgDailyRate = analytics.totalNights > 0 ? analytics.totalRevenue / analytics.totalNights : 0;
    analytics.avgBookingValue = totalBookingsAll > 0 ? analytics.totalRevenue / totalBookingsAll : 0;
    analytics.avgStayLength = totalBookingsAll > 0 ? analytics.totalNights / totalBookingsAll : 0;
    
    // Occupancy rate calculation (simplified)
    const daysInMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
    const totalRoomNights = roomData.length * daysInMonth;
    const occupiedNights = bookingData.filter(booking => {
      const checkIn = new Date(booking[this.BOOKING_COL.CHECK_IN_DATE - 1]);
      const status = booking[this.BOOKING_COL.BOOKING_STATUS - 1];
      return checkIn >= monthStart && (status === CONFIG.STATUS.BOOKING.CHECKED_IN || status === CONFIG.STATUS.BOOKING.CHECKED_OUT);
    }).reduce((sum, booking) => sum + (booking[this.BOOKING_COL.NUMBER_OF_NIGHTS - 1] || 0), 0);
    
    analytics.occupancyRate = totalRoomNights > 0 ? Math.round((occupiedNights / totalRoomNights) * 100) : 0;
    analytics.revPAR = totalRoomNights > 0 ? analytics.monthlyRevenue / roomData.length : 0;
    
    // Percentage calculations
    if (totalBookingsAll > 0) {
      analytics.weekendBookings = Math.round((analytics.weekendBookings / totalBookingsAll) * 100);
      analytics.weekdayBookings = 100 - analytics.weekendBookings;
    }
    
    // Top purpose
    analytics.topPurpose = Object.keys(analytics.purposeStats).reduce((a, b) => 
      analytics.purposeStats[a] > analytics.purposeStats[b] ? a : b, 'None');
    
    // Room performance array
    analytics.roomPerformance = Object.values(analytics.roomStats).map(room => ({
      ...room,
      occupancy: room.nights > 0 ? Math.round((room.nights / daysInMonth) * 100) : 0
    }));
    
    return analytics;
  },
  
  /**
   * Send check-in reminders for today's arrivals
   */
  sendCheckInReminders: function() {
    try {
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      
      const bookingData = SheetManager.getAllData(CONFIG.SHEETS.GUEST_BOOKINGS);
      let sentCount = 0;
      
      bookingData.forEach(row => {
        const checkIn = new Date(row[this.BOOKING_COL.CHECK_IN_DATE - 1]);
        checkIn.setHours(0, 0, 0, 0);
        
        if (checkIn.getTime() === today.getTime() && 
            row[this.BOOKING_COL.BOOKING_STATUS - 1] === CONFIG.STATUS.BOOKING.CONFIRMED) {
          
          const email = row[this.BOOKING_COL.EMAIL - 1];
          if (email) {
            EmailManager.sendGuestCheckInReminder(email, {
              guestName: row[this.BOOKING_COL.GUEST_NAME - 1],
              roomNumber: row[this.BOOKING_COL.ROOM_NUMBER - 1],
              checkInDate: Utils.formatDate(checkIn, 'MMMM dd, yyyy'),
              propertyName: CONFIG.SYSTEM.PROPERTY_NAME
            });
            sentCount++;
          }
        }
      });
      
      SpreadsheetApp.getUi().alert(
        'Check-in Reminders Sent',
        `Sent ${sentCount} check-in reminder(s).`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      
    } catch (error) {
      handleSystemError(error, 'sendCheckInReminders');
    }
  },
  
  /**
   * Get guest room occupancy data for calendar
   */
  getOccupancyData: function(startDate, endDate) {
    const bookingData = SheetManager.getAllData(CONFIG.SHEETS.GUEST_BOOKINGS);
    const occupancyData = {};
    
    bookingData.forEach(booking => {
      const checkIn = new Date(booking[this.BOOKING_COL.CHECK_IN_DATE - 1]);
      const checkOut = new Date(booking[this.BOOKING_COL.CHECK_OUT_DATE - 1]);
      const status = booking[this.BOOKING_COL.BOOKING_STATUS - 1];
      
      if (status === CONFIG.STATUS.BOOKING.CONFIRMED || 
          status === CONFIG.STATUS.BOOKING.CHECKED_IN || 
          status === CONFIG.STATUS.BOOKING.CHECKED_OUT) {
        
        const currentDate = new Date(checkIn);
        while (currentDate < checkOut && currentDate <= endDate) {
          if (currentDate >= startDate) {
            const dateKey = Utils.formatDate(currentDate, 'yyyy-MM-dd');
            const roomNumber = booking[this.BOOKING_COL.ROOM_NUMBER - 1];
            
            if (!occupancyData[dateKey]) {
              occupancyData[dateKey] = {};
            }
            
            occupancyData[dateKey][roomNumber] = {
              guest: booking[this.BOOKING_COL.GUEST_NAME - 1],
              status: status,
              bookingId: booking[this.BOOKING_COL.BOOKING_ID - 1]
            };
          }
          currentDate.setDate(currentDate.getDate() + 1);
        }
      }
    });
    
    return occupancyData;
  }
};

Logger.log('GuestManager module loaded successfully');
