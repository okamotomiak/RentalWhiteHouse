// FormManager.gs - Google Forms Management System
// Handles creation and management of Google Forms for White House Rentals

const FormManager = {
  
  /**
   * Create all system forms with proper field validation
   */
  createAllSystemForms: function() {
    const ui = SpreadsheetApp.getUi();
    
    try {
      const forms = [];
      
      // Create Tenant Application Form
      const tenantForm = this.createTenantApplicationForm();
      forms.push({
        name: 'Tenant Application',
        editUrl: tenantForm.getEditUrl(),
        publishedUrl: tenantForm.getPublishedUrl(),
        id: tenantForm.getId()
      });
      
      // Create Move-Out Request Form
      const moveOutForm = this.createMoveOutRequestForm();
      forms.push({
        name: 'Move-Out Request',
        editUrl: moveOutForm.getEditUrl(),
        publishedUrl: moveOutForm.getPublishedUrl(),
        id: moveOutForm.getId()
      });
      
      // Setup form triggers for automatic processing
      this.setupFormTriggers(tenantForm, moveOutForm);
      
      // Store form information for future reference
      this.storeFormInformation(forms);
      
      ui.alert(
        'Forms Created Successfully!',
        `Created ${forms.length} forms for Belvedere White House Rental:\n\n` +
        forms.map(form => `• ${form.name}`).join('\n') +
        '\n\nForms are now connected to your spreadsheet and will automatically process submissions.',
        ui.ButtonSet.OK
      );
      
      return forms;
      
    } catch (error) {
      ui.alert('Error', `Failed to create forms: ${error.message}`, ui.ButtonSet.OK);
      throw error;
    }
  },
  
  /**
   * Create comprehensive tenant application form
   */
  createTenantApplicationForm: function() {
    const form = FormApp.create('Tenant Application - Belvedere White House Rental');
    
    form.setDescription(`
Welcome to Belvedere White House Rental!

Please complete this application thoroughly. All fields marked with * are required.

We will review your application within 2-3 business days and contact you with our decision.

Thank you for your interest in our community!
    `);
    
    // Basic Information Section
    form.addSectionHeaderItem()
      .setTitle('Personal Information')
      .setHelpText('Please provide your basic contact information');
    
    form.addTextItem()
      .setTitle('Full Name *')
      .setRequired(true)
      .setHelpText('First and Last Name');
    
    form.addTextItem()
      .setTitle('Email Address *')
      .setRequired(true)
      .setValidation(FormApp.createTextValidation()
        .setHelpText('Please enter a valid email address')
        .requireTextIsEmail()
        .build());
    
    form.addTextItem()
      .setTitle('Phone Number *')
      .setRequired(true)
      .setHelpText('Primary contact number');
    
    // Current Housing Section
    form.addSectionHeaderItem()
      .setTitle('Current Housing Information');
    
    form.addParagraphTextItem()
      .setTitle('Current Address *')
      .setRequired(true)
      .setHelpText('Include street address, city, state, and zip code');
    
    // Desired Housing Section
    form.addSectionHeaderItem()
      .setTitle('Desired Housing');
    
    form.addDateItem()
      .setTitle('Desired Move-in Date *')
      .setRequired(true)
      .setHelpText('When would you like to move in?');
    
    form.addMultipleChoiceItem()
      .setTitle('Preferred Room *')
      .setRequired(true)
      .setChoices([
        form.createChoice('Any available room'),
        form.createChoice('Room 101'),
        form.createChoice('Room 102'),
        form.createChoice('Room 103'),
        form.createChoice('Room 104'),
        form.createChoice('Room 201'),
        form.createChoice('Room 202'),
        form.createChoice('Room 203'),
        form.createChoice('Room 204')
      ])
      .setHelpText('Select your preferred room or any available');
    
    // Employment Section
    form.addSectionHeaderItem()
      .setTitle('Employment Information');
    
    form.addMultipleChoiceItem()
      .setTitle('Employment Status *')
      .setRequired(true)
      .setChoices([
        form.createChoice('Full-time employed'),
        form.createChoice('Part-time employed'),
        form.createChoice('Self-employed'),
        form.createChoice('Student'),
        form.createChoice('Retired'),
        form.createChoice('Unemployed'),
        form.createChoice('Other')
      ]);
    
    form.addTextItem()
      .setTitle('Employer/School Name')
      .setHelpText('Current employer or educational institution');
    
    form.addTextItem()
      .setTitle('Monthly Income *')
      .setRequired(true)
      .setHelpText('Gross monthly income in dollars (numbers only)')
      .setValidation(FormApp.createTextValidation()
        .setHelpText('Please enter a valid number')
        .requireNumber()
        .build());
    
    // References Section
    form.addSectionHeaderItem()
      .setTitle('References')
      .setHelpText('Please provide at least one reference');
    
    form.addParagraphTextItem()
      .setTitle('Reference 1 (Required) *')
      .setRequired(true)
      .setHelpText('Name, relationship, phone number, and email if available');
    
    form.addParagraphTextItem()
      .setTitle('Reference 2 (Optional)')
      .setHelpText('Name, relationship, phone number, and email if available');
    
    // Emergency Contact
    form.addSectionHeaderItem()
      .setTitle('Emergency Contact');
    
    form.addParagraphTextItem()
      .setTitle('Emergency Contact Information *')
      .setRequired(true)
      .setHelpText('Name, relationship, and phone number');
    
    // Additional Information
    form.addSectionHeaderItem()
      .setTitle('Additional Information');
    
    form.addParagraphTextItem()
      .setTitle('Tell us about yourself')
      .setHelpText('Brief description of yourself, lifestyle, interests, etc.');
    
    form.addParagraphTextItem()
      .setTitle('Special Needs or Requests')
      .setHelpText('Any accessibility needs or special accommodation requests');
    
    // File Upload for Proof of Income
    form.addFileUploadItem()
      .setTitle('Proof of Income *')
      .setRequired(true)
      .setHelpText('Upload recent pay stub, bank statement, or employment letter (PDF, JPG, PNG)')
      .setFolderName('Belvedere White House Rental - Applications');
    
    // Agreement
    form.addCheckboxItem()
      .setTitle('Application Agreement *')
      .setRequired(true)
      .setChoices([
        form.createChoice('I certify that all information provided is true and complete. I understand that false information may result in denial of my application.')
      ]);
    
    // Connect form to spreadsheet
    form.setDestination(FormApp.DestinationType.SPREADSHEET, SpreadsheetApp.getActiveSpreadsheet().getId());
    
    return form;
  },
  
  /**
   * Create comprehensive move-out request form
   */
  createMoveOutRequestForm: function() {
    const form = FormApp.create('Move-Out Request - Belvedere White House Rental');
    
    form.setDescription(`
30-Day Move-Out Notice

Please submit this form at least 30 days before your intended move-out date as required by your lease agreement.

We will contact you within 24 hours to schedule your final inspection and coordinate the move-out process.

Thank you for being a valued resident!
    `);
    
    // Tenant Information
    form.addSectionHeaderItem()
      .setTitle('Tenant Information');
    
    form.addTextItem()
      .setTitle('Tenant Name *')
      .setRequired(true)
      .setHelpText('Your full name as it appears on the lease');
    
    form.addTextItem()
      .setTitle('Email Address *')
      .setRequired(true)
      .setValidation(FormApp.createTextValidation()
        .requireTextIsEmail()
        .build());
    
    form.addTextItem()
      .setTitle('Phone Number *')
      .setRequired(true);
    
    form.addMultipleChoiceItem()
      .setTitle('Room Number *')
      .setRequired(true)
      .setChoices([
        form.createChoice('Room 101'),
        form.createChoice('Room 102'),
        form.createChoice('Room 103'),
        form.createChoice('Room 104'),
        form.createChoice('Room 201'),
        form.createChoice('Room 202'),
        form.createChoice('Room 203'),
        form.createChoice('Room 204')
      ]);
    
    // Move-Out Details
    form.addSectionHeaderItem()
      .setTitle('Move-Out Details');
    
    form.addDateItem()
      .setTitle('Planned Move-Out Date *')
      .setRequired(true)
      .setHelpText('Must be at least 30 days from today');
    
    form.addParagraphTextItem()
      .setTitle('Forwarding Address *')
      .setRequired(true)
      .setHelpText('Complete address where security deposit refund should be mailed');
    
    form.addMultipleChoiceItem()
      .setTitle('Primary Reason for Moving *')
      .setRequired(true)
      .setChoices([
        form.createChoice('Job relocation'),
        form.createChoice('Buying a home'),
        form.createChoice('Moving closer to family'),
        form.createChoice('Found different housing'),
        form.createChoice('Financial reasons'),
        form.createChoice('Dissatisfied with property'),
        form.createChoice('Other')
      ]);
    
    form.addParagraphTextItem()
      .setTitle('Additional Details')
      .setHelpText('Please explain your reason for moving in more detail');
    
    // Feedback Section
    form.addSectionHeaderItem()
      .setTitle('Feedback (Optional but Appreciated)');
    
    form.addScaleItem()
      .setTitle('Overall Satisfaction')
      .setLabels('Very Dissatisfied', 'Very Satisfied')
      .setBounds(1, 5);
    
    form.addParagraphTextItem()
      .setTitle('What did you like most about living here?');
    
    form.addParagraphTextItem()
      .setTitle('What could we improve?');
    
    form.addMultipleChoiceItem()
      .setTitle('Would you recommend us to others?')
      .setChoices([
        form.createChoice('Definitely'),
        form.createChoice('Probably'),
        form.createChoice('Not sure'),
        form.createChoice('Probably not'),
        form.createChoice('Definitely not')
      ]);
    
    // Connect form to spreadsheet
    form.setDestination(FormApp.DestinationType.SPREADSHEET, SpreadsheetApp.getActiveSpreadsheet().getId());
    
    return form;
  },
  
  /**
   * Setup automatic form processing triggers
   */
  setupFormTriggers: function(tenantForm, moveOutForm) {
    try {
      // Create triggers for form submissions
      ScriptApp.newTrigger('processTenantApplicationSubmission')
        .onFormSubmit()
        .onCreate(tenantForm.getId());
      
      ScriptApp.newTrigger('processMoveOutRequestSubmission')
        .onFormSubmit()
        .onCreate(moveOutForm.getId());
      
      Logger.log('Form triggers created successfully');
      
    } catch (error) {
      Logger.log(`Error setting up form triggers: ${error.toString()}`);
    }
  },
  
  /**
   * Store form information in the system
   */
  storeFormInformation: function(forms) {
    try {
      // Store form URLs in Documents sheet for easy access
      forms.forEach(form => {
        const documentData = [
          Utils.generateId('FORM'),
          'Google Form',
          form.name,
          'Form',
          form.id,
          form.publishedUrl,
          new Date(),
          CONFIG.SYSTEM.MANAGER_EMAIL,
          '',
          'Active',
          'Public',
          'form,application,tenant',
          `Edit URL: ${form.editUrl}`
        ];
        
        SheetManager.addRow(CONFIG.SHEETS.DOCUMENTS, documentData);
      });
      
      Logger.log('Form information stored successfully');
      
    } catch (error) {
      Logger.log(`Error storing form information: ${error.toString()}`);
    }
  },
  
  /**
   * Show all form links with edit and published URLs
   */
  showAllFormLinks: function() {
    try {
      // Get form information from Documents sheet
      const documents = SheetManager.getAllData(CONFIG.SHEETS.DOCUMENTS);
      const formDocs = documents.filter(doc => doc[1] === 'Google Form');
      
      let formLinksHtml = '';
      
      if (formDocs.length > 0) {
        formLinksHtml = formDocs.map(doc => `
          <div style="background: #f8f9fa; padding: 15px; border-radius: 8px; margin-bottom: 15px;">
            <h4 style="margin: 0 0 10px 0; color: #007bff;">${doc[2]}</h4>
            <p><strong>Public Link (share with applicants):</strong><br>
            <a href="${doc[5]}" target="_blank" style="word-break: break-all;">${doc[5]}</a></p>
            <p><strong>Edit Link (for you to modify):</strong><br>
            <a href="${doc[12].replace('Edit URL: ', '')}" target="_blank" style="word-break: break-all;">${doc[12].replace('Edit URL: ', '')}</a></p>
          </div>
        `).join('');
      } else {
        formLinksHtml = '<p style="text-align: center; color: #666;">No forms found. Please create forms first using "Create All Forms".</p>';
      }
      
      const html = HtmlService.createHtmlOutput(`
        <div style="font-family: Arial, sans-serif; padding: 20px;">
          <h2>🔗 Form Links - Belvedere White House Rental</h2>
          
          ${formLinksHtml}
          
          <div style="background: #e3f2fd; padding: 15px; border-radius: 8px; margin-top: 20px;">
            <h4>📋 How to Use These Forms:</h4>
            <ul>
              <li><strong>Public Links:</strong> Share these with prospective tenants and current residents</li>
              <li><strong>Edit Links:</strong> Use these to modify form questions and settings</li>
              <li><strong>Automatic Processing:</strong> Form submissions are automatically added to your spreadsheet</li>
              <li><strong>Email Notifications:</strong> You'll receive email alerts for new submissions</li>
            </ul>
          </div>
          
          <div style="text-align: center; margin-top: 20px;">
            <button onclick="google.script.run.createAllSystemForms()" style="padding: 10px 20px; background: #28a745; color: white; border: none; border-radius: 4px;">
              Create New Forms
            </button>
          </div>
        </div>
      `)
        .setWidth(700)
        .setHeight(600);
      
      SpreadsheetApp.getUi().showModalDialog(html, 'Form Links');
      
    } catch (error) {
      handleSystemError(error, 'showAllFormLinks');
    }
  },
  
  /**
   * Process tenant application form submission
   */
  processTenantApplicationSubmission: function(e) {
    try {
      Logger.log('Processing tenant application submission...');
      
      if (!e || !e.values) {
        Logger.log('No form data received');
        return;
      }
      
      const formData = e.values;
      const timestamp = new Date();
      
      // Generate application ID
      const applicationId = Utils.generateId('APP');
      
      // Extract form data (adjust indices based on your form structure)
      const applicantName = formData[1] || '';
      const applicantEmail = formData[2] || '';
      const applicantPhone = formData[3] || '';
      const currentAddress = formData[4] || '';
      const moveInDate = formData[5] || '';
      const preferredRoom = formData[6] || '';
      const employmentStatus = formData[7] || '';
      
      // Send confirmation email to applicant
      if (applicantEmail && Utils.isValidEmail(applicantEmail)) {
        const subject = 'Application Received - Belvedere White House Rental';
        const body = `Dear ${applicantName},

Thank you for your application to Belvedere White House Rental!

Application Details:
• Application ID: ${applicationId}
• Submitted: ${timestamp.toLocaleString()}
• Preferred Room: ${preferredRoom}
• Desired Move-in: ${moveInDate}

We will review your application and contact you within 2-3 business days with our decision.

If you have any questions, please don't hesitate to contact us.

Best regards,
Belvedere White House Rental Management
Email: ${CONFIG.SYSTEM.MANAGER_EMAIL}`;

        MailApp.sendEmail(applicantEmail, subject, body);
      }
      
      // Send notification to manager
      const managerSubject = 'New Tenant Application - Belvedere White House Rental';
      const managerBody = `A new tenant application has been received:

Application ID: ${applicationId}
Applicant: ${applicantName}
Email: ${applicantEmail}
Phone: ${applicantPhone}
Preferred Room: ${preferredRoom}
Move-in Date: ${moveInDate}
Employment: ${employmentStatus}

Please review the application in your management system.

Submitted: ${timestamp.toLocaleString()}`;

      MailApp.sendEmail(CONFIG.SYSTEM.MANAGER_EMAIL, managerSubject, managerBody);
      
      Logger.log(`Processed tenant application: ${applicationId} for ${applicantName}`);
      
    } catch (error) {
      Logger.log(`Error processing tenant application: ${error.toString()}`);
    }
  },
  
  /**
   * Process move-out request form submission
   */
  processMoveOutRequestSubmission: function(e) {
    try {
      Logger.log('Processing move-out request submission...');
      
      if (!e || !e.values) {
        Logger.log('No form data received');
        return;
      }
      
      const formData = e.values;
      const timestamp = new Date();
      
      // Generate request ID
      const requestId = Utils.generateId('MO');
      
      // Extract form data
      const tenantName = formData[1] || '';
      const tenantEmail = formData[2] || '';
      const tenantPhone = formData[3] || '';
      const roomNumber = formData[4] || '';
      const moveOutDate = formData[5] || '';
      const forwardingAddress = formData[6] || '';
      const reason = formData[7] || '';
      
      // Send confirmation email to tenant
      if (tenantEmail && Utils.isValidEmail(tenantEmail)) {
        const subject = 'Move-Out Request Received - Belvedere White House Rental';
        const body = `Dear ${tenantName},

We have received your 30-day move-out notice.

Move-Out Details:
• Request ID: ${requestId}
• Room: ${roomNumber}
• Planned Move-Out Date: ${moveOutDate}
• Forwarding Address: ${forwardingAddress}

Next Steps:
1. We will contact you within 24 hours to schedule your final inspection
2. Please ensure all personal belongings are removed by your move-out date
3. All keys must be returned during the final inspection
4. Security deposit refund will be processed within 30 days

Thank you for being a valued resident. We wish you all the best in your future endeavors!

Best regards,
Belvedere White House Rental Management
Email: ${CONFIG.SYSTEM.MANAGER_EMAIL}`;

        MailApp.sendEmail(tenantEmail, subject, body);
      }
      
      // Send notification to manager
      const managerSubject = 'Move-Out Request Received - Belvedere White House Rental';
      const managerBody = `A move-out request has been submitted:

Request ID: ${requestId}
Tenant: ${tenantName}
Email: ${tenantEmail}
Phone: ${tenantPhone}
Room: ${roomNumber}
Move-Out Date: ${moveOutDate}
Reason: ${reason}
Forwarding Address: ${forwardingAddress}

Action Required:
• Contact tenant within 24 hours
• Schedule final inspection
• Prepare move-out documentation

Submitted: ${timestamp.toLocaleString()}`;

      MailApp.sendEmail(CONFIG.SYSTEM.MANAGER_EMAIL, managerSubject, managerBody);
      
      Logger.log(`Processed move-out request: ${requestId} for ${tenantName}`);
      
    } catch (error) {
      Logger.log(`Error processing move-out request: ${error.toString()}`);
    }
  }
};

// Global functions for form triggers
function processTenantApplicationSubmission(e) {
  FormManager.processTenantApplicationSubmission(e);
}

function processMoveOutRequestSubmission(e) {
  FormManager.processMoveOutRequestSubmission(e);
}

Logger.log('FormManager module loaded successfully');
