var subPanelCSS = {
  backgroundColor: '#f7f7f7',
  border: '1px solid grey',
  padding: '10px'
};

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [{name: "Xero Settings", functionName: "xeroSetup"}, //This will be for adding user API settings.
                     //{name: "Account Type", functionName: "accountType"}, // Select Payable or Receivables
                     {name: "Manual Upload Invoices", functionName: "uploadInvoices"}, // Push through manaual upload of prices.
                     {name: "Manual Download Invoices", functionName: "handleInvoicesDownload"}, // Push through manual download of orders.
                     {name: "Manual Upload Payments", functionName: "uploadPayments"}, // Push through manual upload of order statuses.
                     {name: "Manual Download Payments", functionName: "handlePaymentsDownload"},
                     //{name: "Automation Setup", functionName: "automation"}, // Automation setup for automatic upload and download of Invoices, download of Invoices, upload of Payments, and Download of Payments. If each function could be listed with own trigger setup.
                     {name: "Clear Data", functionName: "clearData"},
                     //{name: "Test", functionName: "downloadInvoiceOptions"}
                    ];
  ss.addMenu("XERO", menuEntries);  
}

function doGet(e) {
  // https://script.google.com/macros/s/AKfycbyoJ09MtcAJFB-v31CH3sA2L5atNr4jtGW3cYWlqqpvGg_kBEY/exec?oauth_token=KNDX9LRNOC0KNAEB4PSVPDINTGSC6F&oauth_verifier=1144716&org=zX0A5sD18KgGYSSotqTKvG
  // Step 3. Exchange verified Request token for Access Token
  // var p = PropertiesService.getScriptProperties().getProperties();
  Xero.getSettings();
  var payload = 
      {
        "oauth_consumer_key": Xero.getProperty('consumerKey'),
        "oauth_token": e.parameter.oauth_token,
        "oauth_signature_method": "PLAINTEXT",
        "oauth_signature": encodeURIComponent(Xero.getProperty('consumerSecret') + '&' + Xero.getProperty('requestTokenSecret')),
        "oauth_timestamp": ((new Date().getTime())/1000).toFixed(0),
        "oauth_nonce": generateRandomString(Math.floor(Math.round(25))),
        "oauth_version": "1.0",
        "oauth_verifier": e.parameter.oauth_verifier
      };
  var options = {"method": "post", "payload": payload, muteHttpExceptions: true};
  try {
    var response = UrlFetchApp.fetch(accessTokenURL, options);
  }catch(e) {
    Logger.log(e);
    return HtmlService.createHtmlOutput("<html><div>"+ response.getContentText() +"</div></html>");  
  }  
  
  var reoAuthToken = /(oauth_token=)([a-zA-Z0-9]+)/;    
  var tokenMatch = reoAuthToken.exec(response.getContentText());
  var reTokenSecret = /(oauth_token_secret=)([a-zA-Z0-9]+)/;
  var secretMatch = reTokenSecret.exec(response.getContentText())  ;  
    
  if (tokenMatch && tokenMatch[2] != '') {
    PropertiesService.getScriptProperties().setProperty('accessToken', tokenMatch[2]);  
    PropertiesService.getScriptProperties().setProperty('accessTokenSecret', secretMatch[2]);  
    PropertiesService.getScriptProperties().setProperty('isConnected', true);  
    ScriptApp.newTrigger('closeConnection').timeBased().after(30000); 
    return HtmlService.createHtmlOutput("<html><b>Your spreadsheet is now connected to Xero.com for 30 mins.</b></html>");  
  } 
}

function closeConnection() {
  PropertiesService.getScriptProperties().deleteProperty('isConnected');
  PropertiesService.getScriptProperties().setProperty('isConnected', false);  
}

function closeApp() {
  var app = UiApp.getActiveApplication();
  app.close();
  return app;
}

function xeroSetup() {  
  var ss = SpreadsheetApp.getActiveSpreadsheet();    
  var app = UiApp.createApplication();
  app.setHeight(400);
  var scriptProps = PropertiesService.getScriptProperties(); 
  
  var label1 = app.createLabel('XERO Settings').setStyleAttribute('font-weight', 'bold').setStyleAttribute('padding', '5px').setId('label1');        
  var panel1 = app.createVerticalPanel().setId('panel1');
  var grid = app.createGrid(7, 2);
  var absPanel = app.createAbsolutePanel();  
  
  var handler = app.createServerHandler('saveSettings');    
  var clientHandler1 = app.createClientHandler();  
  var clientHandler2 = app.createClientHandler();    
  var clientHandler3 = app.createClientHandler();    
  
  var btnSave = app.createButton('Save Settings', handler);       
  var lblAppType = app.createLabel('Application Type: ');    
  var appTypes = {Private:0, Public:1, Partner:2};
  var listAppType = app.createListBox().setName('appType').addItem('Private').addItem('Public').addItem('Partner').addChangeHandler(clientHandler1).
  addChangeHandler(clientHandler2).addChangeHandler(clientHandler3).setSelectedIndex(appTypes[(scriptProps.getProperty('appType') != null ? scriptProps.getProperty('appType'): 'Private')]);  
  handler.addCallbackElement(listAppType);
  
  var lblAppName = app.createLabel('Application Name: ');  
  var txtAppName = app.createTextBox().setName('userAgent').setWidth("350")
  .setValue((scriptProps.getProperty('userAgent') != null ? scriptProps.getProperty('userAgent'): ""));
  handler.addCallbackElement(txtAppName);
  
  var lblConsumerKey = app.createLabel('Consumer Key: ');  
  var txtConsumerKey = app.createTextBox().setName('consumerKey').setWidth("350")   
  .setValue((scriptProps.getProperty('consumerKey') != null ? scriptProps.getProperty('consumerKey'): ""));
  handler.addCallbackElement(txtConsumerKey);  
  
  var lblConsumerSecret = app.createLabel('Consumer Secret: ');  
  var txtConsumerSecret = app.createTextBox().setName('consumerSecret').setWidth("350")
    .setValue((scriptProps.getProperty('consumerSecret') != null ? scriptProps.getProperty('consumerSecret'): ""));
  handler.addCallbackElement(txtConsumerSecret);
  
  var lblcallBack = app.createLabel('Callback URL:');
  var txtcallBack = app.createTextBox().setName('callBack').setWidth("350")
    .setValue((scriptProps.getProperty('callbackURL') != null ? scriptProps.getProperty('callbackURL'): ""));
  handler.addCallbackElement(txtcallBack);
  
  var lblRSA = app.createLabel('RSA Private Key:');
  var txtareaRSA = app.createTextArea().setName('RSA').setWidth("350").setHeight("150")
    .setValue((scriptProps.getProperty('rsaKey') != null ? scriptProps.getProperty('rsaKey'): ""));  
  
  if (scriptProps.getProperty('appType') == "Private" || scriptProps.getProperty('appType') == null)     
    txtcallBack.setEnabled(false);
  else if (scriptProps.getProperty('appType') == "Public")     
    txtareaRSA.setEnabled(false);
  
  handler.addCallbackElement(txtareaRSA);  
  clientHandler1.validateMatches(listAppType, 'Private').forTargets(txtcallBack).setEnabled(false).forTargets(txtareaRSA).setEnabled(true);
  clientHandler2.validateMatches(listAppType, 'Public').forTargets(txtcallBack).setEnabled(true).forTargets(txtareaRSA).setEnabled(false);
  clientHandler3.validateMatches(listAppType, 'Partner').forTargets(txtcallBack).setEnabled(true).forTargets(txtareaRSA).setEnabled(true);
  
  grid.setBorderWidth(0);
  grid.setWidget(0, 0, lblAppType);
  grid.setWidget(0, 1, listAppType);  
  grid.setWidget(1, 0, lblAppName);
  grid.setWidget(1, 1, txtAppName);
  grid.setWidget(2, 0, lblConsumerKey);
  grid.setWidget(2, 1, txtConsumerKey);  
  grid.setWidget(3, 0, lblConsumerSecret);
  grid.setWidget(3, 1, txtConsumerSecret);
  grid.setWidget(4, 0, lblcallBack);
  grid.setWidget(4, 1, txtcallBack);
  grid.setWidget(5, 0, lblRSA);
  grid.setWidget(5, 1, txtareaRSA);
  grid.setWidget(6, 1, btnSave);  
  panel1.add(grid).setStyleAttributes(subPanelCSS);    
  app.add(label1);
  app.add(panel1);  
  ss.show(app);  
}

function saveSettings(e){
  var ss = SpreadsheetApp.getActiveSpreadsheet();  
  var app = UiApp.getActiveApplication();
  // we have the settings in e.parameter  
  var sProps = PropertiesService.getScriptProperties();    
  var p = e.parameter;  
  var properties = { appType: p.appType, userAgent: p.userAgent, consumerKey: p.consumerKey, consumerSecret: p.consumerSecret, rsaKey: p.RSA, 
                    callbackURL: p.callBack };  
  if (sProps.setProperties(properties))
    ss.toast('Xero Settings saved.', '', 5);  
  else
    ss.toast('Xero settings could not be saved.', '', 5);  
  if (p.appType != 'Private') {
    app.remove(app.getElementById('label1')); 
    app.remove(app.getElementById('panel1'));          
    app.setHeight(50);app.setWidth(100);            
    var link = app.createAnchor('Connect to Xero', true, Xero.connect());
    var handler = app.createServerHandler('closeApp');
    link.addClickHandler(handler);
    app.add(link);
  }
  else
    app.close();
  return app;
}

// Download Functions
function downloadInvoices() {
  downloadOptions('Invoices');
}

function downloadPayments() {
  downloadOptions('Payments');
}

function downloadOptions(item) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();    
  var app = UiApp.createApplication();
  var panel = app.createAbsolutePanel().setStyleAttributes(subPanelCSS).setWidth("475");  
  var lblHeader = app.createLabel('Download ' + item).setStyleAttribute('font-weight', 'bold');
  var gridMain = app.createGrid(4, 2);  
  var optAllHandler = app.createClientHandler();
  var optModifiedAfterHandler = app.createClientHandler();
  var optAll = app.createRadioButton('optAll', 'All').setId('optAll').setValue(true).addClickHandler(optAllHandler);  
  var optModifiedAfter = app.createRadioButton('optModifiedAfter', 'Modified After').setValue(false).addClickHandler(optModifiedAfterHandler);  
  var modifiedAfterDate = app.createDatePicker().setVisible(false).setName('modifiedAfterDate').setValue(new Date());
  
  optAllHandler.forEventSource().setEnabled(true).forTargets(optModifiedAfter).setValue(false).forTargets(modifiedAfterDate).setVisible(false);
  optModifiedAfterHandler.forEventSource().setEnabled(true).forTargets(optAll).setValue(false).forTargets(modifiedAfterDate).setVisible(true);
  
  switch(item) {
    case 'Invoices':
      var clickHandler = app.createServerHandler('handleInvoicesDownload');
      break;
    case 'Payments':
      var clickHandler = app.createServerHandler('handlePaymentsDownload');
      break;
  }
  var btnHandler = app.createClientHandler().forEventSource().setEnabled(false).forTargets(optAll).setEnabled(false).forTargets(optModifiedAfter).setEnabled(false).forTargets(modifiedAfterDate).setEnabled(false);
  var btnDownload = app.createButton('Download', clickHandler).addClickHandler(btnHandler); 
  
  // callback elements
  clickHandler.addCallbackElement(optAll).addCallbackElement(optModifiedAfter).addCallbackElement(modifiedAfterDate);
  gridMain.setWidget(0, 0, optAll);
  gridMain.setWidget(1, 0, optModifiedAfter);  
  gridMain.setWidget(2, 1, modifiedAfterDate);    
  gridMain.setWidget(3, 0, btnDownload);
  
  panel.add(gridMain);  
  app.add(lblHeader);
  app.add(panel);
  
  ss.show(app);
}  

function handleInvoicesDownload(){      
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var sheet    = ss.getSheetByName("Invoices Download");
  var li_sheet = ss.getSheetByName("Download - Invoices Line Items"); 
  
  if (!sheet) ss.insertSheet("Invoices Download");
  if (!li_sheet) ss.insertSheet("Download - Invoices Line Items");  
  
  // Add headers
  if (sheet.getLastRow() == 0) {
    var arrHeaders = ['Type', 'Contact Name', 'Date',	'Due Date', 'Status', 'Line Amount Types', 'Sub Total',	'Total Tax', 'Total', 'Updated Date (UTC)',	
                        'Currency Code', 'Invoice ID', 'Invoice Number', 'Amount Due', 'Amount Paid', 'Amount Credited'];
    sheet.appendRow(arrHeaders).getRange(1, 1, 1, 16).setFontWeight('bold');
  }  
  if (li_sheet.getLastRow() == 0) {    
    var liHeaders = ['Invoice ID', 'Invoice Number', 'Description', 'Quantity',	'Unit Amount', 'Tax Type', 'Tax Amount', 'Line Amount',	'Account Code'];
    li_sheet.appendRow(liHeaders).getRange(1, 1, 1, 9).setFontWeight('bold');
  }    
  
  var pageNo = 1;
  var moreData = true;
  
  try {
    // Connect to Xero
    Xero.connect();   
    
    while (moreData) {
      ss.toast('Downloading page ' + pageNo + ' ...');
      // API call          
      var invoice_info = Xero.fetchData('Invoices', pageNo);
      Logger.log(invoice_info);
      
      if (invoice_info) var invoices = invoice_info.Invoices; 
      else return false;
      
      // If less than 100 records returned, there are no more records
      if (invoices.length < 100) moreData = false; 
      else pageNo++;
      
      // Processing the result and update Invoices data in the sheet    
      for(var i=0; i < invoices.length; i++) {
        var invoice         = invoices[i];      
        var accType         = (invoice.Type != null) ? invoice.Type : "";
        var contactName     = (invoice.Contact != null) ? invoice.Contact.Name : ""; 
        var date            = (invoice.DateString != null) ? invoice.DateString : "";
        var dueDate         = (invoice.DueDateString != null) ? invoice.DueDateString : "";
        var status          = (invoice.Status != null) ? invoice.Status  : "";
        var lineAmountTypes = (invoice.LineAmountTypes != null) ? invoice.LineAmountTypes : "";
        
        // Line Items present in the Middle
        var subTotal       = (invoice.SubTotal != null) ? invoice.SubTotal : "";
        var totalTax       = (invoice.TotalTax != null) ? invoice.TotalTax : "";
        var total          = (invoice.Total != null)    ? invoice.Total    : "";
        var updatedDateUTC = (invoice.UpdatedDateUTC != null) ?  eval('new ' + invoice.UpdatedDateUTC.substr(1, invoice.UpdatedDateUTC.length - 2) + '.toISOString()') : "" ;        
        var currencyCode   = (invoice.CurrencyCode != null) ? invoice.CurrencyCode : "";
        var invoiceID      = (invoice.InvoiceID != null) ? invoice.InvoiceID : "";
        var invoiceNumber  = (invoice.InvoiceNumber != null) ? invoice.InvoiceNumber : "";
        
        // Payments present in the Middle
        var amountDue      = (invoice.AmountDue != null) ? invoice.AmountDue : "";
        var amountPaid     = (invoice.AmountPaid != null) ? invoice.AmountPaid : "";
        var AmountCredited = (invoice.AmountCredited != null) ? invoice.AmountCredited : 0;
        
        // Add Line Items to the Invoice in the end, because all other fields for invoice will be common       
        var arr_line_items = [];
        
        var lineItems =  (invoice.LineItems != null) ? invoice.LineItems : [];
        for(var k=0; k < lineItems.length; k++){
          var item         = lineItems[k];         
          var description  = (item.Description != null ) ? item.Description : "";
          var quantity     = (item.Quantity != null) ? item.Quantity : "";
          var unitAmount   = (item.UnitAmount != null) ? item.UnitAmount : "";
          var taxType      = (item.TaxType != null) ? item.TaxType : "";
          var taxAmount    = (item.TaxAmount != null) ? item.TaxAmount : "";
          var lineAmount   = (item.LineAmount != null) ? item.LineAmount : "";
          var accountCode  = (item.AccountCode != null) ? item.AccountCode : "";
          
          var tmp = [invoiceID,invoiceNumber,description,quantity,unitAmount,taxType,taxAmount,lineAmount,accountCode];
          li_sheet.appendRow(tmp);
        }      
        // Add data in the Spreadsheet    
        var arr =[accType, contactName, date, dueDate, status, lineAmountTypes, subTotal, totalTax, total, updatedDateUTC, currencyCode, invoiceID, invoiceNumber,amountDue, amountPaid, AmountCredited ];
        sheet.appendRow(arr);
      } 
    } 
  }
  catch(e) {
    Browser.msgBox('handleInvoicesDownload: ' + e.message);
  }  
}

function handlePaymentsDownload(){
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Payments Download"); 
  
  if (!sheet) {
    ss.insertSheet("Payments Download");
    sheet = ss.getActiveSheet();
  }
  sheet.activate();
  
  // Add headers
  if (sheet.getLastRow() == 0) {    
    var arrHeaders = ['Payment ID', 'Date', 'Amount', 'Currency Rate', 'Payment Type', 'Status', 'Update Date (UTC)', 'Account ID', 'Contact ID', 'Contact Name',
                      'Invoice Type', 'Invoice ID', 'Invoice Number'];
    sheet.appendRow(arrHeaders).getRange(1, 1, 1, 13).setFontWeight('bold');
  }
  
  var pageNo = 1;
  var moreData = true;  
  
  try {
    // Connect to Xero
    Xero.connect();
    
    ss.toast("Downloading page " + pageNo + "...");
    while (moreData) {
      // API call
      var payments_info = Xero.fetchData('Payments', pageNo); // Data should retrieve in JSON format
      
      if (payments_info) var payments = payments_info.Payments;  
      else return app;    
      
      // Processing the result and update Invoices data in the sheet
      for(var i=0; i < payments.length; i++){
        var payment        = payments[i];      
        var paymentID      = (payment.PaymentID != null) ? payment.PaymentID : "";      
        var date           = (payment.Date != null) ?  eval('new ' + payment.Date.substr(1, payment.Date.length - 2) + '.toISOString()') : "" ;        
        var amount         = (payment.Amount != null) ? payment.Amount : "";
        var currencyRate   = (payment.CurrencyRate != null) ? payment.CurrencyRate : "";
        var paymentType    = (payment.PaymentType != null) ? payment.PaymentType : "";
        var status         = (payment.Status != null) ? payment.Status : "";      
        var updatedDateUTC = (payment.UpdatedDateUTC != null) ?  eval('new ' + payment.UpdatedDateUTC.substr(1, payment.UpdatedDateUTC.length - 2) + '.toISOString()') : "" ;        
        
        var accountId      = "";
        var account        = payment.Account;
        if (account != null && account != undefined && account != ""){
          accountId        = ( account.AccountID != null ) ? account.AccountID : "";
        }
        
        var contactID = "", inv_type = "", inv_id = "", inv_number ="", contactName="";
        var invoice = payment.Invoice;
        if (invoice != undefined && invoice != null &&  invoice != "") {
          inv_type         = (invoice.Type != null)       ? invoice.Type : "";
          inv_id           = (invoice.InvoiceID != null)       ? invoice.InvoiceID : "";
          inv_number       = (invoice.InvoiceNumber != null)       ? invoice.InvoiceNumber : "";
          var contact      = invoice.Contact ;
          
          if(contact != null && contact != undefined && contact != ""){
            contactID      = (contact.ContactID != null) ? contact.ContactID : "";
            contactName    = (contact.Name != null) ? contact.Name : "";
          }
        }      
        // Add data in the Spreadsheet
        var arr = [paymentID, date, amount, currencyRate, paymentType, status, updatedDateUTC, accountId, contactID, contactName, inv_type, inv_id, inv_number];
        sheet.appendRow(arr);      
      }
      
      // More data?
      if (payments.length < 100) moreData = false;
      else pageNo++;
    }    
  }
  catch(e) {
    Browser.msgBox(e.message);
  }  
}


// Upload Functions
function uploadInvoices(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet   = ss.getSheetByName("Invoices Upload");
  var l_sheet = ss.getSheetByName("Upload - Invoice Line Items");
  
  if (!sheet) {
    ss.insertSheet('Invoices Upload');
    sheet   = ss.getSheetByName("Invoices Upload");
  }
  
  if (!l_sheet) {
    ss.insertSheet('Upload - Invoice Line Items');  
    l_sheet = ss.getSheetByName("Upload - Invoice Line Items");
  }
  
  if (sheet.getLastRow() == 0) {    
    var arrHeaders = ['Invoice Number', 'Invoice Type', 'Contact Name',	'Contact Number', 'Issue Date', 'Due Date', 'Line Amount Types', 'Reference', 'Bid', 'URL',	
                        'Currency Code', 'Status', 'Sent to Contact', 'Sub Total', 'Total Tax', 'Total'];
    sheet.appendRow(arrHeaders).getRange(1, 1, 1, 16).setFontWeight('bold');    
  }
  if (l_sheet.getLastRow() == 0) {    
    var liHeaders = ['Invoice Number', 'Description', 'Quantity',	'Unit Amount', 'Tax Type', 'Tax Amount', 'Line Amount',	'Account Code', 'Region'];
    l_sheet.appendRow(liHeaders).getRange(1, 1, 1, 9).setFontWeight('bold');    
  }    
  
  sheet.activate();
  
  if (sheet.getLastRow() <= 1 || l_sheet.getLastRow() <= 1) {
    Browser.msgBox('Error: No data to upload');
    return false;
  }    
  
  var l_lrow = l_sheet.getDataRange().getLastRow();
  var l_data = l_sheet.getDataRange().getValues();
  
  var lrow = sheet.getDataRange().getLastRow();
  var lcol = 17;
  
  var data = sheet.getRange(2, 1, lrow-1, lcol).getValues();
  var payload  = "<Invoices>"; 
  for(var i=0; i < data.length; i++) {

    var inv_number           = data[i][0];
    var inv_type             = data[i][1];
    var contact_name         = data[i][2];
    var contact_number       = data[i][3];
    var issue_date           = data[i][4];
    var due_date             = data[i][5];
    var line_amount_types    = data[i][6];
    var reference            = data[i][7];
    var bid                  = data[i][8];
    var url                  = data[i][9];
    var currency_code        = data[i][10];
    var status               = data[i][11];
    var sent_to_contact      = data[i][12];
    var sub_total            = data[i][13];
    var total_tax            = data[i][14];
    var total                = data[i][15];

    payload += "<Invoice>";
    payload += "<InvoiceNumber>" + inv_number + "</InvoiceNumber>";
    payload += "<Type>" + inv_type + "</Type>";
    payload += "<Contact>";
    payload += "<Name>" + contact_name + "</Name>";
    payload += "<ContactNumber>" + contact_number + "</ContactNumber>";
    payload += "</Contact>";
 
    // Removed the Address related fields in Contact. Add them as required
    payload += "<Date>" + Utilities.formatDate(issue_date, "GMT", "yyyy-MM-dd") + "</Date>";
    payload += "<DueDate>" + Utilities.formatDate(due_date, "GMT", "yyyy-MM-dd") + "</DueDate>";
    payload += "<LineAmountTypes>" + line_amount_types + "</LineAmountTypes>";
    if(inv_type == "ACCREC"){
      payload += "<Reference>" + reference + "</Reference>";
    }
    payload += "<Url>" + url + "</Url>";
    payload += "<CurrencyCode>" + currency_code + "</CurrencyCode>";
    payload += "<Status>" + status + "</Status>";
    payload += "<SentToContact>" + sent_to_contact + "</SentToContact>";
    payload += "<SubTotal>" + sub_total + "</SubTotal>";
    payload += "<TotalTax>" + total_tax + "</TotalTax>";
    payload += "<Total>" + total + "</Total>";

    payload += "<LineItems>";
    for(var j=2; j < l_data.length; j++){
      var l_inv_number = l_data[j][0];
      if(l_inv_number != undefined && l_inv_number != null && l_inv_number != "" && l_inv_number == inv_number){
        payload += "<LineItem>";
        
        var l_desc        = l_data[j][1];
        var l_quantity    = l_data[j][2];
        var l_unitamount  = l_data[j][3];
        var l_taxtype     = l_data[j][4];
        var l_taxamount   = l_data[j][5];
        var l_lineamount  = l_data[j][6];
        var l_accountcode = l_data[j][7];
        var l_region      = l_data[j][8];
        
        payload += "<Description>" + l_desc + "</Description>";
        payload += "<Quantity>" + l_quantity + "</Quantity>";
        payload += "<UnitAmount>" + l_unitamount + "</UnitAmount>";
        payload += "<TaxType>" + l_taxtype + "</TaxType>";
        payload += "<TaxAmount>" + l_taxamount + "</TaxAmount>";
        payload += "<LineAmount>" + l_lineamount + "</LineAmount>";
        payload += "<AccountCode>" + l_accountcode + "</AccountCode>";
        payload += "<Region>" + l_region + "</Region>";
        
        payload += "</LineItem>";
      }      
    }
    payload += "</LineItems>";
    payload += "</Invoice>";
  }
  payload  += "</Invoices>";
  
  try {    
    Xero.connect();
    var response = Xero.uploadData('Invoices', payload);
    if (response) {
      // request successful, Log failed invoices
      var invoices = response.Invoices;    
      logFailedInvoices(invoices);
    }  
  }
  catch(e) {
    Browser.msgBox(e.message);
  }    
}

function logFailedInvoices(invoices){      
  var ss           = SpreadsheetApp.getActiveSpreadsheet();      
  var errorSheet   = ss.getSheetByName("Error Invoices Upload");
  var errorLISheet = ss.getSheetByName("Error Upload - Invoice Line Items"); 
  
  if (!errorSheet) ss.insertSheet("Error Invoices Upload");
  if (!errorLISheet) ss.insertSheet("Error Upload - Invoice Line Items");  
  
  // Add headers
  if (errorSheet.getLastRow() == 0) {
    var arrHeaders = ['Type', 'Contact Name', 'Date',	'Due Date', 'Status', 'Line Amount Types', 'Sub Total',	'Total Tax', 'Total', 'Updated Date (UTC)',	
                        'Currency Code', 'Invoice ID', 'Invoice Number', 'Amount Due', 'Amount Paid', 'Amount Credited'];
    errorSheet.appendRow(arrHeaders).getRange(1, 1, 1, 16).setFontWeight('bold');
  }  
  if (errorLISheet.getLastRow() == 0) {    
    var liHeaders = ['Invoice ID', 'Description', 'Quantity',	'Unit Amount', 'Tax Type', 'Tax Amount', 'Line Amount',	'Account Code'];
    errorLISheet.appendRow(liHeaders).getRange(1, 1, 1, 8).setFontWeight('bold');
  }    
  
  // Processing the result and update Invoices data in the sheet
  var failedInvoices =0;
  for(var i=0; i < invoices.length; i++){
    var invoice         = invoices[i];   
    var StatusAttributeString = (invoice.StatusAttributeString != null) ? invoice.StatusAttributeString : "";
    if (StatusAttributeString == "ERROR") { // an error occured for this invoice
      failedInvoices++;
      var accType         = (invoice.Type != null) ? invoice.Type : "";
      var contactName     = (invoice.Contact.Name != null) ? invoice.Contact.Name : ""; 
      var date            = (invoice.DateString != null) ? invoice.DateString : "";
      var dueDate         = (invoice.DueDateString != null) ? invoice.DueDateString : "";
      var status          = (invoice.Status != null) ? invoice.Status  : "";
      var lineAmountTypes = (invoice.LineAmountTypes != null) ? invoice.LineAmountTypes : "";  
      
      // Line Items present in the Middle
      var subTotal       = (invoice.SubTotal != null) ? invoice.SubTotal : "";
      var totalTax       = (invoice.TotalTax != null) ? invoice.TotalTax : "";
      var total          = (invoice.Total != null)    ? invoice.Total    : "";
      var updatedDateUTC = (invoice.UpdatedDateUTC != null) ?  eval('new ' + invoice.UpdatedDateUTC.substr(1, invoice.UpdatedDateUTC.length - 2) + '.toISOString()') : "" ;        
      var currencyCode   = (invoice.CurrencyCode != null) ? invoice.CurrencyCode : "";
      var invoiceID      = (invoice.InvoiceID != null) ? invoice.InvoiceID : "";
      var invoiceNumber  = (invoice.InvoiceNumber != null) ? invoice.InvoiceNumber : "";
      
      // Payments present in the Middle
      var amountDue      = (invoice.AmountDue != null) ? invoice.AmountDue : "";
      var amountPaid     = (invoice.AmountPaid != null) ? invoice.AmountPaid : "";
      var AmountCredited = (invoice.AmountCredited != null) ? invoice.AmountCredited : 0;
      
      // Add Line Items to the Invoice in the end, because all other fields for invoice will be common       
      var arr_line_items = [];    
      
      var lineItems =  (invoice.LineItems != null) ? invoice.LineItems : [];
      for(var k=0; k < lineItems.length; k++){
        var item         = lineItems[k];         
        var description  = (item.Description != null ) ? item.Description : "";
        var quantity     = (item.Quantity != null) ? item.Quantity : "";
        var unitAmount   = (item.UnitAmount != null) ? item.UnitAmount : "";
        var taxType      = (item.TaxType != null) ? item.TaxType : "";
        var taxAmount    = (item.TaxAmount != null) ? item.TaxAmount : "";
        var lineAmount   = (item.LineAmount != null) ? item.LineAmount : "";
        var accountCode  = (item.AccountCode != null) ? item.AccountCode : "";
       
        var tmp = [invoiceID,description,quantity,unitAmount,taxType,taxAmount,lineAmount,accountCode];
        errorLISheet.appendRow(tmp);
      }
      // Add data in the Spreadsheet    
      var arr =[accType, contactName, date, dueDate, status, lineAmountTypes, subTotal, totalTax, total, updatedDateUTC, currencyCode, invoiceID, invoiceNumber,amountDue, amountPaid, AmountCredited ];
      errorSheet.appendRow(arr);     
    }     
  }
  if (failedInvoices > 0)
    Browser.msgBox(failedInvoices + " invoices could not be saved to Xero due to a validation error. See \"Error Invoices Upload\" & \"Error Invoices Upload Line Items\" sheets.");
}

function uploadPayments(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Payments Upload");
  
  if (!sheet) {
    Browser.msgBox('Error: "Payments Upload" sheet not found.');
    return false;
  }
  else if (sheet.getLastRow() == 0) {
    Browser.msgBox('Error: "Payments Upload" sheet has no data.');
    var headers = ['Invoice ID', 'Invoice Number', 'Account ID', 'Account Code', 'Date', 'Currency Rate', 'Amount', 'Reference'];
    sheet.getRange(1, 1, 1, 8).setValues(headers).setFontWeight('bold');
    return false;
  }
  var lrow = sheet.getDataRange().getLastRow();
  var lcol = 8;
  
  var data = sheet.getRange(3, 1, lrow-2, lcol).getValues();
  
  var payload  = "<Payments>";
  for(var i=0; i < data.length; i++){
    var invoiceID = data[i][0];
    var invNumber = data[i][1];
    var accID     = data[i][2];
    var accCode   = data[i][3];
    var date      = data[i][4];
    var cur_rate  = data[i][5];
    var amount    = data[i][6];
    var reference = data[i][7];
    payload += "<Payment>";
    
    if (invoiceID != null && invoiceID != undefined && invoiceID != "" ) 
      payload += "<Invoice><InvoiceID>" + invoiceID + "</InvoiceID></Invoice>";
    else if (invNumber != null && invNumber != undefined && invNumber != "") {      
      payload += "<Invoice><InvoiceNumber>" + invNumber + "</InvoiceNumber></Invoice>";
    } else {
      Browser.msgBox("Either InvoiceID or Invoice Number should be present. Please see Row:" + i+3);
      return;
    }

    if (accID != null && accID != undefined && accID != "") 
      payload += "<Account><AccountID>" + accID + "</AccountID></Account>";
    else if (accCode != null && accCode != undefined && accCode != "")
      payload += "<Account><Code>" + accCode + "</Code></Account>";
    else {
      Browser.msgBox("Either AccountID or Account Code should be present. Please see Row:" + i+3);
      return;
    }
    
    payload += "<Date>" + Utilities.formatDate(date, "GMT", "yyyy-MM-dd") + "</Date>";
    payload += "<CurrencyRate>" + cur_rate + "</CurrencyRate>";
    payload += "<Amount>" + amount + "</Amount>";
    payload += "<Reference>" + reference + "</Reference>";
    payload += "</Payment>";
    
  }
  payload  += "</Payments>";
  
  try {    
    Xero.connect();
    var method = 'PUT';
    var response = XERO.uploadData('Payments', payload, method);
    if (response.code != 200) {
      // Log failed items if any                      
      if (response.code == 400)  {
        Browser.msgBox(response.Message + ' . See "Error Payments Upload" sheet for details. Response Code: ' + response.code);
        logFailedPayments(response);      
      }  
      else
        Browser.msgBox('Error occured. Response Code: ' + response.code);  
    }  
  }
  catch(e) {
    Browser.msgBox(e.message);
  }   
}

function logFailedPayments(response) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var errorPayments = ss.getSheetByName('Error Payments Upload');
  if (!errorPayments)
    ss.insertSheet('Error Payments Upload');
  if (errorPayments.getLastRow() == 0) {
    // insert column headers
    var colHeaders = ['Invoice ID', 'Invoice Number', 'Account ID', 'Account Code', 'Date', 'Currency Rate', 'Amount', 'Reference', 'Error Message'];
    var range = errorPayments.getRange(1, 1, 1, 9).setValues([colHeaders]).setFontWeight('bold');
  }    
  // read response elements  
  var elements = response.Elements;
  for (var i = 0; i < elements.length; i++) {
    var elem       = elements[i];
    var invoiceID  = elem.Invoice.InvoiceID ? elem.Invoice.InvoiceID : "";
    var invoiceNum = elem.Invoice.InvoiceNumber ? elem.Invoice.InvoiceNumber : "";
    var accID      = elem.Account.AccountID ? elem.Account.AccountID : "";
    var accCode    = elem.Account.Code ? elem.Account.Code : "";
    var date       = elem.Date ? eval("new " + elem.Date.substr(2, elem.Date.length - 2) + ".toISOString()") : "";
    var curRate    = elem.Invoice.CurrencyRate ? elem.Invoice.CurrencyRate : "";
    var amount     = elem.Amount ? elem.Amount : "";    
    var ref        = elem.Invoice.Reference ? elem.Invoice.Reference : "";
    var errors     = elem.ValidationErrors;
    var errorMsg   = "";
    for (var j = 0; j < errors.length; j++) 
      errorMsg += errors[j].Message ? errors[j].Message + ". " : "";
    
    var row = [invoiceID, invoiceNum, accID, accCode, date, curRate, amount, ref, errorMsg];
    errorPayments.appendRow(row);
  }
}

// Utilities
function generateRandomString(n) {    
  var chars = ['a', 'b','c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9'];
  chars.push('A', 'B', 'C', 'D', 'E', 'F','G','H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z');
  var randomString = '';  
  for (i=0; i < n; i++) {
    r  = Math.random();
    r = r * 61; 
    r = Math.round(r);  
    randomString = randomString + chars[r];
  }  
  return randomString;
}  

function convertToUTC(input) {    
  var inter = new Date(Date.parse(input));
  var y = inter.getFullYear();
  var m = inter.getMonth() + 1;  
  var d = inter.getDate();
  mString = m; dString = d; yString = y;
  if (m < 10)
    var mString  = '0' + m;    
  if (d < 10)
    var dString = '0' + d;
  
  var output = yString + '-' + mString + '-' + dString + 'T00:00:00';
  Browser.msgBox('convertToUTC: ' + output);
  return output;
}

/*
function download_line_items_for_invoices(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Xero Invoices Download");
  var osheet = ss.getSheetByName("Download - Invoices Line Items");
  
  var line_items_status_col = 17;
  var lrow  = sheet.getDataRange().getLastRow();
  var data = sheet.getRange(3, line_items_status_col, lrow-2, 1).getValues();
  
  var invoices = [];
  for(var i=0; i < data.length; i++){
    if(data[i][0] != undefined && data[i][0] != null && data[i][0] != "YES"){
      var invoice_id = sheet.getRange(i+3, 12).getValue();
      
      var responses = XERO.download_invoice_for_an_id(invoice_id);
      var x = 1;
      
      var invoice = responses.Response.Invoices.Invoice;
      var line_items = invoice.LineItems.LineItem;
      
      if(line_items[0] != undefined && line_items[0] != null && line_items[0] != ""){
      
        for(var z=0; z < line_items.length; z++){
          var line_item = line_items[z];
          
          var description = (line_item.Description != null) ? line_item.Description : "";
          var quantity    = (line_item.Quantity != null) ? line_item.Quantity : "";
          var unitAmount  = (line_item.UnitAmount != null) ? line_item.UnitAmount : "";
          var taxType     = (line_item.TaxType != null) ? line_item.TaxType : "";
          var taxAmount   = (line_item.TaxAmount != null) ? line_item.TaxAmount : "";
          var lineAmount  = (line_item.LineAmount != null) ? line_item.LineAmount : "";
          var accountCode = (line_item.AccountCode != null) ? line_item.AccountCode : "";
          
          var row = [invoice_id, description, quantity, unitAmount, taxType, taxAmount, lineAmount, accountCode];
          osheet.appendRow(row).activate();
        }
        
      }else{
          var description = (line_items.Description != null) ? line_items.Description : "";
          var quantity    = (line_items.Quantity != null) ? line_items.Quantity : "";
          var unitAmount  = (line_items.UnitAmount != null) ? line_items.UnitAmount : "";
          var taxType     = (line_items.TaxType != null) ? line_items.TaxType : "";
          var taxAmount   = (line_items.TaxAmount != null) ? line_items.TaxAmount : "";
          var lineAmount  = (line_items.LineAmount != null) ? line_items.LineAmount : "";
          var accountCode = (line_items.AccountCode != null) ? line_items.AccountCode : "";
          
          var row = [invoice_id, description, quantity, unitAmount, taxType, taxAmount, lineAmount, accountCode];
          osheet.appendRow(row).activate();
      }
      
      sheet.getRange(i+3, line_items_status_col).setValue("YES");
    }
  }
  
}
*/
