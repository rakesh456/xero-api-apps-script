var baseURL = 'https://api.xero.com';
var requestTokenURL = baseURL + '/oauth/RequestToken';
var authorizeURL = baseURL + '/oauth/Authorize';
var accessTokenURL = baseURL + '/oauth/AccessToken';
var apiEndPoint = baseURL + '/api.xro/2.0';

var Xero = {  
  appType: '', userAgent: '', consumerKey: '', consumerSecret: '', callbackURL: '', rsaKey: '', isConnected: false, 
  
  getProperty: function(propName) {
    if (this[propName] != null)
      return this[propName];
    else
      return false;
  },

  getSettings: function() {
    var p = PropertiesService.getScriptProperties().getProperties();        
    if (p.appType == null || p.appType == '') {
      Browser.msgBox('Please enter Xero Settings.');
      return false;
    }  
    else if (p.userAgent == null || p.userAgent == '' || p.consumerKey == null || p.consumerKey == '' || p.consumerSecret == null || p.consumerSecret == '') {
        Browser.msgBox('Error: Missing Xero Settings (Apllication Name/ Consumer Key/ Consumer Secret).');
        return false;      
    }
    else if (p.appType == 'Public') {
      if (p.callbackURL == null || p.callbackURL == '') {
        Browser.msgBox('Error: Missing Xero Settings (Callback URL.');
        return false;
      } 
    }
    else if (p.appType == 'Partner') {
        if (p.callbackURL == null || p.callbackURL == '' || p.rsaKey == null || p.rsaKey == '' ) {
        Browser.msgBox('Error: Missing Xero Settings (Callback URL/ RSA Key');
        return false;
      }           
    }
       
    this.appType = p.appType;
    this.userAgent = p.userAgent;
    this.consumerKey = p.consumerKey;
    this.consumerSecret = p.consumerSecret;
    this.rsaKey = p.rsaKey;
    this.callbackURL = p.callbackURL;
    this.requestTokenSecret = "";
    this.accessToken = "";
    this.accessTokenSecret = "";    
    if (p.requestTokenSecret != null) 
      this.requestTokenSecret = p.requestTokenSecret;   
    if (p.accessToken != null) 
      this.accessToken = p.accessToken;    
    if (p.accessTokenSecret != null) 
      this.accessTokenSecret = p.accessTokenSecret;        
    if (p.isConnected != null)
      this.isConnected = p.isConnected;
    return true;
  },  
  
  connect: function() {
    this.getSettings();
    if (this.appType != 'Private' /*&& !this.isConnected*/) {
      // Ask user to connect to Xero first
      // Step 1. Get an Unauthorised Request Token    
      var payload = {"oauth_consumer_key": this.consumerKey,
                     "oauth_signature_method": "PLAINTEXT",
                     "oauth_signature": encodeURIComponent(this.consumerSecret + '&'),
                     "oauth_timestamp": ((new Date().getTime())/1000).toFixed(0),
                     "oauth_nonce": generateRandomString(Math.floor(Math.round(25))),
                     "oauth_version": "1.0",
                     "oauth_callback": this.callbackURL};  
      var options = {"method": "post", "payload": payload};
      var response = UrlFetchApp.fetch(requestTokenURL, options);  
      var reoAuthToken = /(oauth_token=)([a-zA-Z0-9]+)/;    
      var tokenMatch = reoAuthToken.exec(response.getContentText());
      var oAuthRequestToken = tokenMatch[2];
      var reTokenSecret = /(oauth_token_secret=)([a-zA-Z0-9]+)/;
      var secretMatch = reTokenSecret.exec(response.getContentText())  ;
      var tokenSecret = secretMatch[2];
      PropertiesService.getScriptProperties().setProperty('requestTokenSecret', tokenSecret);      
      //Logger.log('Request Token = ' + oAuthRequestToken);
      //Logger.log('Request Token Secret = ' + tokenSecret);
      
      return authorizeURL + '?oauth_token=' + oAuthRequestToken;
      
      // Step 2 Show user the link to connect to Xero       
      /*
      var ss = SpreadsheetApp.getActiveSpreadsheet();      
      var app = UiApp.createApplication(); 
      app.setHeight(50);app.setWidth(100);
      var link = app.createAnchor('Connect to Xero', true, authorizeURL + '?oauth_token=' + oAuthRequestToken);
      var handler = app.createServerHandler('closeApp');
      link.addClickHandler(handler);
      app.add(link);
      ss.show(app);      
      */
    }
  },
  
  fetchPrivateAppData: function(item, pageNo) {    
    /* FETCH PRIVATE APPLICATION DATA */
    var method = 'GET';
    var requestURL = apiEndPoint + '/' + item ;    
    var oauth_signature_method = 'RSA-SHA1';
    var oauth_timestamp = (new Date().getTime()/1000).toFixed();
    var oauth_nonce = generateRandomString(Math.floor(Math.random() * 50));
    var oauth_version = '1.0';     
    var signBase = 'GET' + '&' + encodeURIComponent(requestURL) + '&'
    + encodeURIComponent('oauth_consumer_key=' + this.consumerKey + '&oauth_nonce=' + oauth_nonce + '&oauth_signature_method='
                         + oauth_signature_method + '&oauth_timestamp=' + oauth_timestamp + '&oauth_token=' + this.consumerKey + '&oauth_version='
                         + oauth_version + '&page=' + pageNo);  
    if (!this.rsa) {
      this.rsa = new RSAKey();      
      this.rsa.readPrivateKeyFromPEMString(this.rsaKey);        
      var sbSigned = this.rsa.signString(signBase, 'sha1');              
    }
    else
      var sbSigned = this.rsa.signString(signBase, 'sha1');              
    
    var data = new Array();
    for (var i =0; i < sbSigned.length; i += 2) 
      data.push(parseInt("0x" + sbSigned.substr(i, 2)));      
    var oauth_signature = hex2b64(sbSigned);  
    
    var authHeader = "OAuth oauth_token=\"" + this.consumerKey + "\",oauth_nonce=\"" + oauth_nonce + "\",oauth_consumer_key=\"" + this.consumerKey 
    + "\",oauth_signature_method=\"RSA-SHA1\",oauth_timestamp=\"" + oauth_timestamp + "\",oauth_version=\"1.0\",oauth_signature=\""
    + encodeURIComponent(oauth_signature) + "\"";    
    
    var headers = { "User-Agent": this.userAgent, "Authorization": authHeader, "Accept": "application/json"};    
    var options = { muteHttpExceptions: true, "headers": headers}; 
    var response = UrlFetchApp.fetch(requestURL + '?page=' + pageNo, options);
    Logger.log(response);
    if (response.getResponseCode() == 200)
      return JSON.parse(response.getContentText());    
    else
      return false;
  },
  
    fetchData: function(item, pageNo) {
    this.getSettings();
    // lets find out which method to use for fetching data
    switch(this.appType) {
      case 'Private': 
        return this.fetchPrivateAppData(item, pageNo);
        break;
      case 'Public':
        if (!this.isConnected)
          this.connect();
        else
          return this.fetchPublicAppData(item, pageNo);       
        break;
      case 'Partner':
        if (!this.isConnected)
          this.connect();
        else
        return this.fetchPartnerAppData(item);
        break;
    }       
  },
  
  
  fetchPublicAppData: function(item, pageNo) {    
    /* For PUBLIC APPLICATION TYPE */
    this.getSettings(); // get latest settings
    var method = 'GET';
    var requestURL = apiEndPoint + '/' + item ;    
    var oauth_signature_method = 'HMAC-SHA1';
    var oauth_timestamp =  (new Date().getTime()/1000).toFixed();
    var oauth_nonce = generateRandomString(Math.floor(Math.random() * 50));
    var oauth_version = '1.0';  
    //oauth_consumer_key=SRPZNHSGTI5L1WAQJGBTZROYVH3IZ3&oauth_nonce=GuWaMcBr3Bq&oauth_signature_method=HMAC-SHA1&oauth_timestamp=1404572819
    //&oauth_token=NVOXINVPTRTJHZOGMBHEDRUXJ6MPIN&oauth_version=1.0&page=1
    var signBase = 'GET' + '&' + encodeURIComponent(requestURL) + '&' 
    + encodeURIComponent('oauth_consumer_key=' + this.consumerKey + '&oauth_nonce=' + oauth_nonce + '&oauth_signature_method=' + oauth_signature_method 
                         + '&oauth_timestamp=' + oauth_timestamp + '&oauth_token=' + this.accessToken  + '&oauth_version=' + oauth_version + '&page=' + pageNo);  
    Logger.log(signBase);
    var sbSigned = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_1, signBase, 
                                                  encodeURIComponent(this.consumerSecret) + '&' + encodeURIComponent(this.accessTokenSecret));        
    Logger.log('sbSigned: ' + sbSigned);    
    //var oauth_signature = hex2b64(sbSigned);  
    var oauth_signature = Utilities.base64Encode(sbSigned);
    Logger.log(oauth_signature);
    
    var authHeader = "OAuth oauth_consumer_key=\"" + this.consumerKey + "\",oauth_nonce=\"" + oauth_nonce + "\",oauth_token=\"" + this.accessToken 
    + "\",oauth_signature_method=\"" + oauth_signature_method +"\",oauth_timestamp=\"" + oauth_timestamp + "\",oauth_version=\"1.0\",oauth_signature=\""
    + encodeURIComponent(oauth_signature) + "\"";
    
    var headers = { "User-Agent": + this.userAgent, "Authorization": authHeader, "Accept": "application/json" };
    var options = { "headers": headers};
    try {
      var response = UrlFetchApp.fetch(requestURL + '?page='+ pageNo, options);
      return JSON.parse(response.getContentText());  
    }
    catch(e) {
      //Logger.log(response.getContentText());
      Logger.log(e.message);
      Browser.msgBox(e.message);
    } 
  },
  fetchPartnerAppData: function(item) {
    return false;
  },
  
  uploadData: function(item, xml, method) {
    var method = method || 'POST';
    var requestURL = apiEndPoint + '/' + item ;    
    var oauth_signature_method = 'RSA-SHA1';
    var oauth_timestamp = (new Date().getTime()/1000).toFixed();
    var oauth_nonce = generateRandomString(Math.floor(Math.random() * 50));
    var oauth_version = '1.0';
    var signBase = method + '&' + encodeURIComponent(requestURL) + '&' + 'SummarizeErrors=false' + 
    + encodeURIComponent('oauth_consumer_key=' + this.consumerKey + '&oauth_nonce=' + oauth_nonce + '&oauth_signature_method='
                         + oauth_signature_method + '&oauth_timestamp=' + oauth_timestamp + '&oauth_token=' + this.consumerKey + '&oauth_version='
                         + oauth_version + '&order=');  
    if (method == 'POST')
      signBase += '&xml=' + xml;
    
    
    var rsa = new RSAKey();
    rsa.readPrivateKeyFromPEMString(this.rsaKey);
    var sbSigned = rsa.signString(signBase, 'sha1');
    Logger.log(sbSigned);
    
    var data = new Array();
    for (var i =0; i < sbSigned.length; i += 2) 
      data.push(parseInt("0x" + sbSigned.substr(i, 2)));      
    var oauth_signature = hex2b64(sbSigned);  
    
    var authHeader = "OAuth oauth_token=\"" + this.consumerKey + "\",oauth_nonce=\"" + oauth_nonce + "\",oauth_consumer_key=\"" + this.consumerKey 
    + "\",oauth_signature_method=\"RSA-SHA1\",oauth_timestamp=\"" + oauth_timestamp + "\",oauth_version=\"1.0\",oauth_signature=\""
    + encodeURIComponent(oauth_signature) + "\"";    
    var payload = {"order": "", "xml": xml};
    var headers = { "User-Agent": this.userAgent, "Authorization": authHeader, "Accept": "application/json", "muteHttpExceptions": true };
    var options = { "headers": headers, "method": "post", "payload": payload };  
    try {
      var response = UrlFetchApp.fetch(requestURL + '?SummarizeErrors=false', options); 
      if (response.getResponseCode() == 200 && response.getHeaders().Status == "OK") 
        return JSON.parse(response.getContentText());                
      else
        throw "Request Failed: Response Code: " + response.getResponseCode();      
    } 
    catch(e) {
      throw e.message;
    }
    Logger.log(response.getContentText());    // return XML
    return false;
  }  
}

