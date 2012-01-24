var XMLHTTPSUCCESS = 200;
var XMLHTTPREADY = 4;

// Public Method
function CrmService(sOrg, sServer) {
/// CrmService object constructor. allows for calling crm webservice CRUD operations directly from client script.
/// CRM Organization Name. Not required if used on a CRM form.
/// (optional) URL of CRM Server. Url of current window is used if this is null.
this.org = sOrg;
this.server = sServer;

if (sOrg == null) {
if (typeof (ORG_UNIQUE_NAME) != “undefined”) {
this.org = ORG_UNIQUE_NAME;
}
else {
alert(“Error: Org Name must be defined.”);
}
}

// URL was not provided, assume the JS file is running from the CRM Server
if (sServer == null) {
this.server = window.location.protocol + “//” + window.location.host;
}
}

CrmService.prototype.CreateXmlHttp = function() {
var oXmlHttp = null;

if (window.XMLHttpRequest) {
oXmlHttp = new XMLHttpRequest();
}
else {
var arrProgIds = ["Msxml2.XMLHTTP", "Microsoft.XMLHTTP", "MSXML2.XMLHTTP.3.0"];
for (var iCount = 0; iCount < arrProgIds.length; iCount++) {
try {
oXmlHttp = new ActiveXObject(arrProgIds[iCount]);
break;
}
catch (e) { }
}
}

if (oXmlHttp == null) {
alert(“Error: Failed to create XmlHTTP object.”);
}

return oXmlHttp;
}

CrmService.prototype.Associate = function(sEntity1, sId1, sEntity2, sId2, sRelationship, fCallback) {
/// Associate two records that have a m:m relationship
/// Entity name of first record being associated.
/// Id of first record being associated.
/// Entity name of second record being associated.
/// Id of second record being associated.
/// Name of the many-to-many relationship.
/// (Optional) callback function. If this is null then the update will be synchronous.

var sXml = this._GetHeader();
sXml += “”;
sXml += “” + sId1 + “”;
sXml += “” + sEntity1 + “”;
sXml += “” + sId2 + “”;
sXml += “” + sEntity2 + “”;
sXml += “” + sRelationship + “”;
sXml += “”;

return this._ExecuteRequest(sXml, “Execute”, this._GenericCallback, fCallback);
}

CrmService.prototype.Disassociate = function(sEntity1, sId1, sEntity2, sId2, sRelationship, fCallback) {
/// Disassociate two records that have a m:m relationship
/// Entity name of first record being disassociated.
/// Id of first record being disassociated.
/// Entity name of second record being disassociated.
/// Id of second record being disassociated.
/// Name of the many-to-many relationship.
/// (Optional) callback function. If this is null then the update will be synchronous.

var sXml = this._GetHeader();
sXml += “”;
sXml += “” + sId1 + “”;
sXml += “” + sEntity1 + “”;
sXml += “” + sId2 + “”;
sXml += “” + sEntity2 + “”;
sXml += “” + sRelationship + “”;
sXml += “”;

return this._ExecuteRequest(sXml, “Execute”, this._GenericCallback, fCallback);
}

CrmService.prototype.SetState = function(sEntityName, sId, sState, iStatus, fCallback) {
/// Sets the State and Status of a record
/// Entity name of record.
/// Id of record being updated.
/// State (string). eg, “Active”
/// Status (integer). Use -1 for default status.
/// (Optional) callback function. If this is null then the update will be synchronous.

var xml = this._GetHeader();
xml += “”;
xml += “” + sId + “”;
xml += “” + sEntityName + “”;
xml += “” + sState + “”;
xml += “” + iStatus + “”;
xml += “”;

return this._ExecuteRequest(xml, “Execute”, this._GenericCallback, fCallback);
}

CrmService.prototype.Update = function(oBusinessEntity, fCallback) {
/// Update a business entity
/// A BusinessEntity object containing the updated attributes
/// (Optional) callback function. If this is null then the update will be synchronous.

//build soap sXml message
var sXml = this._GetHeader();
sXml += “”;
sXml += this._GetAttributesXml(oBusinessEntity);
sXml += “”;

return this._ExecuteRequest(sXml, “Update”, this._GenericCallback, fCallback);
}

CrmService.prototype.Retrieve = function(sEntityName, sEntityId, asColumnSet, fCallback) {
/// Retrieve a business entity
/// Entity name. e.g., account, contact, etc.
/// Guid of record to retrieve.
/// Array of strings indicating the column set to retrieve.
/// (Optional) Async callback function. If this is null then the retrieve will be synchronous.
/// Returns a BusinessEntity object (if being called synchronously)

//build soap sXml message
var sXml = this._GetHeader();

sXml += “” + sEntityName + “”;
sXml += “” + sEntityId + “”;

if (asColumnSet != null && asColumnSet.length > 0) {
sXml += “”;
for (var i = 0; i < asColumnSet.length; i++) {
sXml += “” + asColumnSet[i] + “”;
}
sXml += “”;
}
sXml += “”;

return this._ExecuteRequest(sXml, “Retrieve”, this._RetrieveCallback, fCallback);
}

CrmService.prototype._RetrieveCallback = function(oXmlHttp, callback) {
///(private) Retrieve message callback.

//oXmlHttp must be completed
if (oXmlHttp.readyState != XMLHTTPREADY) {
return;
}

//check for server errors
if (this._HandleErrors(oXmlHttp)) {
return;
}

//parse sXml into a oBusinessEntity object
var oBE = new BusinessEntity();
var oNodes = oXmlHttp.responseXML.selectSingleNode(“//RetrieveResult”).childNodes;
for (var i = 0; i < oNodes.length; i++) {
var oNode = oNodes[i];
var oObj = new Object();
oObj["value"] = oNode.text;

for (var j = 0; j < oNode.attributes.length; j++) {
oObj[oNode.attributes[j].nodeName] = oNode.attributes[j].nodeValue;
}

oBE.attributes[oNode.baseName] = oObj;
}

//return entity if sync, or call user callback func if async
if (callback != null) {
callback(oBE);
}
else {
return oBE;
}
}

CrmService.prototype.Create = function(oBusinessEntity, fCallback) {
/// Create a business entity
/// A BusinessEntity object containing the record to create.
/// (Optional) Async callback function. If this is null then the create will be synchronous.

//build soap sXml message
var sXml = this._GetHeader();
sXml += “”;
sXml += “”;
sXml += this._GetAttributesXml(oBusinessEntity);
sXml += “”;

var crmServiceObject = this;
return this._ExecuteRequest(sXml, “Create”, this._CreateCallback, fCallback);
}

CrmService.prototype.Delete = function(sEntityName, sEntityId, fCallback) {
/// Delete a business entity
/// Entity name. e.g., account, contact, etc.
/// Guid of record to delete.
/// (Optional) Async callback function. If this is null then the delete will be synchronous.

//build soap sXml message
var sXml = this._GetHeader();
sXml += “”;
sXml += “” + sEntityName + “”;
sXml += “” + sEntityId + “”;
sXml += “”;

return this._ExecuteRequest(sXml, “Delete”, this._GenericCallback, fCallback);
}

CrmService.prototype.Fetch = function(sFetchXml, fCallback) {
/// Execute a FetchXml request. (result is an array of BusinessEntity objects)
/// fetchxml string
/// (Optional) Async callback function. If this is null then the fetch will be synchronous.

var sXml = this._GetHeader();
sXml += “”;
sXml += this._crmXmlEncode(sFetchXml);
sXml += “”;

return this._ExecuteRequest(sXml, “Fetch”, this._FetchCallback, fCallback);
}

CrmService.prototype._FetchCallback = function(oXmlHttp, callback) {
///(private) Fetch message callback.
//oXmlHttp must be completed
if (oXmlHttp.readyState != XMLHTTPREADY) {
return;
}

//check for server errors
if (this._HandleErrors(oXmlHttp)) {
return;
}

//Get the id of the created object from the returned xml
var sFetchResult = oXmlHttp.responseXML.selectSingleNode(“//FetchResult”).text;

var oResultDoc = new ActiveXObject(“Microsoft.XMLDOM”);
oResultDoc.async = false;
oResultDoc.loadXML(sFetchResult);

//parse result sXml into array of BusinessEntity objects
var oResults = new Array(oResultDoc.firstChild.childNodes.length);

var iLen = oResultDoc.firstChild.childNodes.length;
for (var i = 0; i < iLen; i++) {

var oResultNode = oResultDoc.firstChild.childNodes[i];
var oBE = new BusinessEntity();

var iLenInner = oResultNode.childNodes.length;
for (var j = 0; j < iLenInner; j++) {
var oNode = oResultNode.childNodes[j];
var oObj = new Object();
oObj["value"] = oNode.text;

for (var k = 0; k < oNode.attributes.length; k++) {
oObj[oNode.attributes[k].nodeName] = oNode.attributes[k].nodeValue;
}

oBE.attributes[oNode.baseName] = oObj;
}

oResults[i] = oBE;
}

//return entity id if sync, or call user callback func if async
if (callback != null) {
callback(oResults);
}
else {
return oResults;
}
}

CrmService.prototype._ExecuteRequest = function(sXml, sMessage, fInternalCallback, fUserCallback) {
// Create the XMLHTTP object for the Update method.
var oXmlHttp = this.CreateXmlHttp();
oXmlHttp.open(“POST”, this.server + “/mscrmservices/2007/crmservice.asmx”, (fUserCallback != null));
oXmlHttp.setRequestHeader(“Content-Type”, “text/xml; charset=utf-8?);
oXmlHttp.setRequestHeader(“SOAPAction”, “http://schemas.microsoft.com/crm/2007/WebServices/” + sMessage);

if (fUserCallback != null) {
//asynchronous
var crmServiceObject = this;
oXmlHttp.onreadystatechange = function() { fInternalCallback.call(crmServiceObject, oXmlHttp, fUserCallback) };
oXmlHttp.send(sXml);
}
else {
//synchronous
oXmlHttp.send(sXml);
return fInternalCallback.call(this, oXmlHttp, null);
}
}

CrmService.prototype._GetAttributesXml = function(oBusinessEntity) {
var sXml = “”;

for (var sName in oBusinessEntity.attributes) {
var oAttrib = oBusinessEntity.attributes[sName];

if (typeof (oAttrib.value) == “undefined”) {
sXml += “”;
sXml += this._crmXmlEncode(oAttrib);
sXml += “”;
}
else {
sXml += “”;
sXml += oAttrib.value;
sXml += “”;
}
}

return sXml;
}

CrmService.prototype._HandleErrors = function(oXmlHttp) {
/// (private) Handles oXmlHttp errors
if (oXmlHttp.status != XMLHTTPSUCCESS) {
var sError = “CrmService Error:\n\n” + oXmlHttp.responseXML.text;
sError += “\n\nXML:\n” + oXmlHttp.responseXML.xml;
sError += “\n\nHTTP: ” + oXmlHttp.status + ” – ” + oXmlHttp.statusText;
alert(sError);
return true;
}
else {
return false;
}
}

CrmService.prototype._GetHeader = function() {
// Use CRM’s GenerateAuthenticationHeader function if it is available.
if (typeof (GenerateAuthenticationHeader) != “undefined”) {
return “” + GenerateAuthenticationHeader();
}
else {
return “0? + this._crmXmlEncode(this.org) + “00000000-0000-0000-0000-000000000000?;
}
}

CrmService.prototype._GenericCallback = function(oXmlHttp, callback) {
///(private) Generic callback. (used for update and delete messages)
if (oXmlHttp.readyState == XMLHTTPREADY) {
if (!this._HandleErrors(oXmlHttp)) {
if (callback != null) {
callback();
}
}
}
}

CrmService.prototype._CreateCallback = function(oXmlHttp, callback) {
///(private) Create message callback.

//oXmlHttp must be completed
if (oXmlHttp.readyState != XMLHTTPREADY) {
return;
}

//check for server errors
if (this._HandleErrors(oXmlHttp)) {
return;
}

//Get the id of the created object from the returned xml
var oResult = oXmlHttp.responseXML.selectSingleNode(“//CreateResult”);

//return entity id if sync, or call user callback func if async
if (callback != null) {
callback(oResult.text);
}
else {
return oResult.text;
}
}

CrmService.prototype._crmXmlEncode = function(s, charToEncode) {
if (s == null) { return s; } else { if (typeof (s) != “string”) { s = s.toString(); } }
if (typeof (charToEncode) != “undefined” && charToEncode != null) {
if (charToEncode.length > 1) charToEncode = charToEncode.charAt(0);
var sEncodedChar = this._XmlEncode(charToEncode);
var rex = new RegExp(charToEncode, “g”);
return s.replace(rex, sEncodedChar);
}
return this._surrogateAmpersandWorkaround(s, this._XmlEncode);
}

CrmService.prototype._surrogateAmpersandWorkaround = function(s, encodingFunction) {
s = s.replace(/([\uD800-\uDBFF][\uDC00-\uDFFF])/g, function($1) { return “CRMEntityReferenceOpen” + ((($1.charCodeAt(0) – 0xD800) * 0×400) + ($1.charCodeAt(1) & 0x03FF) + 0×10000).toString(16) + “CRMEntityReferenceClose”; });
s = s.replace(/[\uD800-\uDFFF]/g, ‘\uFFFD’);
s = encodingFunction(s);
s = s.replace(/CRMEntityReferenceOpen/g, “&#x”);
s = s.replace(/CRMEntityReferenceClose/g, “;”);
return s;
}

CrmService.prototype._XmlEncode = function(strInput) {
var c;
var HtmlEncode = “”;

if (strInput == null) {
return null;
}
if (strInput == “”) {
return “”;
}

var iLen = strInput.length;
for (var i = 0; i 96) && (c 64) && (c 47) && (c < 58)) ||
(c == 46) ||
(c == 44) ||
(c == 45) ||
(c == 95)) {
HtmlEncode = HtmlEncode + String.fromCharCode(c);
}
else {
HtmlEncode = HtmlEncode + ‘&#’ + c + ‘;’;
}
}

return HtmlEncode;
}

//Object
// BusinessEntity object
function BusinessEntity(sName) {
/// Business Entity object constructor. for use with the CrmService object.
/// Entity Name. e.g., ‘account’
this.name = sName;
this.attributes = new Object();
}

BusinessEntity.prototype.name = “”;
BusinessEntity.prototype.attributes = new Object();

// CrmLookup object
function CrmLookup(sType, sValue) {
/// Lookup attribute constructor. for use with the CrmService object.
/// logical name of the lookup entity; e.g., ‘account’.
/// Guid of the lookup record.
this.type = sType;
this.value = sValue;
}

function Utility()
{ }

/// Get Element by Class Name
/// return type — collection of DOM objects
Utility.prototype.getElementsByClassName = function(oElm, strTagName, oClassNames) {
var arrElements = (strTagName == “*” && oElm.all) ? oElm.all : oElm.getElementsByTagName(strTagName);
var arrReturnElements = new Array();
var arrRegExpClassNames = new Array();
if (typeof oClassNames == “object”) {
for (var i = 0; i < oClassNames.length; i++) {
arrRegExpClassNames.push(new RegExp(“(^|\\s)” + oClassNames[i].replace(/\-/g, “\\-”) + “(\\s|$)”));
}
}
else {
arrRegExpClassNames.push(new RegExp(“(^|\\s)” + oClassNames.replace(/\-/g, “\\-”) + “(\\s|$)”));
}
var oElement;
var bMatchesAll;
for (var j = 0; j < arrElements.length; j++) {
oElement = arrElements[j];
bMatchesAll = true;
for (var k = 0; k < arrRegExpClassNames.length; k++) {
if (!arrRegExpClassNames[k].test(oElement.className)) {
bMatchesAll = false;
break;
}
}
if (bMatchesAll) {
arrReturnElements.push(oElement);
}
}
return (arrReturnElements)
}

// Add event
// obj — DOM object (e.g) crmForm.all.new_vpnlinkid
// type — DOM Event (e.g) click, mouseover etc, note: event name without ‘on’ keyword.
// fn — Function name which you want to add..
Utility.prototype.addEvent = function(obj, type, fn) {
try {
if (obj.attachEvent) {
obj.attachEvent(‘on’ + type, fn);
} else
obj.addEventListener(type, fn, false);
} catch (e) {
alert(e);
}
}

// removeevent
Utility.prototype.removeEvent = function(obj, type, fn) {
if (obj.detachEvent) {
obj.detachEvent(‘on’ + type, fn);
} else
obj.removeEventListener(type, fn, false);
}

var CRM_BOOLEAN = 0;
var CRM_LOOKUP = 1;
var CRM_PICKLIST = 2;
var CRM_STRING = 3;
var CRM_DATETIME = 4;
var CRM_DECIMAL = 5;
var CRM_FLOAT = 6;
var CRM_INTEGER = 7;
var CRM_MEMO = 8;
var CRM_MONEY = 9;
var CRM_PARTYLIST = 10;
var CRM_REGARDING = 11;
var CRM_STATE = 12;

//Utility
//Use to get value from crm entity attributes
//GetValue(new_test1);
//
Utility.prototype.GetValue = function(attribute, type) {
var returnValue = “”;
switch (type) {
case CRM_LOOKUP:
var _lookupItem = new Array;

// This gets the lookup for the attribute.
_lookupItem = attribute.DataValue;

if (_lookupItem != null) {
// If there is data in the field, return it.
if (_lookupItem[0] != null) {
//return text value of the lookup.
returnValue = _lookupItem[0].name;
}
}
break;
default:
var _attribute = crmForm.all.attribute;
returnValue = _attribute.DataValue;
}
return returnValue;
}

Utility.prototype.SetValue = function(attribute, value) {
attribute.DataValue = value;
}

Utility.prototype.AttributeVisible = function(attribute, value) {
var _attribute = attribute;
var _attributeLabel = eval(attribute.id + ‘_c’);
var _attribute_Column = eval(attribute.id + ‘_d’);
if (value != true) {
_attribute.style.display = ‘none’;
_attributeLabel.style.display = ‘none’;
_attribute_Column.style.display = ‘none’;
} else {
_attribute.style.display = ”;
_attributeLabel.style.display = ”;
_attribute_Column.style.display = ”;
}
}

Utility.prototype.SectionVisible = function(section, value) {
var _section = section;
_section = eval(_section.id + ‘_c’);
if (value != true) {
_section.parentElement.parentElement.style.display = ‘none’;
} else {
_section.parentElement.parentElement.style.display = ”;
}
}

Utility.prototype.TabVisible = function(index, value) {
var _tab = eval(‘crmForm.all.tab’ + index + ‘Tab’);

if (value != true) {
_tab.style.visibility = ‘hidden’
} else {
_tab.style.visibility = ”;
}
}

/// Access external resources in crm
/// Page url
/// param Name
/// param Values
Utility.prototype.AccessResource = function(resource, paramName, paramValue) {

var _name = paramName.split(‘,’);
var _values = paramValue.split(‘,’);
var _resource = “”;
var _query;

for (var i = 0; i < _name.length; i++) {
if (i != eval(_name.length – 1)) {
_resource += _name[i] + “=” + _values[i] + “&”;
} else {
_resource += _name[i] + “=” + _values[i];
}
}

_query = resource + ‘?’ + _resource;

var xHReq = new ActiveXObject(“Msxml2.XMLHTTP”);
xHReq.Open(“POST”, param, false);
xHReq.send(_query);
// Capture the result.
return xHReq.responseText;
}