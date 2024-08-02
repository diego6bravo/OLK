<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<%
          
sql = "select IsNull(T0.PrintHeadr,IsNull(T0.CompnyName, '')) CmpName, T1.CardName, IsNull(IsNull(T6.Tel1, T5.Phone1), '') Phone1, IsNull(T6.Name, '') CntctPrsn, " & _
"T2.Street ShipperStreet, T2.Block ShipperBlock, T2.City ShipperCity, T2.ZipCode ShipperZipCode, T2.County ShipperCounty, T2.State ShipperState, T2.Country ShipperCountry, " & _
"T3.Street ShipToStreet, T3.Block ShipToBlock, T3.City ShipToCity, T3.ZipCode ShipToZipCode, T3.County ShipToCounty, T3.State ShipToState, T3.Country ShipToCountry, " & _
"T4.Street ShipFromStreet, T4.Block ShipFromBlock, T4.City ShipFromCity, T4.ZipCode ShipFromZipCode, T4.County ShipFromCounty, T4.State ShipFromState, T4.Country ShipFromCountry " & _
"from OADM T0 inner join R3_ObsCommon..TDOC T1 on T1.LogNum = " & Session("RetVal") & " cross join ADM1 T2 " & _
"inner join CRD1 T3 on T3.CardCode = N'" & Session("UserName") & "' and AdresType = 'B' and T3.Address = N'" & Request("PayToCode") & "' " & _
"inner join OWHS T4 on T4.WhsCode = N'" & Request("WhsCode") & "' " & _
"inner join OCRD T5 on T5.CardCode = T3.CardCode " & _
"left outer join OCPR T6 on T6.CardCode = T5.CardCode and T6.CntctCode = T1.CntctCode "
set rd = Server.CreateObject("ADODB.RecordSet")
set rd = conn.execute(sql)

sql = "select FieldID, Value from OLKShipmentSettings where ShipTypeID = 0"
set rs = Server.CreateObject("ADODB.RecordSet")
rs.open sql, conn, 3, 1
     
Dim strXml
     
Dim dom
set dom = Server.CreateObject("MSXML2.DOMDocument.3.0")

dom.async = False
dom.resolveExternals = False
dom.preserveWhiteSpace = True

'	***** Start AccessRequest ******

set node = dom.createProcessingInstruction("xml", "version='1.0'")
dom.appendChild node
set node = Nothing

Dim root
set root = dom.createElement("AccessRequest")
root.setAttribute "xml:lang", "en-US"
dom.appendChild root

rs.Filter = "FieldID = 0"
Dim el
set el = dom.createElement("AccessLicenseNumber")
el.text = rs(1)
root.appendChild el

rs.Filter = "FieldID = 1"
set el = dom.createElement("UserId")
el.text = rs(1)
root.appendChild el

rs.Filter = "FieldID = 2"
set el = dom.createElement("Password")
el.text = rs(1)
root.appendChild el

set el = Nothing
set root = Nothing

strXml = dom.xml


'	***** Start RatingServiceSelectionRequest ******
set dom = Server.CreateObject("MSXML2.DOMDocument.3.0")

dom.async = False
dom.resolveExternals = False
dom.preserveWhiteSpace = True

set node = dom.createProcessingInstruction("xml", "version='1.0'")
dom.appendChild node
set node = Nothing

set root = dom.createElement("RatingServiceSelectionRequest")
root.setAttribute "xml:lang", "en-US"
dom.appendChild root

'	***** Start Request ******

Dim subNode
set subNode = dom.createElement("Request")
root.appendChild subNode

set el = dom.createElement("RequestAction")
el.text = "Rate"
subNode.appendChild el

set el = dom.createElement("RequestOption")
el.text = "Shop"
subNode.appendChild el

set subSubNode = dom.createElement("TransactionReference")
subNode.appendChild subSubNode

set el = dom.createElement("CustomerContext")
subSubNode.appendChild el

set el = dom.createElement("XpciVersion")
subSubNode.appendChild el

set subNode = dom.createElement("Pickup")
root.appendChild subNode

rs.Filter = "FieldID = 3"
set el = dom.createElement("Code")
el.text = rs(1)
subNode.appendChild el

'set subSubNode = dom.createElement("CustomerClassification") 'US Only
'subNode.appendChild subSubNode

'set el = dom.createElement("Code")
'el.text = "01"
'subSubNode.appendChild el

'	***** Start Shipment ******

set subNode = dom.createElement("Shipment")
root.appendChild subNode

set subSubNode = dom.createElement("Shipper")
subNode.appendChild subSubNode

set el = dom.createElement("Name")
el.text = rd("CmpName")
subSubNode.appendChild el

rs.Filter = "FieldID = 4"
set el = dom.createElement("ShipperNumber") 'Falta Settings
el.text = rs(1)
subSubNode.appendChild el

set subAdd = dom.createElement("Address")
subSubNode.appendChild subAdd

set el = dom.createElement("AddressLine1")
el.text = rd("ShipperStreet")
subAdd.appendChild el

set el = dom.createElement("AddressLine2")
el.text = rd("ShipperBlock")
subAdd.appendChild el

set el = dom.createElement("AddressLine3")
el.text = ""
subAdd.appendChild el

set el = dom.createElement("City") 'Required if no postal code
el.text = rd("ShipperCity")
subAdd.appendChild el

set el = dom.createElement("StateProvinceCode") 'Irland use 5 digit abbreviation for county
el.text = rd("ShipperState")
subAdd.appendChild el

set el = dom.createElement("PostalCode") 'Required use postal code
el.text = rd("ShipperZipCode")
subAdd.appendChild el

set el = dom.createElement("CountryCode") 'Sacar tabla de paises
el.text = rd("ShipperCountry")
subAdd.appendChild el

set subSubNode = dom.createElement("ShipTo")
subNode.appendChild subSubNode

set el = dom.createElement("CompanyName")
el.text = rd("CardName")
subSubNode.appendChild el

set subAdd = dom.createElement("Address")
subSubNode.appendChild subAdd

set el = dom.createElement("AddressLine1")
el.text = rd("ShipToStreet")
subAdd.appendChild el

set el = dom.createElement("AddressLine2")
el.text = rd("ShipToBlock")
subAdd.appendChild el

set el = dom.createElement("AddressLine3")
el.text = ""
subAdd.appendChild el

set el = dom.createElement("City") 'Required if no postal code
el.text = rd("ShipToCity")
subAdd.appendChild el

set el = dom.createElement("StateProvinceCode") 'Irland use 5 digit abbreviation for county
el.text = rd("ShipToState")
subAdd.appendChild el

set el = dom.createElement("PostalCode") 'Required use postal code
el.text = rd("ShipToZipCode")
subAdd.appendChild el

set el = dom.createElement("CountryCode") 'Sacar tabla de paises
el.text = rd("ShipToCountry")
subAdd.appendChild el

set el = dom.createElement("ResidentialAddressIndicator")
subAdd.appendChild el

set subSubNode = dom.createElement("ShipFrom")
subNode.appendChild subSubNode

set el = dom.createElement("CompanyName")
el.text = rd("CmpName")
subSubNode.appendChild el

set subAdd = dom.createElement("Address")
subSubNode.appendChild subAdd

set el = dom.createElement("AddressLine1")
el.text = rd("ShipFromStreet")
subAdd.appendChild el

set el = dom.createElement("AddressLine2")
el.text = rd("ShipFromBlock")
subAdd.appendChild el

set el = dom.createElement("AddressLine3")
el.text = ""
subAdd.appendChild el

set el = dom.createElement("City") 'Required if no postal code
el.text = rd("ShipFromCity")
subAdd.appendChild el

set el = dom.createElement("StateProvinceCode") 'Irland use 5 digit abbreviation for county
el.text = rd("ShipFromState")
subAdd.appendChild el

set el = dom.createElement("PostalCode") 'Required use postal code
el.text = rd("ShipFromZipCode")
subAdd.appendChild el

set el = dom.createElement("CountryCode") 'Sacar tabla de paises
el.text = rd("ShipFromCountry")
subAdd.appendChild el

'	***** Start Service ******

set subNode = dom.createElement("Service")
root.appendChild subNode

set el = dom.createElement("Code")
el.text = "" '{Code}
subNode.appendChild el

set el = dom.createElement("Description")
subNode.appendChild el

set el = dom.createElement("DocumentsOnly")
root.appendChild el

set subNode = dom.createElement("Package")
root.appendChild subNode 

Dim subPack
set subPack = dom.createElement("PackagingType")
subNode.appendChild subPack

rs.filter = "FieldID = 5"
set el = dom.createElement("Code")
el.text = rs(1)
subPack.appendChild el

set el = dom.createElement("Description")
subPack.appendChild el

Dim subDim
set subDim = dom.createElement("Dimensions")
subNode.appendChild subDim

Dim subUOM
set subUOM = dom.createElement("UnitOfMeasurement")
subDim.appendChild subUOM

set el = dom.createElement("Code")
'el.text = "" 'IN = Inches, CM = Centimeters
subUOM.appendChild el

set el = dom.createElement("Description")
subUOM.appendChild el

set el = dom.createElement("Length")
el.text = "{Length}"
subDim.appendChild el

set el = dom.createElement("Width")
el.text = "{Width}"
subDim.appendChild el

set el = dom.createElement("Height")
el.text = "Height"
subDim.appendChild el

Dim subWeight
set subWeight = dom.createElement("PackageWeight")
subNode.appendChild subWeight

set subUOM = dom.createElement("UnitOfMeasurement")
subWeight.appendChild subUOM

set el = dom.createElement("Code")
'el.text = "" 'LBS = Pounds, KGS = Kilos
subUOM.appendChild el

set el = dom.createElement("Description")
subUOM.appendChild el

set el = dom.createElement("Weight")
el.text = "{Weight}"
subWeight.appendChild el

set el = dom.createElement("LargePackageIndicator")
el.text = "" 'N = No, Y = Yes
subNode.appendChild el

set subOpt = dom.createElement("PackageServiceOptions")
subNode.appendChild subOpt

set subDel = dom.createElement("DeliveryConfirmation")
subOpt.appendChild subDel

rs.Filter = "FieldID = 6"
set el = dom.createElement("DCISType")
el.text = rs(1)
subDel.appendChild el

Dim subVer
set subVer = dom.createElement("VerbalConfirmation")
subOpt.appendChild subVer

set el = dom.createElement("Name")
el.text = rd("CntctPrsn")
subVer.appendChild el

set el = dom.createElement("PhoneNumber")
el.text = CStr(rd("Phone1"))
subVer.appendChild el

set subOpt = dom.createElement("ShipmentServiceOptions")
root.appendChild subOpt

set el = dom.createElement("SaturdayPickup")
el.text = "" 'N = No, Y = Yes
subOpt.appendChild el

set el = dom.createElement("SaturdayDelivery")
el.text = "" 'N = No, Y = Yes
subOpt.appendChild el

Dim subOnCall
set subOnCall = dom.createElement("OnCallAir")
subOpt.appendChild subOnCall

Dim subSch
set subSch = dom.createElement("Schedule")
subOnCall.appendChild subSch

rs.Filter = "FieldID = 7"
set el = dom.createElement("PickupDay")
el.text = rs(1)
subSch.appendChild el

rs.Filter = "FieldID = 8"
set el = dom.createElement("Method")
el.text = rs(1)
subSch.appendChild el

set subOpt= Nothing
set subAdd = Nothing
set subDim = Nothing
set subUOM = Nothing
set subSubNode = Nothing
set subNode = Nothing
set el = Nothing
set root = Nothing

strXml = strXml & dom.xml
'***** End RatingServiceSelectionRequest ******
%>
<html>
<body>
<textarea style="width: 100%; height: 100%; font-family: Verdana; font-size: xx-small;">
<%=Server.HTMLEncode(strXml)%>
</textarea>
</body>
</html>
<%

'https://wwwcie.ups.com/ups.app/xml/Rate

' create the object that manages the communication
	' Dim oXMLHttp As XMLHTTP
	' Set oXMLHttp = New XMLHTTP
' prepare the HTTP POST request
	' oXMLHttp.open "POST", "https://www.server.com/path", False
	' oXMLHttp.setRequestHeader "Content-Type", _
	' "application/x-www-form-urlencoded"
' send the request
	'oXMLHttp.send requestString
' server's response will be available in oXMLHttp.responseXML

' Define a variable and initialize it to a new XML message
	' Dim dom
	' Set dom = New DOMDocument30
' Set properties of the variable
	' dom.async = False
	' dom.validateOnParse = False
	' dom.resolveExternals = False
	' dom.preserveWhiteSpace = True
' Identify the message as XML version 1.0
	' Set node = dom.createProcessingInstruction("xml", "version='1.0'")
	' dom.appendChild node
	'Set node = Nothing
' Create the root (book) element and add it to the message
	' Dim root
	' Set root = dom.createElement("book")
	' dom.appendChild root

' Create child elements and add them to the root
	' Dim node
	' Set node = dom.createElement("title")
	' node.text = "HTTP Essentials: ..."
	' root.appendChild node
	' Set node = Nothing
	' Set node = dom.createElement("author")
	' Dim child
	' Set child = dom.createElement("firstname")
	' child.text = "Stephen"
	' node.appendChild child
	' Set child = Nothing
	' Set child = dom.createElement("lastname")
	' child.text = "Thomas"
	' node.appendChild child
	' root.appendChild node
%>