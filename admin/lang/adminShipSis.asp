<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminShipSis.xml"
set docadminShipSis = server.CreateObject("MSXML2.DOMDocument")
docadminShipSis.async = False
DocadminShipSis.Load(server.MapPath(xmlfilename)) 
docadminShipSis.setProperty "SelectionLanguage", "XPath"
set selectedadminShipSisnode = docadminShipSis.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminShipSisnodes=docadminShipSis.documentElement.selectNodes("/languages/language")
function getadminShipSisLngStr(instring)
	temp = selectedadminShipSisnode.selectSingleNode(instring).text
	getadminShipSisLngStr = temp
end function
%>
