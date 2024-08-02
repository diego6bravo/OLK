<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "addresses.xml"
set docaddresses = server.CreateObject("MSXML2.DOMDocument")
docaddresses.async = False
Docaddresses.Load(server.MapPath(xmlfilename)) 
docaddresses.setProperty "SelectionLanguage", "XPath"
set selectedaddressesnode = docaddresses.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedaddressesnodes=docaddresses.documentElement.selectNodes("/languages/language")
function getaddressesLngStr(instring)
	temp = selectedaddressesnode.selectSingleNode(instring).text
	getaddressesLngStr = temp
end function
%>
