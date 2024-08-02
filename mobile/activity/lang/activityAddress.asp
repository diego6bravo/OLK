<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "activityAddress.xml"
set docactivityAddress = server.CreateObject("MSXML2.DOMDocument")
docactivityAddress.async = False
DocactivityAddress.Load(server.MapPath(xmlfilename)) 
docactivityAddress.setProperty "SelectionLanguage", "XPath"
set selectedactivityAddressnode = docactivityAddress.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedactivityAddressnodes=docactivityAddress.documentElement.selectNodes("/languages/language")
function getactivityAddressLngStr(instring)
	temp = selectedactivityAddressnode.selectSingleNode(instring).text
	getactivityAddressLngStr = temp
end function
%>
