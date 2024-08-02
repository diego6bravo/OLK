<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "clientAddresses.xml"
set docclientAddresses = server.CreateObject("MSXML2.DOMDocument")
docclientAddresses.async = False
DocclientAddresses.Load(server.MapPath(xmlfilename)) 
docclientAddresses.setProperty "SelectionLanguage", "XPath"
set selectedclientAddressesnode = docclientAddresses.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedclientAddressesnodes=docclientAddresses.documentElement.selectNodes("/languages/language")
function getclientAddressesLngStr(instring)
	temp = selectedclientAddressesnode.selectSingleNode(instring).text
	getclientAddressesLngStr = temp
end function
%>
