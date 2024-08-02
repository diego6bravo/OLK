<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "clientAddress.xml"
set docclientAddress = server.CreateObject("MSXML2.DOMDocument")
docclientAddress.async = False
DocclientAddress.Load(server.MapPath(xmlfilename)) 
docclientAddress.setProperty "SelectionLanguage", "XPath"
set selectedclientAddressnode = docclientAddress.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedclientAddressnodes=docclientAddress.documentElement.selectNodes("/languages/language")
function getclientAddressLngStr(instring)
	temp = selectedclientAddressnode.selectSingleNode(instring).text
	getclientAddressLngStr = temp
end function
%>
