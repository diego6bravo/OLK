<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "cartEditExpLine.xml"
set doccartEditExpLine = server.CreateObject("MSXML2.DOMDocument")
doccartEditExpLine.async = False
DoccartEditExpLine.Load(server.MapPath(xmlfilename)) 
doccartEditExpLine.setProperty "SelectionLanguage", "XPath"
set selectedcartEditExpLinenode = doccartEditExpLine.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedcartEditExpLinenodes=doccartEditExpLine.documentElement.selectNodes("/languages/language")
function getcartEditExpLineLngStr(instring)
	temp = selectedcartEditExpLinenode.selectSingleNode(instring).text
	getcartEditExpLineLngStr = temp
end function
%>
