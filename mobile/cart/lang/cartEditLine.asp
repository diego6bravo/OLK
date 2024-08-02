<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "cartEditLine.xml"
set doccartEditLine = server.CreateObject("MSXML2.DOMDocument")
doccartEditLine.async = False
DoccartEditLine.Load(server.MapPath(xmlfilename)) 
doccartEditLine.setProperty "SelectionLanguage", "XPath"
set selectedcartEditLinenode = doccartEditLine.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedcartEditLinenodes=doccartEditLine.documentElement.selectNodes("/languages/language")
function getcartEditLineLngStr(instring)
	temp = selectedcartEditLinenode.selectSingleNode(instring).text
	getcartEditLineLngStr = temp
end function
%>
