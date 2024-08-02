<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "UDFQry.xml"
set docUDFQry = server.CreateObject("MSXML2.DOMDocument")
docUDFQry.async = False
DocUDFQry.Load(server.MapPath(xmlfilename)) 
docUDFQry.setProperty "SelectionLanguage", "XPath"
set selectedUDFQrynode = docUDFQry.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedUDFQrynodes=docUDFQry.documentElement.selectNodes("/languages/language")
function getUDFQryLngStr(instring)
	temp = selectedUDFQrynode.selectSingleNode(instring).text
	getUDFQryLngStr = temp
end function
%>
