<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminOpsEdit.xml"
set docadminOpsEdit = server.CreateObject("MSXML2.DOMDocument")
docadminOpsEdit.async = False
DocadminOpsEdit.Load(server.MapPath(xmlfilename)) 
docadminOpsEdit.setProperty "SelectionLanguage", "XPath"
set selectedadminOpsEditnode = docadminOpsEdit.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminOpsEditnodes=docadminOpsEdit.documentElement.selectNodes("/languages/language")
function getadminOpsEditLngStr(instring)
	temp = selectedadminOpsEditnode.selectSingleNode(instring).text
	getadminOpsEditLngStr = temp
end function
%>
