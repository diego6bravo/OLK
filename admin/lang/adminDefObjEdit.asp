<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminDefObjEdit.xml"
set docadminDefObjEdit = server.CreateObject("MSXML2.DOMDocument")
docadminDefObjEdit.async = False
DocadminDefObjEdit.Load(server.MapPath(xmlfilename)) 
docadminDefObjEdit.setProperty "SelectionLanguage", "XPath"
set selectedadminDefObjEditnode = docadminDefObjEdit.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminDefObjEditnodes=docadminDefObjEdit.documentElement.selectNodes("/languages/language")
function getadminDefObjEditLngStr(instring)
	temp = selectedadminDefObjEditnode.selectSingleNode(instring).text
	getadminDefObjEditLngStr = temp
end function
%>
