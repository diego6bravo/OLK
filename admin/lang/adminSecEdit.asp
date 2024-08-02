<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminSecEdit.xml"
set docadminSecEdit = server.CreateObject("MSXML2.DOMDocument")
docadminSecEdit.async = False
DocadminSecEdit.Load(server.MapPath(xmlfilename)) 
docadminSecEdit.setProperty "SelectionLanguage", "XPath"
set selectedadminSecEditnode = docadminSecEdit.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminSecEditnodes=docadminSecEdit.documentElement.selectNodes("/languages/language")
function getadminSecEditLngStr(instring)
	temp = selectedadminSecEditnode.selectSingleNode(instring).text
	getadminSecEditLngStr = temp
end function
%>
