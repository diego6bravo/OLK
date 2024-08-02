<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminCustomSearchEdit.xml"
set docadminCustomSearchEdit = server.CreateObject("MSXML2.DOMDocument")
docadminCustomSearchEdit.async = False
DocadminCustomSearchEdit.Load(server.MapPath(xmlfilename)) 
docadminCustomSearchEdit.setProperty "SelectionLanguage", "XPath"
set selectedadminCustomSearchEditnode = docadminCustomSearchEdit.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminCustomSearchEditnodes=docadminCustomSearchEdit.documentElement.selectNodes("/languages/language")
function getadminCustomSearchEditLngStr(instring)
	temp = selectedadminCustomSearchEditnode.selectSingleNode(instring).text
	getadminCustomSearchEditLngStr = temp
end function
%>
