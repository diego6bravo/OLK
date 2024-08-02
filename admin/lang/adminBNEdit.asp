<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminBNEdit.xml"
set docadminBNEdit = server.CreateObject("MSXML2.DOMDocument")
docadminBNEdit.async = False
DocadminBNEdit.Load(server.MapPath(xmlfilename)) 
docadminBNEdit.setProperty "SelectionLanguage", "XPath"
set selectedadminBNEditnode = docadminBNEdit.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminBNEditnodes=docadminBNEdit.documentElement.selectNodes("/languages/language")
function getadminBNEditLngStr(instring)
	temp = selectedadminBNEditnode.selectSingleNode(instring).text
	getadminBNEditLngStr = temp
end function
%>
