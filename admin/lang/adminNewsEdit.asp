<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminNewsEdit.xml"
set docadminNewsEdit = server.CreateObject("MSXML2.DOMDocument")
docadminNewsEdit.async = False
DocadminNewsEdit.Load(server.MapPath(xmlfilename)) 
docadminNewsEdit.setProperty "SelectionLanguage", "XPath"
set selectedadminNewsEditnode = docadminNewsEdit.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminNewsEditnodes=docadminNewsEdit.documentElement.selectNodes("/languages/language")
function getadminNewsEditLngStr(instring)
	temp = selectedadminNewsEditnode.selectSingleNode(instring).text
	getadminNewsEditLngStr = temp
end function
%>
