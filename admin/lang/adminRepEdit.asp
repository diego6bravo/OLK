<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminRepEdit.xml"
set docadminRepEdit = server.CreateObject("MSXML2.DOMDocument")
docadminRepEdit.async = False
DocadminRepEdit.Load(server.MapPath(xmlfilename)) 
docadminRepEdit.setProperty "SelectionLanguage", "XPath"
set selectedadminRepEditnode = docadminRepEdit.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminRepEditnodes=docadminRepEdit.documentElement.selectNodes("/languages/language")
function getadminRepEditLngStr(instring)
	temp = selectedadminRepEditnode.selectSingleNode(instring).text
	getadminRepEditLngStr = temp
end function
%>
