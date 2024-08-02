<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminCartMinRepEdit.xml"
set docadminCartMinRepEdit = server.CreateObject("MSXML2.DOMDocument")
docadminCartMinRepEdit.async = False
DocadminCartMinRepEdit.Load(server.MapPath(xmlfilename)) 
docadminCartMinRepEdit.setProperty "SelectionLanguage", "XPath"
set selectedadminCartMinRepEditnode = docadminCartMinRepEdit.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminCartMinRepEditnodes=docadminCartMinRepEdit.documentElement.selectNodes("/languages/language")
function getadminCartMinRepEditLngStr(instring)
	temp = selectedadminCartMinRepEditnode.selectSingleNode(instring).text
	getadminCartMinRepEditLngStr = temp
end function
%>
