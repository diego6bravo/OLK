<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminRepEditSaveAs.xml"
set docadminRepEditSaveAs = server.CreateObject("MSXML2.DOMDocument")
docadminRepEditSaveAs.async = False
DocadminRepEditSaveAs.Load(server.MapPath(xmlfilename)) 
docadminRepEditSaveAs.setProperty "SelectionLanguage", "XPath"
set selectedadminRepEditSaveAsnode = docadminRepEditSaveAs.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminRepEditSaveAsnodes=docadminRepEditSaveAs.documentElement.selectNodes("/languages/language")
function getadminRepEditSaveAsLngStr(instring)
	temp = selectedadminRepEditSaveAsnode.selectSingleNode(instring).text
	getadminRepEditSaveAsLngStr = temp
end function
%>
