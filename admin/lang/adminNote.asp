<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminNote.xml"
set docadminNote = server.CreateObject("MSXML2.DOMDocument")
docadminNote.async = False
DocadminNote.Load(server.MapPath(xmlfilename)) 
docadminNote.setProperty "SelectionLanguage", "XPath"
set selectedadminNotenode = docadminNote.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminNotenodes=docadminNote.documentElement.selectNodes("/languages/language")
function getadminNoteLngStr(instring)
	temp = selectedadminNotenode.selectSingleNode(instring).text
	getadminNoteLngStr = temp
end function
%>
