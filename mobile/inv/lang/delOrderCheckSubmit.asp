<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "delOrderCheckSubmit.xml"
set docdelOrderCheckSubmit = server.CreateObject("MSXML2.DOMDocument")
docdelOrderCheckSubmit.async = False
DocdelOrderCheckSubmit.Load(server.MapPath(xmlfilename)) 
docdelOrderCheckSubmit.setProperty "SelectionLanguage", "XPath"
set selecteddelOrderCheckSubmitnode = docdelOrderCheckSubmit.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selecteddelOrderCheckSubmitnodes=docdelOrderCheckSubmit.documentElement.selectNodes("/languages/language")
function getdelOrderCheckSubmitLngStr(instring)
	temp = selecteddelOrderCheckSubmitnode.selectSingleNode(instring).text
	getdelOrderCheckSubmitLngStr = temp
end function
%>
