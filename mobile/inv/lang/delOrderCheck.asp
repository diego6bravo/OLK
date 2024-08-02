<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "delOrderCheck.xml"
set docdelOrderCheck = server.CreateObject("MSXML2.DOMDocument")
docdelOrderCheck.async = False
DocdelOrderCheck.Load(server.MapPath(xmlfilename)) 
docdelOrderCheck.setProperty "SelectionLanguage", "XPath"
set selecteddelOrderChecknode = docdelOrderCheck.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selecteddelOrderChecknodes=docdelOrderCheck.documentElement.selectNodes("/languages/language")
function getdelOrderCheckLngStr(instring)
	temp = selecteddelOrderChecknode.selectSingleNode(instring).text
	getdelOrderCheckLngStr = temp
end function
%>
