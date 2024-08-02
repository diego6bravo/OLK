<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "dateFormat.xml"
set docdateFormat = server.CreateObject("MSXML2.DOMDocument")
docdateFormat.async = False
DocdateFormat.Load(server.MapPath(xmlfilename)) 
docdateFormat.setProperty "SelectionLanguage", "XPath"
set selecteddateFormatnode = docdateFormat.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selecteddateFormatnodes=docdateFormat.documentElement.selectNodes("/languages/language")
function getdateFormatLngStr(instring)
	temp = selecteddateFormatnode.selectSingleNode(instring).text
	getdateFormatLngStr = temp
end function
%>
