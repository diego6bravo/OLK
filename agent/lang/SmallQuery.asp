<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "SmallQuery.xml"
set docSmallQuery = server.CreateObject("MSXML2.DOMDocument")
docSmallQuery.async = False
DocSmallQuery.Load(server.MapPath(xmlfilename)) 
docSmallQuery.setProperty "SelectionLanguage", "XPath"
set selectedSmallQuerynode = docSmallQuery.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedSmallQuerynodes=docSmallQuery.documentElement.selectNodes("/languages/language")
function getSmallQueryLngStr(instring)
	temp = selectedSmallQuerynode.selectSingleNode(instring).text
	getSmallQueryLngStr = temp
end function
%>
