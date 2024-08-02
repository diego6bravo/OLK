<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "viewType.xml"
set docviewType = server.CreateObject("MSXML2.DOMDocument")
docviewType.async = False
DocviewType.Load(server.MapPath(xmlfilename)) 
docviewType.setProperty "SelectionLanguage", "XPath"
set selectedviewTypenode = docviewType.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedviewTypenodes=docviewType.documentElement.selectNodes("/languages/language")
function getviewTypeLngStr(instring)
	temp = selectedviewTypenode.selectSingleNode(instring).text
	getviewTypeLngStr = temp
end function
%>
