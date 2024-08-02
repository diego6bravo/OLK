<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "viewRepVals.xml"
set docviewRepVals = server.CreateObject("MSXML2.DOMDocument")
docviewRepVals.async = False
DocviewRepVals.Load(server.MapPath(xmlfilename)) 
docviewRepVals.setProperty "SelectionLanguage", "XPath"
set selectedviewRepValsnode = docviewRepVals.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedviewRepValsnodes=docviewRepVals.documentElement.selectNodes("/languages/language")
function getviewRepValsLngStr(instring)
	temp = selectedviewRepValsnode.selectSingleNode(instring).text
	getviewRepValsLngStr = temp
end function
%>
