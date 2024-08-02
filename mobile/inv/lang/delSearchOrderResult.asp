<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "delSearchOrderResult.xml"
set docdelSearchOrderResult = server.CreateObject("MSXML2.DOMDocument")
docdelSearchOrderResult.async = False
DocdelSearchOrderResult.Load(server.MapPath(xmlfilename)) 
docdelSearchOrderResult.setProperty "SelectionLanguage", "XPath"
set selecteddelSearchOrderResultnode = docdelSearchOrderResult.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selecteddelSearchOrderResultnodes=docdelSearchOrderResult.documentElement.selectNodes("/languages/language")
function getdelSearchOrderResultLngStr(instring)
	temp = selecteddelSearchOrderResultnode.selectSingleNode(instring).text
	getdelSearchOrderResultLngStr = temp
end function
%>
