<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "viewRepValsCL.xml"
set docviewRepValsCL = server.CreateObject("MSXML2.DOMDocument")
docviewRepValsCL.async = False
DocviewRepValsCL.Load(server.MapPath(xmlfilename)) 
docviewRepValsCL.setProperty "SelectionLanguage", "XPath"
set selectedviewRepValsCLnode = docviewRepValsCL.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedviewRepValsCLnodes=docviewRepValsCL.documentElement.selectNodes("/languages/language")
function getviewRepValsCLLngStr(instring)
	temp = selectedviewRepValsCLnode.selectSingleNode(instring).text
	getviewRepValsCLLngStr = temp
end function
%>
