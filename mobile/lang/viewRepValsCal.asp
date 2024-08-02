<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "viewRepValsCal.xml"
set docviewRepValsCal = server.CreateObject("MSXML2.DOMDocument")
docviewRepValsCal.async = False
DocviewRepValsCal.Load(server.MapPath(xmlfilename)) 
docviewRepValsCal.setProperty "SelectionLanguage", "XPath"
set selectedviewRepValsCalnode = docviewRepValsCal.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedviewRepValsCalnodes=docviewRepValsCal.documentElement.selectNodes("/languages/language")
function getviewRepValsCalLngStr(instring)
	temp = selectedviewRepValsCalnode.selectSingleNode(instring).text
	getviewRepValsCalLngStr = temp
end function
%>
