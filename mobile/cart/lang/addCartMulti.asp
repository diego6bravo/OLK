<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "addCartMulti.xml"
set docaddCartMulti = server.CreateObject("MSXML2.DOMDocument")
docaddCartMulti.async = False
DocaddCartMulti.Load(server.MapPath(xmlfilename)) 
docaddCartMulti.setProperty "SelectionLanguage", "XPath"
set selectedaddCartMultinode = docaddCartMulti.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedaddCartMultinodes=docaddCartMulti.documentElement.selectNodes("/languages/language")
function getaddCartMultiLngStr(instring)
	temp = selectedaddCartMultinode.selectSingleNode(instring).text
	getaddCartMultiLngStr = temp
end function
%>
