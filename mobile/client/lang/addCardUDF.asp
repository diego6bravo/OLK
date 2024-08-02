<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "addCardUDF.xml"
set docaddCardUDF = server.CreateObject("MSXML2.DOMDocument")
docaddCardUDF.async = False
DocaddCardUDF.Load(server.MapPath(xmlfilename)) 
docaddCardUDF.setProperty "SelectionLanguage", "XPath"
set selectedaddCardUDFnode = docaddCardUDF.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedaddCardUDFnodes=docaddCardUDF.documentElement.selectNodes("/languages/language")
function getaddCardUDFLngStr(instring)
	temp = selectedaddCardUDFnode.selectSingleNode(instring).text
	getaddCardUDFLngStr = temp
end function
%>
