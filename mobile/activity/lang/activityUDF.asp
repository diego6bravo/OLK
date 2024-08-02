<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "activityUDF.xml"
set docactivityUDF = server.CreateObject("MSXML2.DOMDocument")
docactivityUDF.async = False
DocactivityUDF.Load(server.MapPath(xmlfilename)) 
docactivityUDF.setProperty "SelectionLanguage", "XPath"
set selectedactivityUDFnode = docactivityUDF.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedactivityUDFnodes=docactivityUDF.documentElement.selectNodes("/languages/language")
function getactivityUDFLngStr(instring)
	temp = selectedactivityUDFnode.selectSingleNode(instring).text
	getactivityUDFLngStr = temp
end function
%>
