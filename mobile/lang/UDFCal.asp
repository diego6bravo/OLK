<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "UDFCal.xml"
set docUDFCal = server.CreateObject("MSXML2.DOMDocument")
docUDFCal.async = False
DocUDFCal.Load(server.MapPath(xmlfilename)) 
docUDFCal.setProperty "SelectionLanguage", "XPath"
set selectedUDFCalnode = docUDFCal.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedUDFCalnodes=docUDFCal.documentElement.selectNodes("/languages/language")
function getUDFCalLngStr(instring)
	temp = selectedUDFCalnode.selectSingleNode(instring).text
	getUDFCalLngStr = temp
end function
%>
