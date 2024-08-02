<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "soCancel.xml"
set docsoCancel = server.CreateObject("MSXML2.DOMDocument")
docsoCancel.async = False
DocsoCancel.Load(server.MapPath(xmlfilename)) 
docsoCancel.setProperty "SelectionLanguage", "XPath"
set selectedsoCancelnode = docsoCancel.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedsoCancelnodes=docsoCancel.documentElement.selectNodes("/languages/language")
function getsoCancelLngStr(instring)
	temp = selectedsoCancelnode.selectSingleNode(instring).text
	getsoCancelLngStr = temp
end function
%>
