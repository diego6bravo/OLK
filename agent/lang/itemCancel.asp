<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "itemCancel.xml"
set docitemCancel = server.CreateObject("MSXML2.DOMDocument")
docitemCancel.async = False
DocitemCancel.Load(server.MapPath(xmlfilename)) 
docitemCancel.setProperty "SelectionLanguage", "XPath"
set selecteditemCancelnode = docitemCancel.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selecteditemCancelnodes=docitemCancel.documentElement.selectNodes("/languages/language")
function getitemCancelLngStr(instring)
	temp = selecteditemCancelnode.selectSingleNode(instring).text
	getitemCancelLngStr = temp
end function
%>
