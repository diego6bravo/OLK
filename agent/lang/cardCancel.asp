<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "cardCancel.xml"
set doccardCancel = server.CreateObject("MSXML2.DOMDocument")
doccardCancel.async = False
DoccardCancel.Load(server.MapPath(xmlfilename)) 
doccardCancel.setProperty "SelectionLanguage", "XPath"
set selectedcardCancelnode = doccardCancel.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedcardCancelnodes=doccardCancel.documentElement.selectNodes("/languages/language")
function getcardCancelLngStr(instring)
	temp = selectedcardCancelnode.selectSingleNode(instring).text
	getcardCancelLngStr = temp
end function
%>
