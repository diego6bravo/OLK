<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "messagedetail.xml"
set docmessagedetail = server.CreateObject("MSXML2.DOMDocument")
docmessagedetail.async = False
Docmessagedetail.Load(server.MapPath(xmlfilename)) 
docmessagedetail.setProperty "SelectionLanguage", "XPath"
set selectedmessagedetailnode = docmessagedetail.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedmessagedetailnodes=docmessagedetail.documentElement.selectNodes("/languages/language")
function getmessagedetailLngStr(instring)
	temp = selectedmessagedetailnode.selectSingleNode(instring).text
	getmessagedetailLngStr = temp
end function
%>
