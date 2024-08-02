<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "messageDetail.xml"
set docmessageDetail = server.CreateObject("MSXML2.DOMDocument")
docmessageDetail.async = False
DocmessageDetail.Load(server.MapPath(xmlfilename)) 
docmessageDetail.setProperty "SelectionLanguage", "XPath"
set selectedmessageDetailnode = docmessageDetail.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedmessageDetailnodes=docmessageDetail.documentElement.selectNodes("/languages/language")
function getmessageDetailLngStr(instring)
	temp = selectedmessageDetailnode.selectSingleNode(instring).text
	getmessageDetailLngStr = temp
end function
%>
