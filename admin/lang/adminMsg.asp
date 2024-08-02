<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminMsg.xml"
set docadminMsg = server.CreateObject("MSXML2.DOMDocument")
docadminMsg.async = False
DocadminMsg.Load(server.MapPath(xmlfilename)) 
docadminMsg.setProperty "SelectionLanguage", "XPath"
set selectedadminMsgnode = docadminMsg.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminMsgnodes=docadminMsg.documentElement.selectNodes("/languages/language")
function getadminMsgLngStr(instring)
	temp = selectedadminMsgnode.selectSingleNode(instring).text
	getadminMsgLngStr = temp
end function
%>
