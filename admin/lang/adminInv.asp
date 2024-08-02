<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminInv.xml"
set docadminInv = server.CreateObject("MSXML2.DOMDocument")
docadminInv.async = False
DocadminInv.Load(server.MapPath(xmlfilename)) 
docadminInv.setProperty "SelectionLanguage", "XPath"
set selectedadminInvnode = docadminInv.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminInvnodes=docadminInv.documentElement.selectNodes("/languages/language")
function getadminInvLngStr(instring)
	temp = selectedadminInvnode.selectSingleNode(instring).text
	getadminInvLngStr = temp
end function
%>
