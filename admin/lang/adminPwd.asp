<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminPwd.xml"
set docadminPwd = server.CreateObject("MSXML2.DOMDocument")
docadminPwd.async = False
DocadminPwd.Load(server.MapPath(xmlfilename)) 
docadminPwd.setProperty "SelectionLanguage", "XPath"
set selectedadminPwdnode = docadminPwd.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminPwdnodes=docadminPwd.documentElement.selectNodes("/languages/language")
function getadminPwdLngStr(instring)
	temp = selectedadminPwdnode.selectSingleNode(instring).text
	getadminPwdLngStr = temp
end function
%>
