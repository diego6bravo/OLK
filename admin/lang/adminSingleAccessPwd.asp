<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminSingleAccessPwd.xml"
set docadminSingleAccessPwd = server.CreateObject("MSXML2.DOMDocument")
docadminSingleAccessPwd.async = False
DocadminSingleAccessPwd.Load(server.MapPath(xmlfilename)) 
docadminSingleAccessPwd.setProperty "SelectionLanguage", "XPath"
set selectedadminSingleAccessPwdnode = docadminSingleAccessPwd.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminSingleAccessPwdnodes=docadminSingleAccessPwd.documentElement.selectNodes("/languages/language")
function getadminSingleAccessPwdLngStr(instring)
	temp = selectedadminSingleAccessPwdnode.selectSingleNode(instring).text
	getadminSingleAccessPwdLngStr = temp
end function
%>
