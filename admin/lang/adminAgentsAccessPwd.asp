<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminAgentsAccessPwd.xml"
set docadminAgentsAccessPwd = server.CreateObject("MSXML2.DOMDocument")
docadminAgentsAccessPwd.async = False
DocadminAgentsAccessPwd.Load(server.MapPath(xmlfilename)) 
docadminAgentsAccessPwd.setProperty "SelectionLanguage", "XPath"
set selectedadminAgentsAccessPwdnode = docadminAgentsAccessPwd.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminAgentsAccessPwdnodes=docadminAgentsAccessPwd.documentElement.selectNodes("/languages/language")
function getadminAgentsAccessPwdLngStr(instring)
	temp = selectedadminAgentsAccessPwdnode.selectSingleNode(instring).text
	getadminAgentsAccessPwdLngStr = temp
end function
%>
