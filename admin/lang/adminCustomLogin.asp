<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminCustomLogin.xml"
set docadminCustomLogin = server.CreateObject("MSXML2.DOMDocument")
docadminCustomLogin.async = False
DocadminCustomLogin.Load(server.MapPath(xmlfilename)) 
docadminCustomLogin.setProperty "SelectionLanguage", "XPath"
set selectedadminCustomLoginnode = docadminCustomLogin.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminCustomLoginnodes=docadminCustomLogin.documentElement.selectNodes("/languages/language")
function getadminCustomLoginLngStr(instring)
	temp = selectedadminCustomLoginnode.selectSingleNode(instring).text
	getadminCustomLoginLngStr = temp
end function
%>
