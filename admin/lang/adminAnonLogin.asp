<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminAnonLogin.xml"
set docadminAnonLogin = server.CreateObject("MSXML2.DOMDocument")
docadminAnonLogin.async = False
DocadminAnonLogin.Load(server.MapPath(xmlfilename)) 
docadminAnonLogin.setProperty "SelectionLanguage", "XPath"
set selectedadminAnonLoginnode = docadminAnonLogin.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminAnonLoginnodes=docadminAnonLogin.documentElement.selectNodes("/languages/language")
function getadminAnonLoginLngStr(instring)
	temp = selectedadminAnonLoginnode.selectSingleNode(instring).text
	getadminAnonLoginLngStr = temp
end function
%>
