<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "changePwdLogon.xml"
set docchangePwdLogon = server.CreateObject("MSXML2.DOMDocument")
docchangePwdLogon.async = False
DocchangePwdLogon.Load(server.MapPath(xmlfilename)) 
docchangePwdLogon.setProperty "SelectionLanguage", "XPath"
set selectedchangePwdLogonnode = docchangePwdLogon.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedchangePwdLogonnodes=docchangePwdLogon.documentElement.selectNodes("/languages/language")
function getchangePwdLogonLngStr(instring)
	temp = selectedchangePwdLogonnode.selectSingleNode(instring).text
	getchangePwdLogonLngStr = temp
end function
%>
