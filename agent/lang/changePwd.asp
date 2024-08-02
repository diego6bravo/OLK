<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "changePwd.xml"
set docchangePwd = server.CreateObject("MSXML2.DOMDocument")
docchangePwd.async = False
DocchangePwd.Load(server.MapPath(xmlfilename)) 
docchangePwd.setProperty "SelectionLanguage", "XPath"
set selectedchangePwdnode = docchangePwd.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedchangePwdnodes=docchangePwd.documentElement.selectNodes("/languages/language")
function getchangePwdLngStr(instring)
	temp = selectedchangePwdnode.selectSingleNode(instring).text
	getchangePwdLngStr = temp
end function
%>
