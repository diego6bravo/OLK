<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "addSO.js.xml"
set docaddSOjs = server.CreateObject("MSXML2.DOMDocument")
docaddSOjs.async = False
DocaddSOjs.Load(server.MapPath(xmlfilename)) 
docaddSOjs.setProperty "SelectionLanguage", "XPath"
set selectedaddSOjsnode = docaddSOjs.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedaddSOjsnodes=docaddSOjs.documentElement.selectNodes("/languages/language")
function getaddSOjsLngStr(instring)
	temp = selectedaddSOjsnode.selectSingleNode(instring).text
	getaddSOjsLngStr = temp
end function
%>
