<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminAutGrp.xml"
set docadminAutGrp = server.CreateObject("MSXML2.DOMDocument")
docadminAutGrp.async = False
DocadminAutGrp.Load(server.MapPath(xmlfilename)) 
docadminAutGrp.setProperty "SelectionLanguage", "XPath"
set selectedadminAutGrpnode = docadminAutGrp.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminAutGrpnodes=docadminAutGrp.documentElement.selectNodes("/languages/language")
function getadminAutGrpLngStr(instring)
	temp = selectedadminAutGrpnode.selectSingleNode(instring).text
	getadminAutGrpLngStr = temp
end function
%>
