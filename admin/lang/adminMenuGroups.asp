<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminMenuGroups.xml"
set docadminMenuGroups = server.CreateObject("MSXML2.DOMDocument")
docadminMenuGroups.async = False
DocadminMenuGroups.Load(server.MapPath(xmlfilename)) 
docadminMenuGroups.setProperty "SelectionLanguage", "XPath"
set selectedadminMenuGroupsnode = docadminMenuGroups.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminMenuGroupsnodes=docadminMenuGroups.documentElement.selectNodes("/languages/language")
function getadminMenuGroupsLngStr(instring)
	temp = selectedadminMenuGroupsnode.selectSingleNode(instring).text
	getadminMenuGroupsLngStr = temp
end function
%>
