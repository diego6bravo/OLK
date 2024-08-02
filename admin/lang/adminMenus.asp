<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminMenus.xml"
set docadminMenus = server.CreateObject("MSXML2.DOMDocument")
docadminMenus.async = False
DocadminMenus.Load(server.MapPath(xmlfilename)) 
docadminMenus.setProperty "SelectionLanguage", "XPath"
set selectedadminMenusnode = docadminMenus.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminMenusnodes=docadminMenus.documentElement.selectNodes("/languages/language")
function getadminMenusLngStr(instring)
	temp = selectedadminMenusnode.selectSingleNode(instring).text
	getadminMenusLngStr = temp
end function
%>
