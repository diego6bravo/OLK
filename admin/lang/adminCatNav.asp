<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminCatNav.xml"
set docadminCatNav = server.CreateObject("MSXML2.DOMDocument")
docadminCatNav.async = False
DocadminCatNav.Load(server.MapPath(xmlfilename)) 
docadminCatNav.setProperty "SelectionLanguage", "XPath"
set selectedadminCatNavnode = docadminCatNav.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminCatNavnodes=docadminCatNav.documentElement.selectNodes("/languages/language")
function getadminCatNavLngStr(instring)
	temp = selectedadminCatNavnode.selectSingleNode(instring).text
	getadminCatNavLngStr = temp
end function
%>
