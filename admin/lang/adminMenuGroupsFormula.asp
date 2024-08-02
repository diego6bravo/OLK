<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminMenuGroupsFormula.xml"
set docadminMenuGroupsFormula = server.CreateObject("MSXML2.DOMDocument")
docadminMenuGroupsFormula.async = False
DocadminMenuGroupsFormula.Load(server.MapPath(xmlfilename)) 
docadminMenuGroupsFormula.setProperty "SelectionLanguage", "XPath"
set selectedadminMenuGroupsFormulanode = docadminMenuGroupsFormula.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminMenuGroupsFormulanodes=docadminMenuGroupsFormula.documentElement.selectNodes("/languages/language")
function getadminMenuGroupsFormulaLngStr(instring)
	temp = selectedadminMenuGroupsFormulanode.selectSingleNode(instring).text
	getadminMenuGroupsFormulaLngStr = temp
end function
%>
