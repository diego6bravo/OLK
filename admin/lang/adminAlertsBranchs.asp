<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminAlertsBranchs.xml"
set docadminAlertsBranchs = server.CreateObject("MSXML2.DOMDocument")
docadminAlertsBranchs.async = False
DocadminAlertsBranchs.Load(server.MapPath(xmlfilename)) 
docadminAlertsBranchs.setProperty "SelectionLanguage", "XPath"
set selectedadminAlertsBranchsnode = docadminAlertsBranchs.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminAlertsBranchsnodes=docadminAlertsBranchs.documentElement.selectNodes("/languages/language")
function getadminAlertsBranchsLngStr(instring)
	temp = selectedadminAlertsBranchsnode.selectSingleNode(instring).text
	getadminAlertsBranchsLngStr = temp
end function
%>
