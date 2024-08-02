<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminBranchs.xml"
set docadminBranchs = server.CreateObject("MSXML2.DOMDocument")
docadminBranchs.async = False
DocadminBranchs.Load(server.MapPath(xmlfilename)) 
docadminBranchs.setProperty "SelectionLanguage", "XPath"
set selectedadminBranchsnode = docadminBranchs.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminBranchsnodes=docadminBranchs.documentElement.selectNodes("/languages/language")
function getadminBranchsLngStr(instring)
	temp = selectedadminBranchsnode.selectSingleNode(instring).text
	getadminBranchsLngStr = temp
end function
%>
