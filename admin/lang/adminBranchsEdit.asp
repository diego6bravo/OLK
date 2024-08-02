<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminBranchsEdit.xml"
set docadminBranchsEdit = server.CreateObject("MSXML2.DOMDocument")
docadminBranchsEdit.async = False
DocadminBranchsEdit.Load(server.MapPath(xmlfilename)) 
docadminBranchsEdit.setProperty "SelectionLanguage", "XPath"
set selectedadminBranchsEditnode = docadminBranchsEdit.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminBranchsEditnodes=docadminBranchsEdit.documentElement.selectNodes("/languages/language")
function getadminBranchsEditLngStr(instring)
	temp = selectedadminBranchsEditnode.selectSingleNode(instring).text
	getadminBranchsEditLngStr = temp
end function
%>
