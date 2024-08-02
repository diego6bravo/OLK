<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminSubmit.xml"
set docadminSubmit = server.CreateObject("MSXML2.DOMDocument")
docadminSubmit.async = False
DocadminSubmit.Load(server.MapPath(xmlfilename)) 
docadminSubmit.setProperty "SelectionLanguage", "XPath"
set selectedadminSubmitnode = docadminSubmit.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminSubmitnodes=docadminSubmit.documentElement.selectNodes("/languages/language")
function getadminSubmitLngStr(instring)
	temp = selectedadminSubmitnode.selectSingleNode(instring).text
	getadminSubmitLngStr = temp
end function
%>
