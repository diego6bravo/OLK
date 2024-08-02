<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminEditDec.xml"
set docadminEditDec = server.CreateObject("MSXML2.DOMDocument")
docadminEditDec.async = False
DocadminEditDec.Load(server.MapPath(xmlfilename)) 
docadminEditDec.setProperty "SelectionLanguage", "XPath"
set selectedadminEditDecnode = docadminEditDec.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminEditDecnodes=docadminEditDec.documentElement.selectNodes("/languages/language")
function getadminEditDecLngStr(instring)
	temp = selectedadminEditDecnode.selectSingleNode(instring).text
	getadminEditDecLngStr = temp
end function
%>
