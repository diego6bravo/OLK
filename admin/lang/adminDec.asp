<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminDec.xml"
set docadminDec = server.CreateObject("MSXML2.DOMDocument")
docadminDec.async = False
DocadminDec.Load(server.MapPath(xmlfilename)) 
docadminDec.setProperty "SelectionLanguage", "XPath"
set selectedadminDecnode = docadminDec.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminDecnodes=docadminDec.documentElement.selectNodes("/languages/language")
function getadminDecLngStr(instring)
	temp = selectedadminDecnode.selectSingleNode(instring).text
	getadminDecLngStr = temp
end function
%>
