<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminBN.xml"
set docadminBN = server.CreateObject("MSXML2.DOMDocument")
docadminBN.async = False
DocadminBN.Load(server.MapPath(xmlfilename)) 
docadminBN.setProperty "SelectionLanguage", "XPath"
set selectedadminBNnode = docadminBN.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminBNnodes=docadminBN.documentElement.selectNodes("/languages/language")
function getadminBNLngStr(instring)
	temp = selectedadminBNnode.selectSingleNode(instring).text
	getadminBNLngStr = temp
end function
%>
