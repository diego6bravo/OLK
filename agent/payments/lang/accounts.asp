<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "accounts.xml"
set docaccounts = server.CreateObject("MSXML2.DOMDocument")
docaccounts.async = False
Docaccounts.Load(server.MapPath(xmlfilename)) 
docaccounts.setProperty "SelectionLanguage", "XPath"
set selectedaccountsnode = docaccounts.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedaccountsnodes=docaccounts.documentElement.selectNodes("/languages/language")
function getaccountsLngStr(instring)
	temp = selectedaccountsnode.selectSingleNode(instring).text
	getaccountsLngStr = temp
end function
%>
