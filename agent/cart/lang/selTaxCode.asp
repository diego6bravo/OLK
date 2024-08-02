<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "selTaxCode.xml"
set docselTaxCode = server.CreateObject("MSXML2.DOMDocument")
docselTaxCode.async = False
DocselTaxCode.Load(server.MapPath(xmlfilename)) 
docselTaxCode.setProperty "SelectionLanguage", "XPath"
set selectedselTaxCodenode = docselTaxCode.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedselTaxCodenodes=docselTaxCode.documentElement.selectNodes("/languages/language")
function getselTaxCodeLngStr(instring)
	temp = selectedselTaxCodenode.selectSingleNode(instring).text
	getselTaxCodeLngStr = temp
end function
%>
