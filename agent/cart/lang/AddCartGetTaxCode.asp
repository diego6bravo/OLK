<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "AddCartGetTaxCode.xml"
set docAddCartGetTaxCode = server.CreateObject("MSXML2.DOMDocument")
docAddCartGetTaxCode.async = False
DocAddCartGetTaxCode.Load(server.MapPath(xmlfilename)) 
docAddCartGetTaxCode.setProperty "SelectionLanguage", "XPath"
set selectedAddCartGetTaxCodenode = docAddCartGetTaxCode.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedAddCartGetTaxCodenodes=docAddCartGetTaxCode.documentElement.selectNodes("/languages/language")
function getAddCartGetTaxCodeLngStr(instring)
	temp = selectedAddCartGetTaxCodenode.selectSingleNode(instring).text
	getAddCartGetTaxCodeLngStr = temp
end function
%>
