<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "submitPayment.xml"
set docsubmitPayment = server.CreateObject("MSXML2.DOMDocument")
docsubmitPayment.async = False
DocsubmitPayment.Load(server.MapPath(xmlfilename)) 
docsubmitPayment.setProperty "SelectionLanguage", "XPath"
set selectedsubmitPaymentnode = docsubmitPayment.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedsubmitPaymentnodes=docsubmitPayment.documentElement.selectNodes("/languages/language")
function getsubmitPaymentLngStr(instring)
	temp = selectedsubmitPaymentnode.selectSingleNode(instring).text
	getsubmitPaymentLngStr = temp
end function
%>
