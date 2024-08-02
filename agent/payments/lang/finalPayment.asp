<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "finalPayment.xml"
set docfinalPayment = server.CreateObject("MSXML2.DOMDocument")
docfinalPayment.async = False
DocfinalPayment.Load(server.MapPath(xmlfilename)) 
docfinalPayment.setProperty "SelectionLanguage", "XPath"
set selectedfinalPaymentnode = docfinalPayment.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedfinalPaymentnodes=docfinalPayment.documentElement.selectNodes("/languages/language")
function getfinalPaymentLngStr(instring)
	temp = selectedfinalPaymentnode.selectSingleNode(instring).text
	getfinalPaymentLngStr = temp
end function
%>
