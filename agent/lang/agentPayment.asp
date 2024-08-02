<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "agentPayment.xml"
set docagentPayment = server.CreateObject("MSXML2.DOMDocument")
docagentPayment.async = False
DocagentPayment.Load(server.MapPath(xmlfilename)) 
docagentPayment.setProperty "SelectionLanguage", "XPath"
set selectedagentPaymentnode = docagentPayment.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedagentPaymentnodes=docagentPayment.documentElement.selectNodes("/languages/language")
function getagentPaymentLngStr(instring)
	temp = selectedagentPaymentnode.selectSingleNode(instring).text
	getagentPaymentLngStr = temp
end function
%>
