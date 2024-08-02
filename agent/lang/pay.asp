<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "pay.xml"
set docpay = server.CreateObject("MSXML2.DOMDocument")
docpay.async = False
Docpay.Load(server.MapPath(xmlfilename)) 
docpay.setProperty "SelectionLanguage", "XPath"
set selectedpaynode = docpay.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedpaynodes=docpay.documentElement.selectNodes("/languages/language")
function getpayLngStr(instring)
	temp = selectedpaynode.selectSingleNode(instring).text
	getpayLngStr = temp
end function
%>
