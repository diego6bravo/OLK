<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "payCheck.xml"
set docpayCheck = server.CreateObject("MSXML2.DOMDocument")
docpayCheck.async = False
DocpayCheck.Load(server.MapPath(xmlfilename)) 
docpayCheck.setProperty "SelectionLanguage", "XPath"
set selectedpayChecknode = docpayCheck.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedpayChecknodes=docpayCheck.documentElement.selectNodes("/languages/language")
function getpayCheckLngStr(instring)
	temp = selectedpayChecknode.selectSingleNode(instring).text
	getpayCheckLngStr = temp
end function
%>
