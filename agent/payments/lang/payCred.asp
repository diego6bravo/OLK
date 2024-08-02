<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "payCred.xml"
set docpayCred = server.CreateObject("MSXML2.DOMDocument")
docpayCred.async = False
DocpayCred.Load(server.MapPath(xmlfilename)) 
docpayCred.setProperty "SelectionLanguage", "XPath"
set selectedpayCrednode = docpayCred.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedpayCrednodes=docpayCred.documentElement.selectNodes("/languages/language")
function getpayCredLngStr(instring)
	temp = selectedpayCrednode.selectSingleNode(instring).text
	getpayCredLngStr = temp
end function
%>
