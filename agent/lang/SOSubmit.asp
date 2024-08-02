<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "SOSubmit.xml"
set docSOSubmit = server.CreateObject("MSXML2.DOMDocument")
docSOSubmit.async = False
DocSOSubmit.Load(server.MapPath(xmlfilename)) 
docSOSubmit.setProperty "SelectionLanguage", "XPath"
set selectedSOSubmitnode = docSOSubmit.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedSOSubmitnodes=docSOSubmit.documentElement.selectNodes("/languages/language")
function getSOSubmitLngStr(instring)
	temp = selectedSOSubmitnode.selectSingleNode(instring).text
	getSOSubmitLngStr = temp
end function
%>
