<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "activeClient.xml"
set docactiveClient = server.CreateObject("MSXML2.DOMDocument")
docactiveClient.async = False
DocactiveClient.Load(server.MapPath(xmlfilename)) 
docactiveClient.setProperty "SelectionLanguage", "XPath"
set selectedactiveClientnode = docactiveClient.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedactiveClientnodes=docactiveClient.documentElement.selectNodes("/languages/language")
function getactiveClientLngStr(instring)
	temp = selectedactiveClientnode.selectSingleNode(instring).text
	getactiveClientLngStr = temp
end function
%>
