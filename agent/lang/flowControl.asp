<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "flowControl.xml"
set docflowControl = server.CreateObject("MSXML2.DOMDocument")
docflowControl.async = False
DocflowControl.Load(server.MapPath(xmlfilename)) 
docflowControl.setProperty "SelectionLanguage", "XPath"
set selectedflowControlnode = docflowControl.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedflowControlnodes=docflowControl.documentElement.selectNodes("/languages/language")
function getflowControlLngStr(instring)
	temp = selectedflowControlnode.selectSingleNode(instring).text
	getflowControlLngStr = temp
end function
%>
