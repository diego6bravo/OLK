<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "addActivity.xml"
set docaddActivity = server.CreateObject("MSXML2.DOMDocument")
docaddActivity.async = False
DocaddActivity.Load(server.MapPath(xmlfilename)) 
docaddActivity.setProperty "SelectionLanguage", "XPath"
set selectedaddActivitynode = docaddActivity.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedaddActivitynodes=docaddActivity.documentElement.selectNodes("/languages/language")
function getaddActivityLngStr(instring)
	temp = selectedaddActivitynode.selectSingleNode(instring).text
	getaddActivityLngStr = temp
end function
%>
