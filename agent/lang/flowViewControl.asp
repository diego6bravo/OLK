<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "flowViewControl.xml"
set docflowViewControl = server.CreateObject("MSXML2.DOMDocument")
docflowViewControl.async = False
DocflowViewControl.Load(server.MapPath(xmlfilename)) 
docflowViewControl.setProperty "SelectionLanguage", "XPath"
set selectedflowViewControlnode = docflowViewControl.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedflowViewControlnodes=docflowViewControl.documentElement.selectNodes("/languages/language")
function getflowViewControlLngStr(instring)
	temp = selectedflowViewControlnode.selectSingleNode(instring).text
	getflowViewControlLngStr = temp
end function
%>
