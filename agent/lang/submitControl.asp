<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "submitControl.xml"
set docsubmitControl = server.CreateObject("MSXML2.DOMDocument")
docsubmitControl.async = False
DocsubmitControl.Load(server.MapPath(xmlfilename)) 
docsubmitControl.setProperty "SelectionLanguage", "XPath"
set selectedsubmitControlnode = docsubmitControl.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedsubmitControlnodes=docsubmitControl.documentElement.selectNodes("/languages/language")
function getsubmitControlLngStr(instring)
	temp = selectedsubmitControlnode.selectSingleNode(instring).text
	getsubmitControlLngStr = temp
end function
%>
