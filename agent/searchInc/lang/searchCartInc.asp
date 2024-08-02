<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "searchCartInc.xml"
set docsearchCartInc = server.CreateObject("MSXML2.DOMDocument")
docsearchCartInc.async = False
DocsearchCartInc.Load(server.MapPath(xmlfilename)) 
docsearchCartInc.setProperty "SelectionLanguage", "XPath"
set selectedsearchCartIncnode = docsearchCartInc.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedsearchCartIncnodes=docsearchCartInc.documentElement.selectNodes("/languages/language")
function getsearchCartIncLngStr(instring)
	temp = selectedsearchCartIncnode.selectSingleNode(instring).text
	getsearchCartIncLngStr = temp
end function
%>
