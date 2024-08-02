<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "searchRecover.xml"
set docsearchRecover = server.CreateObject("MSXML2.DOMDocument")
docsearchRecover.async = False
DocsearchRecover.Load(server.MapPath(xmlfilename)) 
docsearchRecover.setProperty "SelectionLanguage", "XPath"
set selectedsearchRecovernode = docsearchRecover.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedsearchRecovernodes=docsearchRecover.documentElement.selectNodes("/languages/language")
function getsearchRecoverLngStr(instring)
	temp = selectedsearchRecovernode.selectSingleNode(instring).text
	getsearchRecoverLngStr = temp
end function
%>
