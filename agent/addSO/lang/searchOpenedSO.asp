<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "searchOpenedSO.xml"
set docsearchOpenedSO = server.CreateObject("MSXML2.DOMDocument")
docsearchOpenedSO.async = False
DocsearchOpenedSO.Load(server.MapPath(xmlfilename)) 
docsearchOpenedSO.setProperty "SelectionLanguage", "XPath"
set selectedsearchOpenedSOnode = docsearchOpenedSO.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedsearchOpenedSOnodes=docsearchOpenedSO.documentElement.selectNodes("/languages/language")
function getsearchOpenedSOLngStr(instring)
	temp = selectedsearchOpenedSOnode.selectSingleNode(instring).text
	getsearchOpenedSOLngStr = temp
end function
%>
