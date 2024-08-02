<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "ofertsX.xml"
set docofertsX = server.CreateObject("MSXML2.DOMDocument")
docofertsX.async = False
DocofertsX.Load(server.MapPath(xmlfilename)) 
docofertsX.setProperty "SelectionLanguage", "XPath"
set selectedofertsXnode = docofertsX.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedofertsXnodes=docofertsX.documentElement.selectNodes("/languages/language")
function getofertsXLngStr(instring)
	temp = selectedofertsXnode.selectSingleNode(instring).text
	getofertsXLngStr = temp
end function
%>
