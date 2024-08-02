<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "invact.xml"
set docinvact = server.CreateObject("MSXML2.DOMDocument")
docinvact.async = False
Docinvact.Load(server.MapPath(xmlfilename)) 
docinvact.setProperty "SelectionLanguage", "XPath"
set selectedinvactnode = docinvact.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedinvactnodes=docinvact.documentElement.selectNodes("/languages/language")
function getinvactLngStr(instring)
	temp = selectedinvactnode.selectSingleNode(instring).text
	getinvactLngStr = temp
end function
%>
