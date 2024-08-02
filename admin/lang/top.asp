<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "top.xml"
set doctop = server.CreateObject("MSXML2.DOMDocument")
doctop.async = False
Doctop.Load(server.MapPath(xmlfilename)) 
doctop.setProperty "SelectionLanguage", "XPath"
set selectedtopnode = doctop.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedtopnodes=doctop.documentElement.selectNodes("/languages/language")
function gettopLngStr(instring)
	temp = selectedtopnode.selectSingleNode(instring).text
	gettopLngStr = temp
end function
%>
