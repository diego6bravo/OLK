<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "setCart3.xml"
set docsetCart3 = server.CreateObject("MSXML2.DOMDocument")
docsetCart3.async = False
DocsetCart3.Load(server.MapPath(xmlfilename)) 
docsetCart3.setProperty "SelectionLanguage", "XPath"
set selectedsetCart3node = docsetCart3.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedsetCart3nodes=docsetCart3.documentElement.selectNodes("/languages/language")
function getsetCart3LngStr(instring)
	temp = selectedsetCart3node.selectSingleNode(instring).text
	getsetCart3LngStr = temp
end function
%>
