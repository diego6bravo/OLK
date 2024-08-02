<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "actdel.xml"
set docactdel = server.CreateObject("MSXML2.DOMDocument")
docactdel.async = False
Docactdel.Load(server.MapPath(xmlfilename)) 
docactdel.setProperty "SelectionLanguage", "XPath"
set selectedactdelnode = docactdel.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedactdelnodes=docactdel.documentElement.selectNodes("/languages/language")
function getactdelLngStr(instring)
	temp = selectedactdelnode.selectSingleNode(instring).text
	getactdelLngStr = temp
end function
%>
