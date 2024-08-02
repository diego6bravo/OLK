<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "cards.xml"
set doccards = server.CreateObject("MSXML2.DOMDocument")
doccards.async = False
Doccards.Load(server.MapPath(xmlfilename)) 
doccards.setProperty "SelectionLanguage", "XPath"
set selectedcardsnode = doccards.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedcardsnodes=doccards.documentElement.selectNodes("/languages/language")
function getcardsLngStr(instring)
	temp = selectedcardsnode.selectSingleNode(instring).text
	getcardsLngStr = temp
end function
%>
