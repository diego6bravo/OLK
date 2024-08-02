<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "addCard.xml"
set docaddCard = server.CreateObject("MSXML2.DOMDocument")
docaddCard.async = False
DocaddCard.Load(server.MapPath(xmlfilename)) 
docaddCard.setProperty "SelectionLanguage", "XPath"
set selectedaddCardnode = docaddCard.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedaddCardnodes=docaddCard.documentElement.selectNodes("/languages/language")
function getaddCardLngStr(instring)
	temp = selectedaddCardnode.selectSingleNode(instring).text
	getaddCardLngStr = temp
end function
%>
