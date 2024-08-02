<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "B1.xml"
set docB1 = server.CreateObject("MSXML2.DOMDocument")
docB1.async = False
DocB1.Load(server.MapPath(xmlfilename)) 
docB1.setProperty "SelectionLanguage", "XPath"
set selectedB1node = docB1.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedB1nodes=docB1.documentElement.selectNodes("/languages/language")
function getB1LngStr(instring)
	temp = selectedB1node.selectSingleNode(instring).text
	getB1LngStr = temp
end function
%>
