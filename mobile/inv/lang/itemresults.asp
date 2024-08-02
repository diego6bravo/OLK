<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "itemresults.xml"
set docitemresults = server.CreateObject("MSXML2.DOMDocument")
docitemresults.async = False
Docitemresults.Load(server.MapPath(xmlfilename)) 
docitemresults.setProperty "SelectionLanguage", "XPath"
set selecteditemresultsnode = docitemresults.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selecteditemresultsnodes=docitemresults.documentElement.selectNodes("/languages/language")
function getitemresultsLngStr(instring)
	temp = selecteditemresultsnode.selectSingleNode(instring).text
	getitemresultsLngStr = temp
end function
%>
