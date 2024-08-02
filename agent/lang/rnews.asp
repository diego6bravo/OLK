<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "rnews.xml"
set docrnews = server.CreateObject("MSXML2.DOMDocument")
docrnews.async = False
Docrnews.Load(server.MapPath(xmlfilename)) 
docrnews.setProperty "SelectionLanguage", "XPath"
set selectedrnewsnode = docrnews.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedrnewsnodes=docrnews.documentElement.selectNodes("/languages/language")
function getrnewsLngStr(instring)
	temp = selectedrnewsnode.selectSingleNode(instring).text
	getrnewsLngStr = temp
end function
%>
