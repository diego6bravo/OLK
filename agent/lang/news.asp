<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "news.xml"
set docnews = server.CreateObject("MSXML2.DOMDocument")
docnews.async = False
Docnews.Load(server.MapPath(xmlfilename)) 
docnews.setProperty "SelectionLanguage", "XPath"
set selectednewsnode = docnews.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectednewsnodes=docnews.documentElement.selectNodes("/languages/language")
function getnewsLngStr(instring)
	temp = selectednewsnode.selectSingleNode(instring).text
	getnewsLngStr = temp
end function
%>
