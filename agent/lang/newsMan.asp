<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "newsMan.xml"
set docnewsMan = server.CreateObject("MSXML2.DOMDocument")
docnewsMan.async = False
DocnewsMan.Load(server.MapPath(xmlfilename)) 
docnewsMan.setProperty "SelectionLanguage", "XPath"
set selectednewsMannode = docnewsMan.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectednewsMannodes=docnewsMan.documentElement.selectNodes("/languages/language")
function getnewsManLngStr(instring)
	temp = selectednewsMannode.selectSingleNode(instring).text
	getnewsManLngStr = temp
end function
%>
