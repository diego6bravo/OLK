<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "newsEdit.xml"
set docnewsEdit = server.CreateObject("MSXML2.DOMDocument")
docnewsEdit.async = False
DocnewsEdit.Load(server.MapPath(xmlfilename)) 
docnewsEdit.setProperty "SelectionLanguage", "XPath"
set selectednewsEditnode = docnewsEdit.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectednewsEditnodes=docnewsEdit.documentElement.selectNodes("/languages/language")
function getnewsEditLngStr(instring)
	temp = selectednewsEditnode.selectSingleNode(instring).text
	getnewsEditLngStr = temp
end function
%>
