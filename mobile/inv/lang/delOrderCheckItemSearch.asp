<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "delOrderCheckItemSearch.xml"
set docdelOrderCheckItemSearch = server.CreateObject("MSXML2.DOMDocument")
docdelOrderCheckItemSearch.async = False
DocdelOrderCheckItemSearch.Load(server.MapPath(xmlfilename)) 
docdelOrderCheckItemSearch.setProperty "SelectionLanguage", "XPath"
set selecteddelOrderCheckItemSearchnode = docdelOrderCheckItemSearch.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selecteddelOrderCheckItemSearchnodes=docdelOrderCheckItemSearch.documentElement.selectNodes("/languages/language")
function getdelOrderCheckItemSearchLngStr(instring)
	temp = selecteddelOrderCheckItemSearchnode.selectSingleNode(instring).text
	getdelOrderCheckItemSearchLngStr = temp
end function
%>
