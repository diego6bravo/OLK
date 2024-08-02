<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "verfyQuery.xml"
set docverfyQuery = server.CreateObject("MSXML2.DOMDocument")
docverfyQuery.async = False
DocverfyQuery.Load(server.MapPath(xmlfilename)) 
docverfyQuery.setProperty "SelectionLanguage", "XPath"
set selectedverfyQuerynode = docverfyQuery.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedverfyQuerynodes=docverfyQuery.documentElement.selectNodes("/languages/language")
function getverfyQueryLngStr(instring)
	temp = selectedverfyQuerynode.selectSingleNode(instring).text
	getverfyQueryLngStr = temp
end function
%>
