<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "itemDetailsData.js.xml"
set docitemDetailsDatajs = server.CreateObject("MSXML2.DOMDocument")
docitemDetailsDatajs.async = False
DocitemDetailsDatajs.Load(server.MapPath(xmlfilename)) 
docitemDetailsDatajs.setProperty "SelectionLanguage", "XPath"
set selecteditemDetailsDatajsnode = docitemDetailsDatajs.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selecteditemDetailsDatajsnodes=docitemDetailsDatajs.documentElement.selectNodes("/languages/language")
function getitemDetailsDatajsLngStr(instring)
	temp = selecteditemDetailsDatajsnode.selectSingleNode(instring).text
	getitemDetailsDatajsLngStr = temp
end function
%>
