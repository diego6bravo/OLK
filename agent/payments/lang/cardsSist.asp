<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "cardsSist.xml"
set doccardsSist = server.CreateObject("MSXML2.DOMDocument")
doccardsSist.async = False
DoccardsSist.Load(server.MapPath(xmlfilename)) 
doccardsSist.setProperty "SelectionLanguage", "XPath"
set selectedcardsSistnode = doccardsSist.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedcardsSistnodes=doccardsSist.documentElement.selectNodes("/languages/language")
function getcardsSistLngStr(instring)
	temp = selectedcardsSistnode.selectSingleNode(instring).text
	getcardsSistLngStr = temp
end function
%>
