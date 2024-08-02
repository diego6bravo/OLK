<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "topGetValueSelect.xml"
set doctopGetValueSelect = server.CreateObject("MSXML2.DOMDocument")
doctopGetValueSelect.async = False
DoctopGetValueSelect.Load(server.MapPath(xmlfilename)) 
doctopGetValueSelect.setProperty "SelectionLanguage", "XPath"
set selectedtopGetValueSelectnode = doctopGetValueSelect.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedtopGetValueSelectnodes=doctopGetValueSelect.documentElement.selectNodes("/languages/language")
function gettopGetValueSelectLngStr(instring)
	temp = selectedtopGetValueSelectnode.selectSingleNode(instring).text
	gettopGetValueSelectLngStr = temp
end function
%>
