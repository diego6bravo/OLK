<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "topGetDocLink.xml"
set doctopGetDocLink = server.CreateObject("MSXML2.DOMDocument")
doctopGetDocLink.async = False
DoctopGetDocLink.Load(server.MapPath(xmlfilename)) 
doctopGetDocLink.setProperty "SelectionLanguage", "XPath"
set selectedtopGetDocLinknode = doctopGetDocLink.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedtopGetDocLinknodes=doctopGetDocLink.documentElement.selectNodes("/languages/language")
function gettopGetDocLinkLngStr(instring)
	temp = selectedtopGetDocLinknode.selectSingleNode(instring).text
	gettopGetDocLinkLngStr = temp
end function
%>
