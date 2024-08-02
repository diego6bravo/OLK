<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "cartBreakDown.xml"
set doccartBreakDown = server.CreateObject("MSXML2.DOMDocument")
doccartBreakDown.async = False
DoccartBreakDown.Load(server.MapPath(xmlfilename)) 
doccartBreakDown.setProperty "SelectionLanguage", "XPath"
set selectedcartBreakDownnode = doccartBreakDown.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedcartBreakDownnodes=doccartBreakDown.documentElement.selectNodes("/languages/language")
function getcartBreakDownLngStr(instring)
	temp = selectedcartBreakDownnode.selectSingleNode(instring).text
	getcartBreakDownLngStr = temp
end function
%>
