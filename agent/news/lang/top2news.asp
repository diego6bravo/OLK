<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "top2news.xml"
set doctop2news = server.CreateObject("MSXML2.DOMDocument")
doctop2news.async = False
Doctop2news.Load(server.MapPath(xmlfilename)) 
doctop2news.setProperty "SelectionLanguage", "XPath"
set selectedtop2newsnode = doctop2news.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedtop2newsnodes=doctop2news.documentElement.selectNodes("/languages/language")
function gettop2newsLngStr(instring)
	temp = selectedtop2newsnode.selectSingleNode(instring).text
	gettop2newsLngStr = temp
end function
%>
