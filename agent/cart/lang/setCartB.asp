<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "setCartB.xml"
set docsetCartB = server.CreateObject("MSXML2.DOMDocument")
docsetCartB.async = False
DocsetCartB.Load(server.MapPath(xmlfilename)) 
docsetCartB.setProperty "SelectionLanguage", "XPath"
set selectedsetCartBnode = docsetCartB.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedsetCartBnodes=docsetCartB.documentElement.selectNodes("/languages/language")
function getsetCartBLngStr(instring)
	temp = selectedsetCartBnode.selectSingleNode(instring).text
	getsetCartBLngStr = temp
end function
%>
