<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "setCartS.xml"
set docsetCartS = server.CreateObject("MSXML2.DOMDocument")
docsetCartS.async = False
DocsetCartS.Load(server.MapPath(xmlfilename)) 
docsetCartS.setProperty "SelectionLanguage", "XPath"
set selectedsetCartSnode = docsetCartS.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedsetCartSnodes=docsetCartS.documentElement.selectNodes("/languages/language")
function getsetCartSLngStr(instring)
	temp = selectedsetCartSnode.selectSingleNode(instring).text
	getsetCartSLngStr = temp
end function
%>
