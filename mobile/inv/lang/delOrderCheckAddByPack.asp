<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "delOrderCheckAddByPack.xml"
set docdelOrderCheckAddByPack = server.CreateObject("MSXML2.DOMDocument")
docdelOrderCheckAddByPack.async = False
DocdelOrderCheckAddByPack.Load(server.MapPath(xmlfilename)) 
docdelOrderCheckAddByPack.setProperty "SelectionLanguage", "XPath"
set selecteddelOrderCheckAddByPacknode = docdelOrderCheckAddByPack.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selecteddelOrderCheckAddByPacknodes=docdelOrderCheckAddByPack.documentElement.selectNodes("/languages/language")
function getdelOrderCheckAddByPackLngStr(instring)
	temp = selecteddelOrderCheckAddByPacknode.selectSingleNode(instring).text
	getdelOrderCheckAddByPackLngStr = temp
end function
%>
