<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "itmConfDetail.xml"
set docitmConfDetail = server.CreateObject("MSXML2.DOMDocument")
docitmConfDetail.async = False
DocitmConfDetail.Load(server.MapPath(xmlfilename)) 
docitmConfDetail.setProperty "SelectionLanguage", "XPath"
set selecteditmConfDetailnode = docitmConfDetail.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selecteditmConfDetailnodes=docitmConfDetail.documentElement.selectNodes("/languages/language")
function getitmConfDetailLngStr(instring)
	temp = selecteditmConfDetailnode.selectSingleNode(instring).text
	getitmConfDetailLngStr = temp
end function
%>
