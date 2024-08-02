<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "soConfDetail.xml"
set docsoConfDetail = server.CreateObject("MSXML2.DOMDocument")
docsoConfDetail.async = False
DocsoConfDetail.Load(server.MapPath(xmlfilename)) 
docsoConfDetail.setProperty "SelectionLanguage", "XPath"
set selectedsoConfDetailnode = docsoConfDetail.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedsoConfDetailnodes=docsoConfDetail.documentElement.selectNodes("/languages/language")
function getsoConfDetailLngStr(instring)
	temp = selectedsoConfDetailnode.selectSingleNode(instring).text
	getsoConfDetailLngStr = temp
end function
%>
