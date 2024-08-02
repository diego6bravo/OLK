<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "cxcDocDetail.xml"
set doccxcDocDetail = server.CreateObject("MSXML2.DOMDocument")
doccxcDocDetail.async = False
DoccxcDocDetail.Load(server.MapPath(xmlfilename)) 
doccxcDocDetail.setProperty "SelectionLanguage", "XPath"
set selectedcxcDocDetailnode = doccxcDocDetail.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedcxcDocDetailnodes=doccxcDocDetail.documentElement.selectNodes("/languages/language")
function getcxcDocDetailLngStr(instring)
	temp = selectedcxcDocDetailnode.selectSingleNode(instring).text
	getcxcDocDetailLngStr = temp
end function
%>
