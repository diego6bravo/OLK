<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "cxcRctDetail.xml"
set doccxcRctDetail = server.CreateObject("MSXML2.DOMDocument")
doccxcRctDetail.async = False
DoccxcRctDetail.Load(server.MapPath(xmlfilename)) 
doccxcRctDetail.setProperty "SelectionLanguage", "XPath"
set selectedcxcRctDetailnode = doccxcRctDetail.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedcxcRctDetailnodes=doccxcRctDetail.documentElement.selectNodes("/languages/language")
function getcxcRctDetailLngStr(instring)
	temp = selectedcxcRctDetailnode.selectSingleNode(instring).text
	getcxcRctDetailLngStr = temp
end function
%>
