<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "crdConfDetail.xml"
set doccrdConfDetail = server.CreateObject("MSXML2.DOMDocument")
doccrdConfDetail.async = False
DoccrdConfDetail.Load(server.MapPath(xmlfilename)) 
doccrdConfDetail.setProperty "SelectionLanguage", "XPath"
set selectedcrdConfDetailnode = doccrdConfDetail.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedcrdConfDetailnodes=doccrdConfDetail.documentElement.selectNodes("/languages/language")
function getcrdConfDetailLngStr(instring)
	temp = selectedcrdConfDetailnode.selectSingleNode(instring).text
	getcrdConfDetailLngStr = temp
end function
%>
