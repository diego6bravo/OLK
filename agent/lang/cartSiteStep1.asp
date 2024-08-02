<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "cartSiteStep1.xml"
set doccartSiteStep1 = server.CreateObject("MSXML2.DOMDocument")
doccartSiteStep1.async = False
DoccartSiteStep1.Load(server.MapPath(xmlfilename)) 
doccartSiteStep1.setProperty "SelectionLanguage", "XPath"
set selectedcartSiteStep1node = doccartSiteStep1.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedcartSiteStep1nodes=doccartSiteStep1.documentElement.selectNodes("/languages/language")
function getcartSiteStep1LngStr(instring)
	temp = selectedcartSiteStep1node.selectSingleNode(instring).text
	getcartSiteStep1LngStr = temp
end function
%>
