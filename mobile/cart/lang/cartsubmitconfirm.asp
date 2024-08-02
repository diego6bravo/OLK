<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "cartsubmitconfirm.xml"
set doccartsubmitconfirm = server.CreateObject("MSXML2.DOMDocument")
doccartsubmitconfirm.async = False
Doccartsubmitconfirm.Load(server.MapPath(xmlfilename)) 
doccartsubmitconfirm.setProperty "SelectionLanguage", "XPath"
set selectedcartsubmitconfirmnode = doccartsubmitconfirm.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedcartsubmitconfirmnodes=doccartsubmitconfirm.documentElement.selectNodes("/languages/language")
function getcartsubmitconfirmLngStr(instring)
	temp = selectedcartsubmitconfirmnode.selectSingleNode(instring).text
	getcartsubmitconfirmLngStr = temp
end function
%>
