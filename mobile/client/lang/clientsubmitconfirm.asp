<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "clientsubmitconfirm.xml"
set docclientsubmitconfirm = server.CreateObject("MSXML2.DOMDocument")
docclientsubmitconfirm.async = False
Docclientsubmitconfirm.Load(server.MapPath(xmlfilename)) 
docclientsubmitconfirm.setProperty "SelectionLanguage", "XPath"
set selectedclientsubmitconfirmnode = docclientsubmitconfirm.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedclientsubmitconfirmnodes=docclientsubmitconfirm.documentElement.selectNodes("/languages/language")
function getclientsubmitconfirmLngStr(instring)
	temp = selectedclientsubmitconfirmnode.selectSingleNode(instring).text
	getclientsubmitconfirmLngStr = temp
end function
%>
