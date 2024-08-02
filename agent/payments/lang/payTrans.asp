<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "payTrans.xml"
set docpayTrans = server.CreateObject("MSXML2.DOMDocument")
docpayTrans.async = False
DocpayTrans.Load(server.MapPath(xmlfilename)) 
docpayTrans.setProperty "SelectionLanguage", "XPath"
set selectedpayTransnode = docpayTrans.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedpayTransnodes=docpayTrans.documentElement.selectNodes("/languages/language")
function getpayTransLngStr(instring)
	temp = selectedpayTransnode.selectSingleNode(instring).text
	getpayTransLngStr = temp
end function
%>
