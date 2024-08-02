<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "processTrans.xml"
set docprocessTrans = server.CreateObject("MSXML2.DOMDocument")
docprocessTrans.async = False
DocprocessTrans.Load(server.MapPath(xmlfilename)) 
docprocessTrans.setProperty "SelectionLanguage", "XPath"
set selectedprocessTransnode = docprocessTrans.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedprocessTransnodes=docprocessTrans.documentElement.selectNodes("/languages/language")
function getprocessTransLngStr(instring)
	temp = selectedprocessTransnode.selectSingleNode(instring).text
	getprocessTransLngStr = temp
end function
%>
