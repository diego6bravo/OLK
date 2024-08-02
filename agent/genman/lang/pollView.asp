<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "pollView.xml"
set docpollView = server.CreateObject("MSXML2.DOMDocument")
docpollView.async = False
DocpollView.Load(server.MapPath(xmlfilename)) 
docpollView.setProperty "SelectionLanguage", "XPath"
set selectedpollViewnode = docpollView.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedpollViewnodes=docpollView.documentElement.selectNodes("/languages/language")
function getpollViewLngStr(instring)
	temp = selectedpollViewnode.selectSingleNode(instring).text
	getpollViewLngStr = temp
end function
%>
