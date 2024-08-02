<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "activityConfDetail.xml"
set docactivityConfDetail = server.CreateObject("MSXML2.DOMDocument")
docactivityConfDetail.async = False
DocactivityConfDetail.Load(server.MapPath(xmlfilename)) 
docactivityConfDetail.setProperty "SelectionLanguage", "XPath"
set selectedactivityConfDetailnode = docactivityConfDetail.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedactivityConfDetailnodes=docactivityConfDetail.documentElement.selectNodes("/languages/language")
function getactivityConfDetailLngStr(instring)
	temp = selectedactivityConfDetailnode.selectSingleNode(instring).text
	getactivityConfDetailLngStr = temp
end function
%>
