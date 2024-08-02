<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "ImageList.xml"
set docImageList = server.CreateObject("MSXML2.DOMDocument")
docImageList.async = False
DocImageList.Load(server.MapPath(xmlfilename)) 
docImageList.setProperty "SelectionLanguage", "XPath"
set selectedImageListnode = docImageList.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedImageListnodes=docImageList.documentElement.selectNodes("/languages/language")
function getImageListLngStr(instring)
	temp = selectedImageListnode.selectSingleNode(instring).text
	getImageListLngStr = temp
end function
%>
