<% If Session("VendId") = "" Then response.redirect "default.asp" %>
<!--#include file="myHTMLEncode.asp"-->
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus %>
<!--#include file="pdfAccess.asp"-->
<%

makeAction = ""
For each itm in Request.Form
	If itm <> "newLng" Then
		makeAction = makeAction & "&" & itm & "=" & Server.URLEncode(Request(itm))
	End If
Next
makeAction = 	"searchCartPDF.asp?dbID=" & Session("id") & "&myRnd=" & myRnd & "&excell=Y&olkimgpath=" & Session("olkimgpath") & _
				"&branch=" & Session("branch") & "&vendid=" & Session("vendid") & "&pdf=Y&newLng=" & Session("myLng") & _
				makeAction
Dim theDoc

Set theDoc = Server.CreateObject("ABCpdf6.Doc")

theScale = 0.8
theDoc.Rect.Magnify 1 / theScale, 1 / theScale
theDoc.Transform.Magnify theScale, theScale, 0, 0
theDoc.Rect.Inset 20 / theScale, 20 / theScale
theDoc.VPos = 0.5
theDoc.HtmlOptions.Timeout = 180000
varx = Request.ServerVariables("URL")
	theID = theDoc.AddImage(GetHTTPStr & Request.ServerVariables("SERVER_NAME") & Mid(varx,1,Len(varx)-11) & makeAction, 1)
	Do
	  If theDoc.GetInfo(theID, "Truncated") <> "1" Then Exit Do
	  theDoc.Page = theDoc.AddPage()
	  theID = theDoc.AddImage(theID)
	Loop
	For i = 1 To theDoc.PageCount
	  theDoc.PageNumber = i
	  theDoc.Flatten
	Next
theData = theDoc.GetData()
Response.ContentType = "application/pdf"
Response.AddHeader "content-disposition", "inline; filename=searchCart.pdf"
Response.BinaryWrite theData
set theDoc = nothing

%>