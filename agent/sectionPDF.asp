<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="myHTMLEncode.asp"-->
<!--#include file="pdfAccess.asp"-->
<%
On Error Resume Next
makeAction = 	"sectionCleanPDF.asp?dbID=" & Session("id") & "&myRnd=" & myRnd & _
				"&UserType=" & userType & "&branch=" & Session("branch") & "&vendid=" & Session("vendid") & "&UserName=" & Session("username") & "&pdf=Y&newLng=" & Session("myLng")

For each itm in Request.QueryString
	makeAction = makeAction & "&" & itm & "=" & Request(itm)
Next
Dim theDoc

Set theDoc = Server.CreateObject("ABCpdf6.Doc")

theScale = 0.8
theDoc.Rect.Magnify 1 / theScale, 1 / theScale
theDoc.Transform.Magnify theScale, theScale, 0, 0
theDoc.Rect.Inset 20 / theScale, 20 / theScale
theDoc.VPos = 0.5
varx = Request.ServerVariables("URL")

theID = theDoc.AddImageUrl(GetHTTPStr & Request.ServerVariables("SERVER_NAME") & Mid(varx,1,Len(varx)-14) & makeAction, 1,,true)
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
Response.AddHeader "content-disposition", "inline; filename=cartPdf.pdf"
Response.BinaryWrite theData

set theDoc = nothing

%>