<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="chkLogin.asp" -->
<!--#include file="myHTMLEncode.asp"-->
<!--#include file="pdfAccess.asp"-->
<%
makeAction = 	"cxcDocDetailPrint.asp?CardCode=" & Request("CardCode") & "&DocEntry=" & Request("DocEntry") & "&isEntry=" & Request("isEntry") & "&pop=Y&DocType=" & Request("DocType") & _
				"&dbID=" & Session("ID") & "&myRnd=" & myRnd & "&UserType=" & userType & "&useraccess=" & Session("useraccess") & "&vendid=" & Session("vendid") & "&pdf=Y&newLng=" & Session("myLng")

Dim theDoc

Set theDoc = Server.CreateObject("ABCpdf6.Doc")

theScale = 0.8
theDoc.Rect.Magnify 1 / theScale, 1 / theScale
theDoc.Transform.Magnify theScale, theScale, 0, 0
theDoc.Rect.Inset 20 / theScale, 20 / theScale
theDoc.VPos = 0.5
varx = Request.ServerVariables("URL")

	theID = theDoc.AddImageUrl(GetHTTPStr & Request.ServerVariables("SERVER_NAME") & Mid(varx,1,Len(varx)-19) & makeAction, 1,,true)
	Do
	  If theDoc.GetInfo(theID, "Truncated") <> "1" Then Exit Do
	  theDoc.Page = theDoc.AddPage()
	  theID = theDoc.AddImage(theID)
	Loop
	For i = 1 To theDoc.PageCount
	  theDoc.PageNumber = i
	  theDoc.Flatten
	Next
'theDoc.Save Server.MapPath(Mid(varx,1,Len(varx)-23) & "temp/") & "\cxcDocDetail.pdf" 
theData = theDoc.GetData()
Response.ContentType = "application/pdf"
'Response.AddHeader "content-length", UBound(theData) - LBound(theData) + 1
Response.AddHeader "content-disposition", "inline; filename=cxcDocDetail.pdf"
Response.BinaryWrite theData


set theDoc = nothing
%>
