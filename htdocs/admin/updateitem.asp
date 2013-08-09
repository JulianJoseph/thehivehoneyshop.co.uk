<!--#include file="adminutils.asp"-->
<%
	If Session("LoggedIn") <> Session("PassWord") Then
		Response.Redirect("default.asp")
	End If
	Response.Buffer = true 
	Dim sSQL
	Dim Q
	Dim oConn
	Dim sConnStr
	Dim i
	Dim productdetails
	Dim disc
	Dim bPublish
	
	Q = Chr(39)
	Response.Write "<p>" & Request.Form
	
	sSQL = "UPDATE Products " & _
			" SET Title = " & Q & FixSpecialChars(Request.Form("title")) & Q  & ", " & vbCrLf & _ 
			" Description = " & Q & FixSpecialChars(Request.Form("description")) & Q  & ", " & vbCrLf & _ 
			" CategoryID = " & Trim(Request.Form("categoryid"))  & ", " & vbCrLf & _ 
			" ImageFile = " & Q & Trim(Request.Form("imagefile")) & Q & ", "
			If Request.Form("publish") = "on" Then
				bPublish = "True"
			Else
				bPublish = "False"
			End If
	sSQL = sSQL & 	" Publish = " & bPublish & vbCrLf & _
			" WHERE ProductID = " & Request.Form("productid")
	Response.Write "<p>" & sSQL
	
	Set oConn = Server.CreateObject("ADODB.Connection")
	oConn.Open gsConnStr
	oConn.Execute sSQL
	
	productdetails = Split(Request.Form("productdetailids"),",")
	for i = 0 to ubound(productdetails)
		if request.Form("delete" & productdetails(i)) = "on" then
			sSQL = "DELETE * FROM ProductDetail WHERE ProductDetailID = " & productdetails(i)	
		else
			sSQL = "UPDATE ProductDetail " & _
				" SET Price = " & Request.Form("price" & productdetails(i)) & ", " & vbCrLf & _
				" ProductCode = " & Q & Request.Form("productcode" & productdetails(i)) & Q & ", " & vbCrLf & _
				" Size = " & Q & FixSpecialChars(Request.Form("size" & productdetails(i))) & Q & ", "
				If Request.Form("discontinued" & productdetails(i)) = "on" Then
					disc = "True"
				Else
					disc="False"
				End If
			sSQL = sSQL & " Discontinued = " & disc & vbCrLf & _
				" WHERE ProductDetailID = " & productdetails(i)	
		end if
		oConn.Execute sSQL
		Response.Write "<p>" & sSQL
	next
	
	'process new products
	if Request.Form("productcode_n") <> "" then
		If Request.Form("discontinued_n") = "on" Then
			disc = "True"
		Else
			disc = "False"
		End If
		sSQL = "INSERT INTO ProductDetail (ProductCode, Price, Size, Discontinued, ProductID) " & vbcrlf & _
				"VALUES(" & Q & Request.Form("productcode_n") & Q & ", " & _
				Q & Request.Form("price_n") & Q & ", " & _ 
				Q & Request.Form("size_n") & Q & ", " & _
				disc & ", " & _
				Request.Form("productid") & ")"
				response.Write "<P>" & sSQL				
				oConn.Execute sSQL
	end if
	
	
	oConn.Close 
	Set oConn = Nothing
	Response.Write "<p>" & i & " records updated"
	Response.Redirect "update_catalogue.asp?productid=" & Request.Form("productid")
	
%>