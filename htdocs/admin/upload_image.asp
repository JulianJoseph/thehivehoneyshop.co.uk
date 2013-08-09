<%@ Language=VBScript %>
<!--#include file="adminutils.asp"-->
<html>
	<head>
		<TITLE>The Hive Honey Shop: Web Site Administration</TITLE>
		<LINK rel="stylesheet" href="hive.css" type="text/css">
	</head>
<body>
<h2>Uploading Image File</h2>
<%
	Dim sSQL
	Dim oConn
	Dim Q
	Dim sFileName
	
	Q = Chr(39)

	Set Upload = Server.CreateObject("Persits.Upload")
	Count = Upload.Save(Server.MapPath("../product_images"))
	
	sFileName  = Upload.Files(1).FileName
	
	If Count = 1 Then
		sSQL = "UPDATE Products SET ImageFile = " & Q & sFileName & Q & " WHERE ProductID = " & Request.QueryString("productid")& ";" 
		'response.Write "<P>" & sSQL & "</p>"
		Set oConn = Server.CreateObject("ADODB.Connection")
		oConn.open gsConnStr
		oConn.execute sSQL
		oConn.close 
		Response.Write "<p>" & sFileName & " has been successfully uploaded to The Hive Honeyshop web site.<br>Press <i>Return to Item</i> to see a preview of the image.</p>"
		Response.Write "<p>Note that uploaded images should be sized with a width to height ratio of 1.5:1 for correct display.</p>"
	Else
		Response.Write "There was a problem uploadeding the file."
	End If
%>
<p><input type=button value="Return to Item" onclick="window.location.href='update_catalogue.asp?productid=<%=Request.Querystring("productid")%>';"></p>
<!--<p><input type=button value=Close onclick="history.back();" ID="btnBack" NAME="btnBack"></p>-->

</body> 
</html>