<%@ Language=VBScript %>
<!--#include file="adminutils.asp"-->
<%
	Response.Expires = -1000 'Makes the browser not cache this page
	Response.Buffer = True 'Buffers the content so our Response.Redirect will work
	If (Session("LoggedIn") <> Session("PassWord")) Or (Session("LoggedIn") = "") Then
		Response.Redirect("default.asp")
	End If


	Dim sSQL
	Dim oRS
	'Dim sConnStr
	Dim Q
	Dim s
	
	Q = Chr(34)

	sSQL = "SELECT Description, DatabaseProductID, Title , ImageFile, Publish, ProductCode, Size, Price, Discontinued FROM qAdminAuditReport"

%>
<html>
<head>
	<title>The Hive Honey Shop: Web Site Administration</title>
	<LINK rel="stylesheet" href="hive.css" type="text/css">
</head>
<body>
	<!--#include file="header.asp"-->
	<h3>Database Product Audit Report</h3>
<%

	Set oRS=Server.CreateObject("ADODB.Recordset")
	oRS.Open sSQL, gsConnStr
	If Not oRS.BOF And Not oRS.EOF Then %>
		<TABLE BORDER="1" CELLPADDING="0" CELLSPACING="0">
			<tr>
				<td><i>Category</i></td>
				<td><i>Prod ID</i></td>
				<td width="45%" ><i>Product Title </i></td>
				<td><i>Publi- shed?</i></td>
				<td><i>Product Code</i></td>
				<td><i>Size</i></td>
				<td><i>Price</i></td>
				<td><i>Discon- tinued</i></td>
			</tr>

		<%While Not oRS.EOF
			%>
			<tr >
			<td style="font-size:8pt"><%=oRS("Description")%>&nbsp;</td>
			<td style="font-size:8pt"><%=oRS("DatabaseProductID")%>&nbsp;</td>
			<td style="font-size:8pt" width="45%"><%=oRS("Title")%>
			<%
			Dim filefullresolved
			Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
			If oRS("ImageFile")<>"" then
'				filefullresolved = Server.MapPath("..") & "\product_images\" & oRS("ImageFile")
'				If objFSO.FileExists(filefullresolved)=False Then 
'					RESPONSE.WRITE "<BR>LARGE IMAGE <input type=""checkbox"" checked=""1"">"
'				End If
				
				RESPONSE.WRITE  "<BR><input type=""checkbox"""
				filefullresolved = Server.MapPath("..") & "\product_images\" & oRS("ImageFile")
				If objFSO.FileExists(filefullresolved)=True THEN
				RESPONSE.WRITE " checked=""1"""
				End IF
				RESPONSE.WRITE ">" & oRS("ImageFile")

				RESPONSE.WRITE "<input type=""checkbox"""
				filefullresolved = Server.MapPath("..") & "\product_images\thumbs\t_" & oRS("ImageFile")
				If objFSO.FileExists(filefullresolved)=True THEN
				RESPONSE.WRITE " checked=""1"""
				End IF
				RESPONSE.WRITE ">THUMB t_..." '& " t_" & oRS("ImageFile")
				
				

			End if
			Set objFSO = nothing
			%>
			
			&nbsp;</td>
			<td style="font-size:8pt"><% If oRS("Publish")="False" Then response.write "No"%>&nbsp;</td>
			<td style="font-size:8pt"><%=oRS("ProductCode")%>&nbsp;</td>
			<td style="font-size:8pt"><%=oRS("Size")%>&nbsp;</td>
			<td align="right" style="font-size:8pt;background-color:<% If oRS("Price")=0 Then response.write "red"%>;"><%= FormatNumber(oRS("Price"),2)%>&nbsp;</td>
			<td style="font-size:8pt"><% If oRS("Discontinued")="True" Then response.write "Yes" %>&nbsp;</td>
			</tr>
			<%
			oRS.MoveNext
		Wend
		%>
		</TABLE>
	<p>
<a href="logout.asp">Log Out</a>&nbsp;<a href="mainpage.asp">Back</a> &nbsp;
	</p>
		<%
	End If
	oRS.Close
	Set oRS = Nothing
%>
</body>
</html>
