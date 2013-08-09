<%@ Language=VBScript %>
<!--#include file="adminutils.asp"-->
<%
	Response.Expires = -1000 'Makes the browser not cache this page
	Response.Buffer = True 'Buffers the content so our Response.Redirect will work
	If (Session("LoggedIn") <> Session("PassWord")) Or (Session("LoggedIn") = "") Then
		Response.Redirect("default.asp")
	End If
		'On Error Resume Next
		Response.Expires = -1000
		Dim oRS
		Dim sSQL
		Dim fo
		Dim sConnStr
		Dim oFile
		Dim sPath
		Dim Q
		Dim fld
		Dim cats()
		Dim i
		Dim s
		Dim productdetails
		Dim iCat
		Dim sCat
		
		sSQL = "SELECT CategoryID, Description FROM Category ORDER BY CategoryID"
		Set oRS=Server.CreateObject("ADODB.Recordset")
		oRS.Open sSQL, gsConnStr
		If Not oRS.BOF And Not oRS.EOF Then
			Do While Not oRS.EOF
				Redim Preserve cats(1,i)
				cats(0,i)=oRS.Fields("CategoryID")
				cats(1,i)=oRS.Fields("Description")
				oRS.MoveNext 
				i = i + 1
			Loop
		End If
		oRS.Close
		Set oRS=Nothing

		If Request.QueryString("categoryid") <> "" Then
			Session("CategoryID") = Request.QueryString("categoryid")
		End If
		iCat = Session("CategoryID")


		sSQL = "SELECT * FROM Products WHERE ProductID = " & Request.QueryString("productid")
		Set oRS=Server.CreateObject("ADODB.Recordset")		
		oRS.Open sSQL, gsConnStr, adOpenDynamic, adLockOptimistic
		'======================================================================	
		If Not oRS.BOF And Not oRS.EOF Then
			oRS.MoveFirst 
			%>
		<html>
		<head>
			<TITLE>The Hive Honey Shop: Web Site Administration</TITLE>
			<LINK rel="stylesheet" href="hive.css" type="text/css">
			<script language=javascript>
			<!--
				function copyFileName(f)
				{	
					//get full path and file name, split into array and return last element
					var a = f.split("\\");
					var name = a[a.length-1];								
					window.frmUpdateItem.imagefile.value=name;
				}
			//-->
			</script>
			
		</head>
		<body>
			<!--#include file="header.asp"-->
			<h2><%=oRS.Fields("Title")%></h2>			
			<form id=frmUpdateItem name=frmUpdateItem action="updateitem.asp" method="post">
			<input type="hidden" value="<%=Request.QueryString("productid")%>" name=productid id=productid>
			<TABLE>
				<TR><TD>Title</TD><TD><INPUT type=text value="<%=oRS.Fields("Title")%>" size="<%=oRS.Fields("Title").DefinedSize%>" id=title name=title></TD></TR>
				<tr>
					<td>Publish?</td>
					<td><input id="publish" value="on" name="publish" type="checkbox" <%If oRS("Publish") Then %> CHECKED <% End If%>></td>
				</tr>

		<%	'Image information
			Dim sProdID
			sProdID = Request.QueryString("productid")
			sImageFileMain = oRS("ImageFile")
			If sImageFileMain <> "" Then 
					sImageFileThumb = "t_" & oRS("ImageFile")
			else	sImageFileThumb = ""
			End if
		%>

				<TR><TD>Description</TD><TD><TEXTAREA id=description name=description cols="75" rows="15"><%=oRS.Fields("Description")%></TEXTAREA></TD></TR>
				<TR><TD>Category</TD><TD>
						<SELECT name=categoryid id=categoryid size="1">
						<%
							For i = 0 To Ubound(cats,2)
								s = "<OPTION VALUE=" & Q
								s = s & cats(0,i) & Q 
								If cats(0,i)=oRS.Fields("CategoryID") Then
									s = s & " SELECTED "								
								End If
								s = s & ">"
								s = s & cats(1,i)
								Response.Write vbTab & vbTab & s & vbcrlf								
							Next
						%>
						</SELECT></TD></TR>						
				<TR><TD>Image File</TD><TD><INPUT type=text value="<%=oRS.Fields("ImageFile")%>" size="<%=oRS.Fields("ImageFile").DefinedSize%>" id="imagefile" name=imagefile style="background:#ccc"> (For reference only. DO NOT edit this field)</TD></TR>
			</TABLE>			
			<%
			
		End If 
		'======================================================================	
		oRS.Close
		Set oRS=Nothing

		sSQL = "SELECT ProductCode, ProductID, Price, Size, Discontinued, ProductDetailID " & _
				"FROM ProductDetail WHERE ProductID = " & Request.QueryString("productid")

		Set oRS=Server.CreateObject("ADODB.Recordset")		
		oRS.Open sSQL, gsConnStr, adOpenDynamic, adLockOptimistic
		If Not oRS.BOF And Not oRS.EOF Then
			oRS.MoveFirst 
			%>
				<table  border="1" cellspacing="0" cellpadding="1">
					<tr><td><i>Product Code</i></td><td><i>Price</i></td><td><i>Size</i></td><td><i>Discontinued</i></td><td><i>Delete?</i></td></tr>
			<%
			While Not oRS.EOF
				i = oRS.Fields("ProductDetailID")
				productdetails = productdetails & oRS.Fields("ProductDetailID") & ","
				%>
				<tr>
					<td><input type="hidden" value="<%=oRS.Fields("ProductDetailID")%>" id=productdetailid<%=i%> name=productdetailid<%=i%>>
					<input type="text" value="<%=oRS.Fields("ProductCode")%>" size="<%=oRS.Fields("ProductCode").DefinedSize%>" id=productcode<%=i%> name=productcode<%=i%>></td>
					<td><input type="text" value="<%=oRS.Fields("Price")%>" size="<%=oRS.Fields("Price").DefinedSize%>" id=price<%=i%> name=price<%=i%>></td>
					<td><input type="text" value="<%=oRS.Fields("Size")%>" size="<%=oRS.Fields("Size").DefinedSize%>" id=size<%=i%> name=size<%=i%>></td>
					<td><input type="checkbox" <%If oRS.Fields("Discontinued") Then %>CHECKED<%End If%> id=discontinued<%=i%> name=discontinued<%=i%>></td>
					<td><input type="checkbox" id=delete<%=i%> name=delete<%=i%>></td>					
				</tr>
				<%
				oRS.MoveNext
			Wend
			%>
				<tr><td><input type="hidden" value="" id=productdetailid_n name=productdetailid_n>
					<input type="text" value="" size="<%=oRS.Fields("ProductCode").DefinedSize%>" id=productcode_n name=productcode_n></td>
					<td><input type="text" value="" size="<%=oRS.Fields("Price").DefinedSize%>" id=price_n name=price_n></td>
					<td><input type="text" value="" size="<%=oRS.Fields("Size").DefinedSize%>" id=size_n name=size_n></td>
					<td><input type="checkbox" id="discontinued_n" name=discontinued_n></td><td>&nbsp;</td> </tr>
			<%			
				productdetails = mid(productdetails, 1, len(productdetails)-1)
			%>
				</table>
				<input type="hidden" value="<%=productdetails%>" id=productdetailids name=productdetailids>

			<%
			oRS.Close
			Set oRS=Nothing
		End If

%>
<input type="submit" value="Save">
<input type="reset" value="Cancel">
<input type="button" value="Close" onclick="window.location.href='list_products.asp?categoryid=<%=iCat%>';">
</form>


<%
' Check if the Main and Thuymb images exist
Dim sImageFileMainExists, sImageFileThumbExists
Dim filefullresolved
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
If sImageFileMain <> "" then
	filefullresolved = Server.MapPath("..") & "\product_images\" & sImageFileMain 
	sImageFileMainExists=objFSO.FileExists(filefullresolved)
	filefullresolved = Server.MapPath("..") & "\product_images\thumbs\" & sImageFileThumb
	sImageFileThumbExists=objFSO.FileExists(filefullresolved)
Else
	sImageFileMainExists = False
	sImageFileThumbExists = false
End if
Set objFSO = nothing
%>



<TABLE BORDER="1"  cellspacing="0" cellpadding="1">

	<TR>
	<TD>Main Image<br /><img height="200" border="1" align="left" src="../product_images/<%=sImageFileMain%>"></TD>
	<TD>Filename: <%=sImageFileMain%><br />
		<% If sImageFileMainExists=False Then
				IF	sImageFileMain<>"" then 
					response.write "IMAGE FILE MISSING ON WEBSERVER!"
				Else
					response.write "NO IMAGE FILE UPLOADED. Once you upload this image, this will set the filename." 
				End if
			End IF
		%>
		<br />					

		<FORM METHOD="POST" ENCTYPE="multipart/form-data" ACTION="upload_image.asp?productid=<%=sProdID%>" ID="Form1">
			<TABLE BORDER="0" ID="Table1" cellpadding="0" cellspacing="0">
				<tr>
					<td><br>Select the main large image file to upload:</td></tr>
				<tr><td>	
						<INPUT TYPE="FILE" SIZE="50" NAME="txtImageFile" ID="txtImageFile" value="<%=sImageFileMain%>" onchange="copyFileName(this.value);" >
						<INPUT TYPE="SUBMIT" VALUE="Upload" ID="Submit1" NAME="Submit1"></td>
						<!--<INPUT type="hidden" name="frmCalling" id="frmCalling" value="update_catalogue.asp?productid=<%=Request.QueryString("productid")%>&categoryid=<%=Request.QueryString("categoryid")%>">-->
				</tr>
			</TABLE>
		</FORM>
	</TD></TR>
	<TR><TD>Thumbnail Image<BR /><img height="50" border="1" align="left" src="../product_images/thumbs/<%=sImageFileThumb%>"></TD>
	<%If sImageFileMain <> "" And sImageFileMainExists = True Then%>
	<TD>
		Thumbnail Filename: <%=sImageFileThumb%><BR />
		<% If sImageFileThumbExists=False then 
		response.write "IMAGE FILE MISSING ON WEBSERVER!" 
		End if%><br />
		<FORM METHOD="POST" ENCTYPE="multipart/form-data" ACTION="upload_image_thumb.asp?productid=<%=sProdID%>" ID="Form1">
			<TABLE BORDER="0" ID="Table1" cellpadding="0" cellspacing="0">
				<tr>
					<td>The thumbnail image MUST be named <%=sImageFileThumb%>, (it is based upon the main image filename).<br>
					The height should be exactly 50 pixels (but the horizontal length with be proportional to the original image dimensions).<br>
					Select a thumb nail image file to upload:</td></tr>
				<tr><td>	
						<INPUT TYPE="FILE" SIZE="50" NAME="txtImageFile" ID="txtImageFile" accept="image/jpeg" >
						<INPUT TYPE="SUBMIT" VALUE="Upload" ID="Submit1" NAME="Submit1"></td>
						<!--<INPUT type="hidden" name="frmCalling" id="frmCalling" value="update_catalogue.asp?productid=<%=Request.QueryString("productid")%>&categoryid=<%=Request.QueryString("categoryid")%>">-->
				</tr>
			</TABLE>
		</FORM>
		</TD></TR>
	<%Else%>
		<TD>Please upload your main image before you upload the thumbnail.</TD></TR>				
	<%End If%>


</TABLE>

<p><a href="logout.asp">Log Out</a></p>
</body>
</html>


