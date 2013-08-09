<%@ Language=VBScript%>
<!--#include file="../private/db/connection.asp"-->
<%
	Dim iItemID
	Dim s
	Dim sTitle
	Dim sDesc
	Dim i
	Dim n
	Dim cPrice
	Dim sPrice
	Dim sProdDesc
	Dim sItemDesc
	Dim sPictureCode
	Dim sCode
	Dim rs
	Dim sSQL
	Dim sImageFile
	Dim Q
	Dim winAttribs
	
	Q = Chr(34)
	iItemID = Trim(Request.QueryString("itemid"))
	sSQL = "SELECT * FROM qProductDetail WHERE ProductID = " & iItemID
	Set rs = Server.CreateObject("ADODB.RecordSet")
	
	rs.Open sSQL, gsConnStr
	
	If Not (rs.BOF) And (Not rs.EOF) Then
		sTitle = rs("Title")
		'sPictureCode = rs("PictureCode")
		sProdDesc = rs("Title")
		sDesc = rs("Description")
	End If
	
	'If (sPictureCode <> "") Then
	'	sTitle = sTitle & "&nbsp;(" & LCase(sPictureCode) & ")"
	'End If
	
	If rs("Description") <> "" Then
		If InStr(1, rs("Description"), vbCrLf) > 0 Then
			sDesc = Replace(rs("Description"), vbCrLf, "<br/><br/>")
		End If
	End If
	
	If rs("ImageFile") <> "" Then
		sImageFile = "product_images/" & rs("ImageFile")				
	End If
	
	'winAttribs = "'menubar=no,toolbar=no,resizable=no,status=no,titlebar=no,scrollbars=yes,location=no,height=400,width=400'"
	'winAttribs = "'menubar=no,toolbar=no,resizable=no,status=no,titlebar=no,scrollbars=yes,location=no'"
	winAttribs = "'resizable=0'"
	
%>

<html>
<head>
	<title>The Hive Honey Shop - Product Detail</title>	
</head>
<body bgcolor="#FFFFCC">
    <% =gConnStr %>
	<font face="verdana,arial,sans-serif" size="2" color="#000000">
	<table cellspacing="0" cellpadding="0" border="0" width="100%">
	<%
		s = "<TR><TD colspan=3 valign=top><font face=""verdana,arial,sans-serif"" color=""#6D1746""><H3>" & sTitle & "</H3></font></TD></TR>"

		If sImageFile <> "" Then
			s = s & "<TR><TD colspan=""3"" align=center><IMG style=""cursor:hand"" title=""Click to enlarge"" onclick=" & Q & "window.open('full_image.asp?" & sImageFile & "',''," & winAttribs & ");" & Q & " src=" & Q & sImageFile & Q & " vspace=""15"" height=""250"" align=center></TD></TR>"
			s = s & "<tr><td colspan=""3"" valign=top align=center><i><font color=""#6D1746"" size=""1"">Click on image to enlarge</font><i></td></tr>"

			's = s & "<TR><TD colspan=""3"" align=center><IMG  src=" & Q & sImageFile & Q & " vspace=""15"" align=center height=""250""></TD></TR>"
		End If

		s = s & "<TR><TD colspan=""3"" valign=top><font face=""verdana,arial,sans-serif"" size=""2"" color=""#000000""><P>" & sDesc & "</P></font></TD></TR>"
		s = s & "<TR><TD colspan=""3""><HR noshade color=#CCCCCC size=""1""></TD></TR>"

		Do While Not rs.EOF 
			n = n + 1						
			sItemDesc = sProdDesc & "&nbsp;" & rs("Size")
			If IsNumeric(rs("Price")) Then
				cPrice = CCur(rs("Price"))
				sPrice = FormatCurrency(rs("Price"),2)
			End If
			If LCase(rs("Discontinued")) = "true" Then
				sPrice = "Discontinued"
			End If
			
			sCode = rs("ProductCode")
			s = s & "<TR><TD valign=bottom><P><font size=""1"">" & sCode & "</font></P></TD>" & vbCrLf
			s = s & "<TD valign=bottom><P><font size=""1"">" & rs("Size") & "</font></P></TD>"  & vbCrLf
			s = s & "<TD valign=bottom><P style=""text-align:right""><font size=""1"">" & sPrice & "</font></P></TD></TR>" & vbCrLf
			If (LCase(rs("Discontinued")) = "false") Then
				s = s & "<FORM action=""hive_shopping.asp#basket"" target=""shop_window"" method=post ID=frmBasket" & n & " Name=frmBasket" & n & "> " & vbCrLf
				s = s & "<TR><TD colspan=""2"" align=right>"
				's = s & "<A href=""javascript:frmBasket" & n & ".submit();window.close();"">Add to Basket</A></TD><TD align=right>" & vbCrLf
				s = s & "<A href=""#"" onClick=""document.forms['frmBasket" & n & "'].submit();window.close();return false;""><font size=""1"">Add to Basket</font></A></TD><TD align=right>" & vbCrLf
				s = s & "<P><font size=""1"">Quantity:</font><input type=""text"" name=""quantity"" size=""3"" value=""1"" ID=intProductQty></P></TD></TR>" & vbCrLf
				s = s & "<INPUT type=""hidden"" name=""action"" value=""add"">"  & vbCrLf
				s = s & "<INPUT type=""hidden"" name=""item"" value=""" & sCode & """>"  & vbCrLf
				s = s & "<INPUT type=""hidden"" name=""itemdesc"" value=""" & sItemDesc & """>"  & vbCrLf
				s = s & "<INPUT type=""hidden"" name=""uprice"" value=""" & cPrice & """>"  & vbCrLf
				s = s & "</FORM>" & vbCrLf
			End If
			s = s & "<TR><TD colspan=""3""><HR noshade color=#CCCCCC size=""1""></TD></TR>" & vbCrLf
			
			rs.MoveNext 
		Loop
		rs.Close 
		Set rs = Nothing
		
		Response.Write s
	%>
		<TR><TD colspan="3"><P>&nbsp;</TD></TR>
		<TR><TD colspan="3"><P align=center><I><font size="1"><A href="javascript:window.close();">Close Window</A></font></I></P></TD></TR>	
		<TR><TD colspan="3"><P>&nbsp;</TD></TR>
		<TR><TD colspan="3" align=center><IMG src="images/2bees.gif"></TD></TR>
		<TR><TD colspan="3"><P>&nbsp;</TD></TR>
	</table>
	</font>

</BODY>
</HTML>



