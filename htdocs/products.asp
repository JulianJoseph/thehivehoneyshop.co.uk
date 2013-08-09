<%@ Language=VBScript%>
<!--#include file="../private/db/connection.asp"-->
<!--../../../private/clients/hivehoneyshop/connection.asp-->
<%
	Const adOpenStatic = 3
	const adLockReadOnly = 1
	Dim iCategory
	Dim sCat
	Dim s(1)
	Dim Q
	Dim sHead
	Dim sFoot
	'Dim i
	Dim iCount
	Dim e
	
	Dim rs
	Dim sSQL
	
	Q = Chr(34)
		
	iCategory = Trim(Request.QueryString("category"))
	Session("category")=iCategory
	
	Set rs = Server.CreateObject("ADODB.RecordSet")
	sSQL = "SELECT Description, HeaderText, FooterText FROM Category WHERE CategoryID = " & iCategory
	rs.Open sSQL, gsConnStr
	
	If (Not rs.BOF) And (Not rs.EOF) Then
'		sHead = "<p size=1 align=left><I>" & rs("HeaderText") & "</I></P>" ' JDE 12Aug06 strip format
'		sFoot = "<p size=1 align=left><I>" & rs("FooterText") & "</I></P>"
		sHead = rs("HeaderText") 
		sFoot = rs("FooterText")
		sCat = rs("Description")
		rs.Close 
	Else
		sHead = ""
		sFoot = ""
		sCat = "Category could not be found"
	End If
	
%>
<html>
<head>
	<title>The Hive Honey Shop : <%=sCat%></title>
	<link href="hive.css" rel="stylesheet" type="text/css" />
	<!--#include file="include/javascript.asp"-->	
	<script LANGUAGE=javascript>
		window.name = "shop_window";
			function popupwin(s)
			{
				var win = window.open("itemdetail.asp?itemid=" + s ,"_blank","height=550,width=500,location=no,menubar=no,resizable=no,status=no,toolbar=no,titlebar=no,scrollbars=yes");
			}						

	</script>
</head>
<!--#include file="include/body.htm"-->
<script language="JavaScript">fwLoadMenus();</script>
<div align="center">

	<!--contact details-->
	<!--#include file="include/contact.htm"-->
	
	<!--masthead-->
	<!--#include file="include/masthead.htm"-->

	<table style="border:solid;border-width:1px" cellpadding="5" cellspacing="0" width="760" bordercolor=#6D1746 bordercolorlight=#6D1746 bordercolordark=#6D1746>
	<tr>
		<td align=left valign=top colspan="2" rowspan="1">
			<font face="verdana,arial,sans-serif" color="#6D1746">
			<h3><%=sCat%></h3>
			</font>
						<p><font face="verdana,arial,sans-serif" size="2" color="#000000"><%=sHead%></font></p>

			<!-- font face="verdana,arial,sans-serif" size="2" color="#000000" -->
			<img src="images/beebutton.gif" border="0">&nbsp;<I>Please click on the item description to find out more about the product</I>
			<!-- hr noshade color=#CCCCCC size="1" -->
</td></tr>
<tr><td valign="top" align="left" width="380" >
			<%
				sSQL = "SELECT Title, ProductID, CategoryID, ProductPrice, Description, ProductCode, ImageFile " & _
						"FROM qProducts WHERE CategoryID = " & iCategory & " " & _
						"ORDER BY Title;"
				rs.Open sSQL, gsConnStr, adOpenStatic, adLockReadOnly
				Do while not rs.EOF 
					iCount = iCount + 1
'					If iCount > 15 Then
'						e = 1
'					End If
					' JDE 12 Aug 06 - Changed layout so that it alternates rather than produces a long list of 15 or less items in one column
					If (iCount Mod 2 = 1 ) then
						e=0
					Else 
						e=1
					End If
						

					's = ""
'					s(e) = s(e) & "<a href=" & Q & "javascript:popupwin('" & rs("ProductID") & "')" & Q & "><img align=texttop border=""0"" src=""images/beebutton.gif"" alt=""Click for details""></a>&nbsp;"
'					s(e) = s(e) & rs("Title") & "&nbsp;"								
'		JDE  -19 Feb 2006 - Change order of product list HREF
'					s(e) = s(e) & "<img align=texttop border=""0"" src=""images/beebutton.gif"" alt=""Click for details"">" 


' Removed font formatting and added css
s(e) = s(e) & "<table width=""100%"" cellpadding=3 style=""border:solid;border-width:1px;border-color:#000000""><tr><td>" '<font face=""verdana,arial,sans-serif"" size=""2"" color=""#000000"">"
					If rs("ImageFile") <> "" THEN
						s(e) = s(e) & "<img height=""50"" align=""right"" border=""1"" src=""./product_images/thumbs/t_" &  rs("ImageFile") & """  onclick=" & Q & "javascript:popupwin('" & rs("ProductID") & "');" & Q & " >"
					End if
					s(e) = s(e) & "<a href=" & Q & "javascript:popupwin('" & rs("ProductID") & "')" & Q & ">"
					s(e) = s(e) & rs("Title") & "</a>"

					If Trim(rs("ProductCode")) <> "" Then
'						s(e) = s(e) & "<br />(#" & rs("ProductCode") & ") "	 & rs("ProductPrice") & ""	
						s(e) = s(e) & "<br />"	 & rs("ProductPrice") & ""	
					End If


					s(e) = s(e) & "</td></tr></table>" '</font>
'					s(e) = s(e) & "<hr noshade color=#CCCCCC size=""1"">"
					s(e) = s(e) & "<br />"

					rs.MoveNext 
				Loop
				rs.Close 
				Set rs = Nothing
				Response.Write s(0)
'			</font>		moved up to comment it out!

			%>

		</td>
		<td valign="top" align="left" width="380" >
			<%=s(1)%>
			<p><font face="verdana,arial,sans-serif" size="1" color="#000000"><%=sFoot%></font></p>
		</td>
	</tr>
	<tr>
		<td align="left" valign="bottom"><font face="verdana,arial,sans-serif" size="1" color="#6D1746"><i>Site Last Updated: <%=Day(Now) & "-" & MonthName(Month(Now), True) & "-" & Year(Now)%></i></font></td>
		<td align="right"><img src="images/2bees.gif"></td>
	</tr>
	
	</table>
</div>
</body>
</html>



