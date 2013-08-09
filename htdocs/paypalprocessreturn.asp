<%@ Language=VBScript%>
<!--#include file="../private/db/connection.asp"-->
<html>
<head>
	<title>The Hive Honey Shop : PayPal Order Completed</title>
	<!--#include file="include/javascript.asp"-->
	<link href="hive.css" rel="stylesheet" type="text/css" />
</head>
<!--#include file="include/body.htm"-->
<script language="JavaScript1.2">fwLoadMenus();</script>
<div align="center">

	<!--contact details-->
	<!--#include file="include/contact.htm"-->
	
	<!--masthead-->
	<!--#include file="include/masthead.htm"-->

	<table style="border:solid;border-width:1px" cellpadding="5" cellspacing="0" width="760" bordercolor=#6D1746 bordercolorlight=#6D1746 bordercolordark=#6D1746>
	<tr>
		<td rowspan="2" valign="top">
			<table>
			<tr>
			 <td><img src="images/spacer.gif" width="103" height="1" border="0"></td>
			 <td><img src="images/spacer.gif" width="1" height="1" border="0"></td>
			</tr>
			<tr>
			 <td><img name="history2_a" src="images/history2_a.jpg" width="103" height="139" border="0"></td>
			 <td><img src="images/spacer.gif" width="1" height="139" border="0"></td>
			</tr>
			<tr>
			 <td><img name="history2_b" src="images/history2_b.jpg" width="103" height="91" border="0"></td>
			 <td><img src="images/spacer.gif" width="1" height="91" border="0"></td>
			</tr>
			<tr>
			 <td><img name="history2_c" src="images/history2_c.jpg" width="103" height="76" border="0"></td>
			 <td><img src="images/spacer.gif" width="1" height="76" border="0"></td>
			</tr>
			<tr>
			 <td><img name="history2_d" src="images/history2_d.jpg" width="103" height="89" border="0"></td>
			 <td><img src="images/spacer.gif" width="1" height="89" border="0"></td>
			</tr>	
			</table>	
		</td>
	</tr>
	<tr>
		<td align="left">
			<font face="verdana,arial,sans-serif" color="#6D1746">
			<h3>Thank you</h3>
			</font>			
			


			<% 'response.Write( CStr(Request("amt"))) %>
					<%' PayPal Return Thankyou
				sSQL = "SELECT SiteText From SiteText WHERE ID=7"
				Set oRS = Server.CreateObject("ADODB.Recordset")
				oRS.Open sSQL, gsConnStr
				response.Write oRS("SiteText")
				oRS.Close 
				Session("basket") = empty
				%>

		</td>
	</tr>
	<tr>
		<td colspan="2" align="right"><img src="images/2bees.gif"></td>
	</tr>
	</table>
</div>
</body>
</html>


