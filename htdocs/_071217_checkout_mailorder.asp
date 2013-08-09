<%@ Language=VBScript%>
<!--#include file="include/javascript.asp"-->
<!-- #include file="hivepaypalconfig.asp" -->

<%
	Response.Expires =-1000
	Const CODE = 0
	Const DESC = 1
	Const PRICE = 2
	Const GIFTWRAP = 3
	Const QTY = 4
	Const PP = 6.95
	Const GIFTPRICE = 2.5
	
	Dim aBasket ' as array
	Dim sAction ' as string
	Dim sItemCode ' as string
	Dim sItemDesc ' as string
	Dim cUnitPrice	'as currency
	Dim iQuantity ' as integer
	Dim cTotalPrice 'as currency
	Dim cGiftWrap ' as currency
	Dim cSubTotal 
	Dim i
	Dim aUpdate
	Dim Q
	Dim chkGift
	Dim valGift
	Q = Chr(34)
	
	Function GetCategory()
		If IsEmpty(Session("category")) Then
			Session("category") = 1
		End If
		GetCategory = Session("category")
	End Function
	'===============================================================
	Sub AddItemToBasket(ItemCode, ItemDesc, UnitPrice, Quantity, Gift)
		Dim iIndex
		Dim iNewElement
		iIndex = ItemExists(ItemCode)
		If (iIndex >=1) Then
			aBasket(QTY, iIndex) = aBasket(QTY, iIndex) + Quantity
		Else			
			If (Not IsArray(aBasket)) Then
				iNewElement = 1
				Redim aBasket(4, iNewElement)
			Else
				iNewElement = ArrayUBound(aBasket,2) + 1			
				If iNewElement = 0 Then 'this is an empty array
					iNewElement = 1
				End If
				Redim Preserve aBasket(4, iNewElement)
			End If
			aBasket(CODE, iNewElement) = ItemCode
			aBasket(DESC, iNewElement) = ItemDesc
			aBasket(PRICE, iNewElement) = UnitPrice
			aBasket(GIFTWRAP, iNewElement) = Gift
			aBasket(QTY, iNewElement) = Quantity
		End If
		'Response.Write "<P>Ubound: " & ubound(aBasket,2)
	End Sub
	'===============================================================
	Sub RemoveItemFromBasket(ItemCode, Quantity)
		Dim iIndex
		Dim iNewElement
		iIndex = ItemExists(ItemCode)
		If (iIndex >=1) Then			
			If aBasket(QTY, iIndex) > 0 Then
				If (Quantity < 0 ) Then
					'remove all 
					aBasket(QTY, iIndex) = 0
				Else
					aBasket(QTY, iIndex) = aBasket(QTY, iIndex) - Quantity
				End If
			End If
			'Response.Write "<P>New Qty: " & aBasket(QTY, iIndex)
		End If
 	End Sub
	'===============================================================
	Sub ShowBasket
		Dim bEmptyBasket
		%>
	<!--#include file="../private/db/connection.asp"-->
	<html>
	<head>
		<title>The Hive Honey Shop: Mail Order Confirmation</title>
	<link href="hive.css" rel="stylesheet" type="text/css" />
	</head>
	<!--#include file="include/body.htm"-->
	<script language="JavaScript1.2">fwLoadMenus();</script>
	<div align="center">

	<a name="top">

	<!--contact details-->
	<!--#include file="include/contact.htm"-->
	
	<!--masthead-->
	<!--#include file="include/masthead.htm"-->

	<table style="border:solid;border-width:1px" cellpadding="5" cellspacing="0" width="760" bordercolor=#6D1746 bordercolorlight=#6D1746 bordercolordark=#6D1746>
	<tr>
	<td>
		<font face="verdana,arial,sans-serif" color="#6D1746">
			<h3>Mail Order Checkout</h3>
		</font>
	</td></tr>
	<tr><td>
	<% 'Mail Order Top
		Dim sSQL 
		sSQL = "SELECT SiteText From SiteText WHERE ID=5"
		Set oRS = Server.CreateObject("ADODB.Recordset")
		oRS.Open sSQL, gsConnStr
		response.Write oRS("SiteText")
		oRS.Close 
%>
	

						<TABLE border="1" cellspacing="0" cellpadding="5" bordercolor=#999999 width="750">
						<TR><TD><I><font face="verdana,arial,sans-serif" size="2" color="#000000">Code</font></I></TD>
						<TD><I><font face="verdana,arial,sans-serif" size="2" color="#000000">Description</font></I></TD>
						<TD><I><font face="verdana,arial,sans-serif" size="2" color="#000000">Quantity</font></I></TD>
						<TD><I><font face="verdana,arial,sans-serif" size="2" color="#000000">Gift Wrap?</font><BR><font face="verdana,arial,sans-serif" size="1" color="#000000">(£2.50 per item)</font></I></TD>
						<TD><I><font face="verdana,arial,sans-serif" size="2" color="#000000">Unit Price</font></I></TD><TD><I><font face="verdana,arial,sans-serif" size="2" color="#000000">Total Price</font></I></TD><TR>

				<!-- PayPal Configuration --> 
				

				<%
							
			If ArrayUBound(aBasket,2) > -1 Then
				'check for empty basket
				bEmptyBasket = True
				For i = 1 To ArrayUbound(aBasket,2)	
					If (aBasket(QTY, i) > 0) Then
						bEmptyBasket = False
						Exit For
					End If
				Next
				If (bEmptyBasket=False) Then
					%>

<%
					For i = 1 To ArrayUbound(aBasket,2)	
						If aBasket(QTY, i) > 0 Then	
							If aBasket(GIFTWRAP, i) > 0 Then	
								chkGift = "Yes"
								chkGiftDesc = " (Giftwrapped)"
							Else
								chkGift = "No"
								chkGiftDesc = ""
							End If
							cTotalPrice = CInt(aBasket(QTY, i)) * (CCur(aBasket(PRICE, i)) + CCur(aBasket(GIFTWRAP, i)))
							
Response.Write "<TR><TD><font face=""verdana,arial,sans-serif"" size=""2"" color=""#000000"">" & aBasket(CODE, i) & "&nbsp;<input name=""item_number_" & i &""" type=""hidden"" value="""& aBasket(CODE, i) &"""></font></TD>" & _
		"<TD><font face=""verdana,arial,sans-serif"" size=""2"" color=""#000000"">" & aBasket(DESC, i) & chkGiftDesc & "</font><input name=""item_name_" & i &""" type=""hidden"" value="""& aBasket(DESC, i) & chkGiftDesc &"""></TD>" & _
		"<TD align=center class=shop><font face=""verdana,arial,sans-serif"" size=""2"" color=""#000000"">" & aBasket(QTY, i) & " </font></TD>" & _
		"<TD class=shop>" &  chkGift & "</TD>" & _
		"<TD align=right><font face=""verdana,arial,sans-serif"" size=""2"" color=""#000000"">" & FormatCurrency(aBasket(PRICE, i),2) & "</font><input name=""amount_" & i &""" type=""hidden"" value="""& aBasket(PRICE, i) &"""></TD>" & _
		"<TD align=right><font face=""verdana,arial,sans-serif"" size=""2"" color=""#000000"">" & FormatCurrency(cTotalPrice,2) & "</font></TD></TR>" & vbcrlf & vbcrlf
							cSubTotal = cSubTotal + cTotalPrice
						End If
					Next	
					Response.Write "<TR><TD colspan=""5"" align=right class=shop><font face=""verdana,arial,sans-serif"" size=""2"" color=""#000000"">SubTotal</font></TD><TD align=right><font face=""verdana,arial,sans-serif"" size=""2"" color=""#000000"">" & FormatCurrency(cSubTotal,2) & "</font></TD></TR>" & vbcrlf
					Response.Write "<TR><TD colspan=""5"" align=right class=shop><font face=""verdana,arial,sans-serif"" size=""2"" color=""#000000"">Postage & Packing</font></TD><TD align=right><font face=""verdana,arial,sans-serif"" size=""2"" color=""#000000"">" & FormatCurrency(PP,2) & "</font></TD></TR>" & vbcrlf
					Response.Write "<TR><TD colspan=""5"" align=right class=shop><B><font face=""verdana,arial,sans-serif"" size=""2"" color=""#000000"">Total</font></B></TD><TD align=right class=shop><B><font face=""verdana,arial,sans-serif"" size=""2"" color=""#000000"">" & FormatCurrency(cSubTotal + PP,2) & "</font></B></TD></TR>" & vbcrlf
					Response.Write "</TABLE>"
				Else					
					%><TR><TD align=middle colspan="6" height="100"><font face="verdana,arial,sans-serif" size="2" color="#000000"><i>Your basket is empty</i></font></TD></TR></TABLE><%

				End If								
		Else
			bEmptyBasket = True
					%><TR><TD align=middle class=shop colspan="6" height="100"><font face="verdana,arial,sans-serif" size="2" color="#000000"><i>Your basket is empty</i></TD></TR></TABLE><%
		End If
			%>	<BR>
		</tr></td>
		<tr><td>	
	<table border="1" cellspacing="0" cellpadding="6" width="740" height="10" bgcolor=#cccccc>
					<TR>
		
				<TD valign=bottom width="25%" align=center><A href="orderform.asp" target="_blank"><B><font face="verdana,arial,sans-serif" size="2">Show Printable Mail Order Form</font></B></A></TD>
				</table>


	<% 'Mail Order Bottom
		sSQL = "SELECT SiteText From SiteText WHERE ID=6"
		Set oRS = Server.CreateObject("ADODB.Recordset")
		oRS.Open sSQL, gsConnStr
		response.Write oRS("SiteText")
		oRS.Close 
%>

		<img src="images/2bees.gif" align="right">

			</div>
			</body>
		</html>
		<%	
	End Sub	
	'===============================================================
	'Returns IIF equivilent results
	Function Iif(condition,test, value1, value2 )
	If condition=test Then Iif=Value1 Else Iif=value2
	End Function
	'===============================================================
	'Returns index of element if exists
	'Else returns -1
	Function ItemExists(ItemCode)
		ItemExists = -1		
		If ArrayUbound(aBasket,2)< 0 Then Exit Function
		For i = 1 To ArrayUbound(aBasket, 2)
			If (ItemCode = aBasket(CODE,i)) Then
				ItemExists = i
				Exit For
			End If
		Next
	End Function
	'===============================================================
	Function ArrayUBound(v, d)
		On Error Resume Next
		Dim u
		u = uBound(v,d)
		If Err <> 0 Then
			ArrayUbound = -1
			ExitFunction
		Else
			ArrayUbound = u
		End If
	End Function
	'===============================================================
	' Get a reference to the cart if it exists otherwise create it
	If Not IsArray(aBasket) Then
		Redim aBasket(4,0)
	End If


	'Response.Write "<P>Session: " & IsArray(Session("basket"))
	
	If IsEmpty(Session("basket"))  Then
		' change bahviour - ifempty basket then go back to shopping basket page
		response.Redirect "hive_shopping.asp" 
	'	Session("basket") = aBasket
	Else
		aBasket = Session("basket")
	End If
	
	sAction = CStr(Request.Form("action"))
	
	Select Case sAction
	Case "add"
		' Get all the parameters passed to the script
		sItemCode = CStr(Request.Form("item"))
		iQuantity = CInt(Request.Form("quantity"))
		sItemDesc  = CStr(Request.Form("itemdesc"))
		cUnitPrice = CCur(Request.Form("uprice"))
		cGiftWrap = 0
		'Response.Write "<P>" & iQuantity
		AddItemToBasket sItemCode, sItemDesc, cUnitPrice, iQuantity, cGiftWrap
		ShowBasket
	Case "del"
		' Get all the parameters passed to the script
		sItemCode = CStr(Request.Form("item"))
		iQuantity = CInt(Request.Form("quantity"))
		RemoveItemFromBasket sItemCode, iQuantity
		ShowBasket
	Case "empty"
		Erase aBasket
		ShowBasket
	Case Else
		ShowBasket
	End Select	

	Session("basket") = aBasket
%>








</body>
</html> 