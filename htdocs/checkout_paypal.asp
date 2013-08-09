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
	Const PP = 4.95
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
		<title>The Hive Honey Shop: Online Order Confirmation</title>
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
			<h3>Online Checkout</h3>
		</font>
	</td></tr>
	<tr><td>
	<%
		Dim sSQL 
		sSQL = "SELECT SiteText From SiteText WHERE ID=3"
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
<form method="post" name="paypal_form" action="<%=paypal_url%>">
<input type="hidden" name="cmd" value="<%=paypal_cmd%>"> 
<input type="hidden" name="business" value="<%=paypal_business%>"> 
<input type="hidden" name="upload" value="1">
<input type="hidden" name="currency_code" value="<%=paypal_currency_code%>">
<input type="hidden" name="lc" value="<%=paypal_lc%>">
<input type="hidden" name="encrypted" value="-----BEGIN PKCS7-----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-----END PKCS7-----
">

<%
					For i = 1 To ArrayUbound(aBasket,2)	
						If aBasket(QTY, i) > 0 Then	
						'	If aBasket(GIFTWRAP, i) > 0 Then	
						'		chkGift = "Yes"
						'		chkGiftDesc = " (Giftwrapped)"
						'	Else
								chkGift = "No"
								chkGiftDesc = ""
						'	End If
							cTotalPrice = CInt(aBasket(QTY, i)) * (CCur(aBasket(PRICE, i))  )
							'CCur(aBasket(GIFTWRAP, i))
							
Response.Write "<TR><TD><font face=""verdana,arial,sans-serif"" size=""2"" color=""#000000"">" & aBasket(CODE, i) & "&nbsp;<input name=""item_number_" & i &""" type=""hidden"" value="""& aBasket(CODE, i) &"""></font></TD>" & _
		"<TD><font face=""verdana,arial,sans-serif"" size=""2"" color=""#000000"">" & aBasket(DESC, i) & chkGiftDesc & "</font><input name=""item_name_" & i &""" type=""hidden"" value="""& aBasket(DESC, i) & chkGiftDesc &"""></TD>" & _
		"<TD align=center class=shop><font face=""verdana,arial,sans-serif"" size=""2"" color=""#000000"">" & aBasket(QTY, i) & " </font><input name=""quantity_" & i &""" type=""hidden"" value="""& aBasket(QTY, i) &"""></TD>" & _
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


<%		
' Check for online site order status
	Dim sSiteOnlineEnabled
	sSiteOnlineEnabled = 0
	sSQL = "SELECT OnlineOrderEnabled FROM admin where id=1"
	Set oRS=Server.CreateObject("ADODB.Recordset")
	oRS.Open sSQL, gsConnStr
	If Not oRS.BOF And Not oRS.EOF Then	sSiteOnlineEnabled = oRS("OnlineOrderEnabled")
sSiteOnlineStatus = "[ Online Site Orders "
	If sSiteOnlineEnabled = True Then 
%>
	<%
		sSQL = "SELECT SiteText From SiteText WHERE ID=4"
		Set oRS = Server.CreateObject("ADODB.Recordset")
		oRS.Open sSQL, gsConnStr
		response.Write oRS("SiteText")
		oRS.Close 
%>

		<input type="image" align="right" ALT="Pay with Paypal" src="./paypal_x-click-but6.gif"/><br/>
		<%Else
			Response.write "<h2><b>Sorry, but the Paypal facility is currently offline for maintenance. Please use the Mail Order facility in the interim. Thankyou.</b></h2>"		
		End If
		%>



		
		
		
		
		</form> 

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