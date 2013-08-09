<%@ Language=VBScript%>
<%
	On Error Resume Next
	Const CODE = 0
	Const DESC = 1
	Const PRICE = 2
	Const GIFTWRAP = 3
	Const QTY = 4
	Const PP = 4.95
	Const GIFTPRICE = 2.5
	Dim aBasket


	If IsEmpty(Session("basket")) Then
		Redim aBasket(4,5)
	Else
		aBasket = Session("basket")
	End If

	If ArrayUBound(aBasket,2) < 0 Then
		Redim aBasket(4,5)	
	End If
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
	Function FormatPrice(price)
		If price > 0 Then
			FormatPrice = FormatCurrency(price,2)
		Else
			FormatPrice = ""
		End If
	End Function
	'===============================================================
	Function FormatTotalPrice(price,pp)
		If price > 0 Then
			FormatTotalPrice = FormatCurrency(price + pp,2)
		Else
			FormatTotalPrice = ""
		End If
	End Function
	'===============================================================

%>
<HTML>
<HEAD>
		<TITLE>The Hive Honey Shop: Order Form</TITLE>
		<STYLE>
		<!--
			BODY {color:#000000;font-family:verdana;font-size:7pt}
			P {color:#000000;font-family:verdana;font-size:7pt}
			TD {color:#000000;font-family:verdana;font-size:7pt}
			.heading {color:#000000;font-family:verdana;font-size:12pt;font-weight:bold}
			.textbox {border-width:1px;border-style:solid;border-collapse:collapse;height:22px;padding-left:5px;padding-top:2px;font-weight:bold;font-size:7pt}
			.itembox {border-width:1px;border-style:solid;border-collapse:collapse;height:22px;padding-left:5px;padding-top:2px;font-size:7pt}
			.parabox {border-width:1px;border-style:solid;border-collapse:collapse;height:22px;padding-left:5px;padding-top:2px;font-family:verdana;font-weight:bold;font-size:7pt;overflow:auto}
			TD.shop {color:#000000;font-family:verdana;font-size:7pt}
		-->
		</STYLE>

</HEAD>
<BODY  bottommargin=5 rightmargin=25 leftmargin=25 topmargin=0 onLoad="frmOrder.txtFullName.focus();">
<FONT size="1"><A href="javascript:window.close();">[Close Window]</A>&nbsp;<A href="javascript:window.print();">[Print Form]</A></FONT>
<FORM name=frmOrder id=frmOrder>
<DIV align=center>
<SPAN class=heading>ORDER FORM</SPAN>
<BR>The Hive Honey Shop&nbsp;&#149;&nbsp;93 Northcote Road&nbsp;&#149;&nbsp;London&nbsp;&#149;&nbsp;SW11&nbsp;6PL&nbsp;&#149; Tel:&nbsp;+44&nbsp;(0)&nbsp;20&nbsp;7924&nbsp;6233&nbsp;&#149;&nbsp;Fax:&nbsp;+44&nbsp;(0)&nbsp;20&nbsp;7228&nbsp;7176&nbsp;&#149;&nbsp;email:&nbsp;order@thehivehoneyshop.co.uk
<BR>BLOCK CAPITALS PLEASE<br><br>
</DIV>
	<TABLE cellpadding="0" cellspacing="0" width="100%" bordercolor=#000000 bordercolorlight=#000000 bordercolordark=#000000>
			<TR><TD colspan="2"><I>Billing Address</I></TD><TD colspan="2"><I>Delivery Address (if different)</I></TD></TR>
			<TR><TD>Full Name</TD><TD><INPUT name=txtFullName id=txtFullName class=textbox tabindex=1 type=text size=40></INPUT><TD>Full Name</TD><TD><INPUT class=textbox tabindex=7 type=text size=40 ></INPUT></TR>
			<TR><TD>Address</TD><TD><INPUT class=textbox tabindex=2 type=text size=40></INPUT><TD>Address</TD><TD><INPUT class=textbox tabindex=8 type=text size=40></INPUT></TR>
			<TR><TD>&nbsp;</TD><TD><INPUT class=textbox tabindex=3 type=text size=40></INPUT><TD>&nbsp;</TD><TD><INPUT class=textbox tabindex=9 type=text size=40></INPUT></TR>
			<TR><TD>Post Code</TD><TD><INPUT class=textbox tabindex=4 type=text size=40></INPUT><TD>Post Code</TD><TD><INPUT class=textbox tabindex=10 type=text size=40></INPUT></TR>
			<TR><TD>Email:</TD><TD><INPUT class=textbox tabindex=5 type=text size=40 id=text1 name=text1></INPUT><TD>Email:</TD><TD><INPUT class=textbox tabindex=11 type=text size=40 id=text2 name=text2></INPUT></TR>
			<TR><TD>Daytime Tel:</TD><TD><INPUT class=textbox tabindex=6 type=text size=40></INPUT><TD>Daytime Tel:</TD><TD><INPUT class=textbox tabindex=12 type=text size=40></INPUT></TR>
	</TABLE>
	<br>
<TABLE class=itembox border="1" cellspacing="0" cellpadding="1" width="100%" bordercolor=#000000 bordercolorlight=#000000 bordercolordark=#000000>
	<TR><TD class=shop><I>Code</I></TD><TD class=shop><I>Description</I></TD><TD class=shop><I>Quantity</I></TD>
	<TD class=shop><I>Gift Wrap?<BR><FONT size="1">(£2.50 per item)</FONT></I></TD>
	<TD class=shop><I>Unit Price</I></TD><TD class=shop><I>Total Price</I></TD><TR>
	
	<%		If ArrayUBound(aBasket,2) > -1 Then
				'check for empty basket
				bEmptyBasket = True
				For i = 1 To ArrayUbound(aBasket,2)	
					If (aBasket(QTY, i) > 0) Then
						bEmptyBasket = False
						Exit For
					End If
				Next
				If (bEmptyBasket=True) Then
					Redim aBasket(4,5)
					bEmptyBasket=False
					
				End If
				If (bEmptyBasket=False) Then
					For i = 1 To ArrayUbound(aBasket,2)	
						If aBasket(QTY, i) > 0 Then	
							If aBasket(GIFTWRAP, i) > 0 Then	
								chkGift = "YES"															
							Else
								chkGift = ""
							End If
							cTotalPrice = CInt(aBasket(QTY, i)) * (CCur(aBasket(PRICE, i)) + CCur(aBasket(GIFTWRAP, i)))
							
							Response.Write "<TR><TD class=shop>" & aBasket(CODE, i) & "&nbsp;</TD><TD class=shop>" & aBasket(DESC, i) & "</TD><TD align=center class=shop>" &  CStr(aBasket(QTY, i)) & "</TD>" & _
											"<TD class=shop>" &  chkGift  & "</TD>" & _
											"<TD align=right class=shop>" & FormatPrice(aBasket(PRICE, i)) & "</TD><TD align=right class=shop>" & FormatPrice(cTotalPrice) & "</TD></TR>" & vbcrlf & vbcrlf
							cSubTotal = cSubTotal + cTotalPrice
						Else 'print blank form
							Response.Write "<TR><TD class=shop>&nbsp;</TD><TD class=shop>&nbsp;</TD><TD align=center class=shop>&nbsp;</TD>" & _
											"<TD class=shop>&nbsp;</TD>" & _
											"<TD align=right class=shop>&nbsp;</TD><TD align=right class=shop>&nbsp;</TD></TR>" & vbcrlf & vbcrlf
						End If
					Next	
					Response.Write "<TR><TD colspan=""5"" align=right class=shop>SubTotal</TD><TD align=right class=shop>" & FormatPrice(cSubTotal) & "</TD></TR>" & vbcrlf
					Response.Write "<TR><TD colspan=""5"" align=right class=shop>Postage & Packing (UK only - for international delivery, please phone for details)</TD><TD align=right class=shop>" & FormatCurrency(PP,2) & "</TD></TR>" & vbcrlf
					Response.Write "<TR><TD colspan=""5"" align=right class=shop><B>Total</B></TD><TD align=right class=shop><B>" & FormatTotalPrice(cSubTotal,PP) & "</B></TD></TR>" & vbcrlf
			End If
		End If
	%>
	</TABLE>
	<p><b>P&P Info</b><br>
	Order as many products as you like and pay only one delivery charge of &pound4.95 per address. Applies to UK Mainland only.</p>
	<B>Method of Payment</B><br>
	<INPUT type=checkbox ID="Checkboxchq" NAME="Checkboxchq"></INPUT>Pay <% Response.Write FormatTotalPrice(cSubTotal,PP) %> by Cheque or Postal Order. (Payable to <B>The Hive</B>. Please write your name and address on the backs of any cheques).<br>
	<BR>
	<INPUT type=checkbox ID="Checkboxcc" NAME="Checkboxcc"></INPUT>Pay <% Response.Write FormatTotalPrice(cSubTotal,PP) %> by Credit Card, using the card details below<BR>
	<TABLE border="0" cellspacing="0" cellpadding="5" width="100%" bordercolor=#000000 bordercolorlight=#000000 bordercolordark=#000000>
	<TR><TD>Type of card</TD><TD><SELECT class=textbox type=select size=1 style="width:300px">
									<OPTION SELECTED>&nbsp;
									<OPTION>Visa
									<OPTION>Mastercard
									<OPTION>Visa Debit
									<OPTION>Mastercard Debit
									<OPTION>Switch
									<OPTION>Maestro
									<OPTION>Delta
									<OPTION>Solo 
								</SELECT></TD>
	<TD>Issue No (if applicable)</TD><TD><INPUT class=textbox type=text size=5></INPUT></TD>
								</TR>
	<TR><TD>Full name of cardholder</TD><TD><INPUT class=textbox type=text size=60></INPUT></TD>
	<TD>Start Date (if applicable)</TD><TD><INPUT class=textbox type=text size=15></INPUT></TD></TR>
	<TR><TD>Card Number</TD><TD><INPUT class=textbox type=text size=60></INPUT></TD>
	<TD>Expiry Date</TD><TD><INPUT class=textbox type=text size=15></INPUT></TD></TR>
	<TR><TD>Cardholder's signature</TD><TD><INPUT class=textbox type=text size=60></INPUT></TD>
	<td>Security Code<br></td>
	<td><INPUT class=textbox type=text size=10></INPUT></td></TR>	
	<TR><TD></TD><TD></TD><TD COLSPAN=2>(The Security Code is the <span style="text-decoration:underline">last</span> 3 numbers on the signature strip)</TD></TR>
	</TABLE>
	<HR noshade color=#cccccc size="1">
	<TABLE cellpadding="0" cellspacing="0" width="100%" bordercolor=#000000 bordercolorlight=#000000 bordercolordark=#000000>
			<TR><TD valign=top><I>Special Delivery Instructions</I></TD></TR>
			<TR><TD valign=top><TEXTAREA style="height:50px" type=text class=parabox cols="100" rows="4"></TEXTAREA></TD></TR>
	</TABLE>
	<BR>
	<TABLE cellpadding="0" cellspacing="0" width="100%" bordercolor=#000000 bordercolorlight=#000000 bordercolordark=#000000>
			<TR><TD valign=top><I>Gift Message</I></TD></TR>
			<TR><TD valign=top><TEXTAREA style="height:50px" type=text class=parabox cols="100" rows="4"></TEXTAREA></TD></TR>
	</TABLE>
	<br>
	Mead &amp; Honey Liqueurs contain alcohol and can only be purchased by those over 18 years of age. In order for us to dispatch, please tick the confirmation box. 
<INPUT type=checkbox ID="Checkbox1" NAME="Checkbox1"></INPUT>&nbsp;I am over 18 years old<br>
	Please tick this box if you do not wish to receive news and updates from The Hive Honey Shop <INPUT type=checkbox></INPUT>		
<BR>	
	<B>Thank you for your order!</B>
	<BR>If you are planning to come to London, please call into The Hive Honey Shop and say hello!
	<BR><br>	
	<TABLE class=itembox border="1" cellspacing="0" cellpadding="1" width="100%" bordercolor=#000000 bordercolorlight=#000000 bordercolordark=#000000>
			<TR>
				<TD rowspan="2"><B>Office Use Only</B></TD>
				<TD>Date Order Taken</TD>
				<TD>Taken By</TD>
				<TD>N/Processed</TD>
				<TD>Processed Date</TD>
				<TD>Packed By & Date</TD>
				<TD>Courier No. & Date</TD>
				<TD>Post Office No. & Date</TD>
			</TR>
			<TR height="30">
				<TD>&nbsp;</TD>
				<TD>&nbsp;</TD>
				<TD>&nbsp;</TD>
				<TD>&nbsp;</TD>
				<TD>&nbsp;</TD>
				<TD>&nbsp;</TD>
				<TD>&nbsp;</TD>
			</TR>
	</TABLE>
	
</FORM>	
</BODY>
</HTML>