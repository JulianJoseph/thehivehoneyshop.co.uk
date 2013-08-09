<%
	
	Response.Buffer = true
	Sub UpdateBasket
		Dim aQty
		Dim u
		Dim j
		Dim d
		Dim index
		Dim newqty
		Dim newgiftwrap
		Dim aBasket
		Dim lastelm
		Const NAME=0
		Const QUANTITY=1
		Const GIFTWRAP=1
		Const GFTWRP = 3
		Const QTY=4
		Dim r

		aBasket = Session("basket")
		aQty = Split(Request.Form, "&")

		u = UBound(aQty)
		For j = 0 to u
			d = split(aQty(j), "=")
			If Instr(1,LCase(d(NAME)),"quantity") Then
				index = CInt(Right(d(NAME), 1))				
				aBasket(GFTWRP, index) = 0
				newqty = CInt(d(QUANTITY))
				aBasket(QTY, index) = newqty
				Response.Write "<P>" & aBasket(DESC, index) & " - " & newqty
			ElseIf Instr(1,LCase(d(NAME)),"giftwrap") Then
				index = CInt(Right(d(NAME), 1))				
				'newgiftwrap = d(GIFTWRAP)
				aBasket(GFTWRP, index) = 2.5
				Response.Write "<P>" & aBasket(DESC, index) & " - " & newgiftwrap
			End If
		Next
		Session("basket") = aBasket
		'Response.Write "<P>" & Request.Form
	End Sub
	
	UpdateBasket
	r = Request.QueryString("redirect")
	Response.Write "<P>" & r
	Response.Redirect r

%>
