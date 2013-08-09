<!--#include file="adminutils.asp"-->
<%

	Response.Buffer = true 
	Dim sSQL
	Dim Q
	Dim oConn
	Dim i
	Dim rs
	Dim CatID
	Dim ProdID
	Dim Ref
	Dim bError
	
	
	Q = Chr(39)
	CatID = Request.Form("CategoryID")
	Ref = "mainpage.asp"
	
	If Request.Form("editcat") = "editcategory" Then
		sSQL = "UPDATE Category SET Description = " & Q & FixSpecialChars(request.Form("txtDescription")) & Q & _
				", HeaderText = " & Q & FixSpecialChars(request.Form("txtHeaderText")) & Q & _
				", FooterText = " & Q & FixSpecialChars(request.Form("txtFooterText")) & Q & _
				"WHERE CategoryID = " & CatID
		Set oConn = Server.CreateObject("ADODB.Connection")
		oConn.open gsConnStr
		oConn.execute sSQL
		oConn.close 
		refreshMenuItems
	End If
	
	Response.Write("<P>" & Ref)
	
	'code for deleting categories
	For i = 1 To Request.Form.Count - 1 
		'loop through all form elements and check for items marked for deletion
		Response.Write "<P>" & request.Form.Key(i)
		if LCase(Left(Request.Form.Key(i),12)) = "delcategory_" Then
			'check that there are no products for this category and if there are, do not allow delete
			sSQL = "SELECT COUNT(*) AS CountProducts FROM Products WHERE CategoryID = " & Mid(Request.Form.Key(i), 13, Len(Request.Form.Key(i))-11) & " " & _
					"AND [Description] <> 'Dummy';"
			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.Open sSQL, gsConnStr
			If rs("CountProducts") > 0 Then
				'there are products for this category, so do not allow deletion to continue
				Response.Redirect Ref & "?action=delete_forbidden"
				Response.End
			Else
				Set oConn = Server.CreateObject("ADODB.Connection")
				oConn.Open gsConnStr

				'delete any dummy products
				sSQL = "DELETE * FROM Products WHERE CategoryID = " & Mid(Request.Form.Key(i), 13, Len(Request.Form.Key(i))-11)
				oConn.Execute sSQL
				'there are no products, so it is ok to delete
				sSQL = "DELETE * FROM Category WHERE CategoryID = " & Mid(Request.Form.Key(i), 13, Len(Request.Form.Key(i))-11)
				oConn.Execute sSQL
				oConn.Close 
				refreshMenuItems
				End If
		End If
	Next


'Handle Save
If request.Form("btnOnlineOrders")="Enable Online Orders" Then 
		Set oConn = Server.CreateObject("ADODB.Connection")
		oConn.Open gsConnStr
		sSQL = "Update Admin SET OnlineOrderEnabled = true WHERE ID=1" 
		oConn.Execute sSQL
		oConn.Close 
End IF
If request.Form("btnOnlineOrders")="Disable Online Orders" Then 
		Set oConn = Server.CreateObject("ADODB.Connection")
		oConn.Open gsConnStr
		sSQL = "Update Admin SET OnlineOrderEnabled = false WHERE ID=1" 
		oConn.Execute sSQL
		oConn.Close 
End If

	'code to insert new category
	If request.Form("NewCategory") <> "" Then
		'insert new category row
		sSQL = "INSERT INTO Category(Description) VALUES(" & Q & FixSpecialChars(Request.Form("NewCategory")) & Q & ")"
		Set oConn = Server.CreateObject("ADODB.Connection")		
		oConn.Open gsConnStr
		
		'start transaction as we are entering dummy child entities for the new category
		oConn.BeginTrans
		
		oConn.Execute sSQL
		If Err <> 0 Then
			Response.Write "An Error Occurred: " & err.Description 
			oConn.RollBackTrans
			oConn.Close
			Set oConn = Nothing 
			Err.Clear 
			Response.End 
		End If
		
		'fetch new category id from database
		sSQL = "SELECT CategoryID FROM Category WHERE Description = " & Q & FixSpecialChars(Request.Form("NewCategory")) & Q
		response.Write "<P>" & sSQL
		Set rs = server.CreateObject("ADODB.Recordset")
		rs.Open sSQL, oConn
		If Err <> 0 Then
			Response.Write "An Error Occurred: " & err.Description 
			oConn.RollBackTrans
			oConn.Close
			Set oConn = Nothing 
			Err.Clear 
			Response.End 
		End If
		
		'now create a new dummy product for the new category so that the admin page will display correctly
		If Not rs.BOF And Not rs.EOF Then
			CatID = rs(0)
			rs.Close 
			Set rs = Nothing
			sSQL = "INSERT INTO Products(Title, CategoryID) VALUES('Dummy'," & CatID & ")"
			oConn.Execute sSQL
			If Err <> 0 Then
				response.Write "An Error Occurred: " & err.Description 
				oConn.RollBackTrans
				oConn.Close
				Set oConn = Nothing 
				Err.Clear 
				Response.End 
			End If
			
			'fetch productID for the new dummy product - use this to allocate to dummy item detail
			sSQL = "SELECT ProductID FROM Products WHERE CategoryID = " & CatID
			Set rs = server.CreateObject("ADODB.Recordset")
			rs.Open sSQL, oConn
			If Err <> 0 Then
				Response.Write "An Error Occurred: " & err.Description 
				oConn.RollBackTrans
				oConn.Close
				Set oConn = Nothing 
				Err.Clear 
				response.End 
			End If
			
			'create dummy product detail record for the new dummy product
			If Not rs.BOF And Not rs.EOF Then
				ProdID = rs(0)
				rs.Close
				Set rs = Nothing
				sSQL = "INSERT INTO ProductDetail (ProductID, ProductCode) VALUES(" & ProdID & ", 'Dummy');"
				Response.Write "<P>" & sSQL
				oConn.Execute sSQL
				If Err <> 0 Then
					Response.Write "An Error Occurred: " & err.Description 
					oConn.RollBackTrans
					oConn.Close
					Set oConn = Nothing 					
					Err.Clear 
					Response.End 
				End If
			Else
				oConn.RollBackTrans
				bError = true
			End If
		Else
			oConn.RollBackTrans
			bError = true
		End If
		If Not bError Then 
			oConn.CommitTrans
			refreshMenuItems
		End If
		oConn.Close 
	End If
	
	Set oConn = Nothing
	Response.Redirect Ref

'=====================================================================
Function refreshMenuItems
		'Refreshes cached list of menu items held in application variable
		Dim rs
		Dim sSQL
		Dim varArray()
		Dim i
		Dim Q
		'If IsEmpty(Application("MenuItems"))  Then
			Set rs = Server.CreateObject("ADODB.Recordset")
			sSQL = "SELECT Description, CategoryID FROM Category ORDER BY Description;"
			rs.Open sSQL, gsConnStr
			i = 0
			Q = Chr(34)
			Do While Not rs.EOF
				ReDim Preserve varArray(1, i)
				varArray(0,i) = rs("Description")
				varArray(1,i) = rs("CategoryID")
				i = i + 1
				rs.MoveNext 
			Loop	
			rs.Close
			Set rs = Nothing
			Application("MenuItems") = varArray
End Function
	
%>