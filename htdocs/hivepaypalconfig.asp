<%
'------------------------------------------------------------------------
'config.asp
'
'ASP Toolkit for PayPal v0.50
'http://www.paypal.com/pdn
'
'Copyright (c) 2004 PayPal Inc
'
'Released under Common Public License 1.0
'http://opensource.org/licenses/cpl.php
' Live config settings
'------------------------------------------------------------------------ 
%>



<%

'Configuration Settings 
' paypal_business="123@shit.com"
paypal_business="websales@thehivehoneyshop.co.uk"
 paypal_site_url="http://thehivehoneyshop.co.uk/"
 paypal_image_url=""
' paypal_success_url="asp_toolkit/success.asp"
 paypal_success_url="paypaldemo/success.asp"
 paypal_cancel_url="paypaldemo/error.asp"
 paypal_notify_url="paypaldemo/ipn/ipn.asp"
 paypal_return_method="2" ' 1=GET 2=POST  
 paypal_currency_code="GBP" ' [USD,GBP,JPY,CAD,EUR 
 paypal_lc="GB"
'Production  URL (Live)  
 paypal_url="https://www.paypal.com/cgi-bin/webscr" 
'Sandbox URL (Test)  
'paypal_url="https://www.sandbox.paypal.com/cgi-bin/webscr" 
 paypal_bn="toolkit-asp"
 paypal_cmd="_cart"
' paypal_cmd="_xclick"



'Payment Page Settings  
 paypal_display_comment="0" ' 0=yes 1=no  
 paypal_comment_header="Delivery Instructions"
 paypal_continue_button_text="Continue >>"
 paypal_background_color="" ' ""=white 1=black  
 paypal_display_shipping_address="1" ' ""=yes 1=no  
 paypal_display_comment="1" ' ""=yes 1=no  



'Product Settings  
 paypal_item_name=Request.Form("item_name")
 paypal_item_number=Request.Form("item_number")
 paypal_amount=Request.Form("amount") 
 paypal_on0=Request.Form("on0")
 paypal_os0=Request.Form("os0")
 paypal_on1=Request.Form("on1")
 paypal_os1=Request.Form("os1")
 paypal_quantity=Request.Form("quantity")
 paypal_edit_quantity="" ' 1=yes ""=no  
 paypal_invoice=""
 paypal_tax=""




'Shipping and Taxes  
 paypal_shipping_amount=Request.Form("shipping_amount")
 paypal_shipping_amount_per_item=""
 paypal_handling_amount=""
 paypal_custom_field=""



'Customer Settings  
 paypal_firstname=Request.Form("firstname")
 paypal_lastname=Request.Form("lastname")
 paypal_address1=Request.Form("address1")
 paypal_address2=Request.Form("address2")
 paypal_city=Request.Form("city")
 paypal_state=Request.Form("state")
 paypal_zip=Request.Form("zip")
 paypal_email=Request.Form("email")
 paypal_phone_1=Request.Form("phone1")
 paypal_phone_2=Request.Form("phone2")
 paypal_phone_3=Request.Form("phone3")

%>
