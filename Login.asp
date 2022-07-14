


<%
' Get Login and Password
num = replace(trim(request.form("num")),"'","''")
inkey = replace(trim(request.form("inkey")),"'","''")
ipaddress  = Request.ServerVariables("REMOTE_ADDR")


' Reset System Message
Session("SystemMessage") = ""
%>
<!--#include file="include/SQLConn.inc.asp" -->

<%
	Session.Timeout=600
	response.expires=0

	' Check if the Login Name Exist
	' --------------------------------------------------------------------------------------------------------------------	
	
	'set Rs1 = server.createobject("adodb.recordset")

    SQL1 = "Select * from Member where LoginName = '" & num & "' and pwd = '" & inkey & "'"

     Set Rs1 = Conn.Execute(SQL1)
       
	If  Rs1.EoF Then

			
					
					'Unsucess login attempt
					'**********************
					
					
					'Session.Abandon
					session("shell_power")= 0
					Session("SystemMessage") = "Login failed."
	
          			Response.Redirect "default.asp"	
	
			Else
			
					'Sucessful Login
					'***********************

                    Getdate = now()

                    SQL2 = "Insert into AuditLog (Description, DoneBy, CreateDate)"

                    SQL2 = SQL2 & " values ('log in successfully', '"& rs1("MemberID") &"' , '" & getdate &"')"
  
                    Conn.execute(SQL2)
					
					
			        Session("id") = num
			        Session("MemberID") = rs1("MemberID")
			        Session("shell_power") = Rs1("UserLevel")
   					Response.Redirect "FileOnServer.asp?sid=" & Session("SessionID")
   				
	end if







%>
