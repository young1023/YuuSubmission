<!--#include file="Common.js" -->
<!--#include file="SQLConn.inc.asp" -->
<table width="100%" border="0" cellspacing="0" cellpadding="4" class="Common">


			<tr valign="top" align="center">
			<td>
			
		<% if session("shell_power") > 0 then %>
			
			<table align=center border=0 cellpadding=5 cellspacing=0 width=160>
			
			<tr> 
				<td height=28   class="NaviBar">User Menu</td>
			</tr>
			
			
<%
						
						
		' Generate User menu
	
        ' set Rs = server.createobject("adodb.recordset")

	SQL1 = "Select * From MenuUserLevelTable t Left Join Menu m on m.ID = t.MenuID "

    SQL1 = SQL1 & "where t.UserLevel = " & Session("shell_power") & " Order by m.OrderID"

    

    Set Rs = Conn.Execute(SQL1)
					
						  
						If Not Rs.EoF Then
	  						Do While Not Rs.EoF

						%>  
									
									
									<tr> 
									<td height=8   class="NaviBar">
									<% If Rs("PageLink") <> "" Then %>
									<a href="<% = Rs("PageLink") %>?sid=<%=SessionID%>" style="TEXT-DECORATION: none"><% = Rs("MenuName") %></a>
									<% Else %>
									<% = Rs("MenuName") %>
									<% End If %>
									</td>
									</tr>
									
									
<%
	Rs.Movenext  
	
	Loop 
 
	End If


%>
	
	<% if session("shell_power") > 7 and session("shell_power") < 10 then %>
								
	<tr> 
		<td height=8   class="NaviBar"><a href="MenuSetup.asp?sid=<%=SessionID%>" style="TEXT-DECORATION: none">Menu Setup</a></td>
					</tr>
	
		<% End If %>
			<tr>
			<td height=28   class="NaviBar"><a href="logout.asp?r=0&sid=<%=SessionID%>" style="TEXT-DECORATION: none">Logout</a> </td>
			</tr>
        		

	</table>

	
	<% End If %>
	
			</td>
			</tr>
		
</table>
