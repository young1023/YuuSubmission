
<!--#include file="include/SessionHandler.inc.asp" -->
<%
if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if



Title = "User Level"
%>

<%
' Add Menu
'*********
response.expires = 0

pageno = trim(request("pageno"))

UserLevel = trim(request(replace("UserLevel1","'","''")))

' Which Main Menu to show
'************************
FunctionID = Trim(Request("FunctionID"))

If FunctionID = "" Then

FunctionID = 1

End IF

Message = ""

if trim(request("action_button")) = "add" then


		LevelName = trim(request(replace("LevelName1","'","''")))
		
		LevelNumber = trim(request(replace("Level1","'","''")))

	    strsql="Select * from UserLevel where LevelName='"& LevelName & "' "

        'or LevelNumber = " & LevelNumber
	    
		set acres=conn.execute(strsql)

	    if  acres.eof then
	    
			strsql="insert into UserLevel (LevelName, LevelNumber) values ('"& LevelName & "', "& LevelNumber & ")"
		
			conn.execute strsql
		else
		   
		   Message =  "This Level Name Exits,Please Input Other Name"
		   
		end if
		set acres=nothing
end if

' modify the menu

if trim(request("action_button")) = "Modify" then

	vLevelName = split(trim(request("LevelName2")),",")
	
	vLevel = split(trim(request("Level2")),",")
	
	vid = split(trim(request("mid")),",")
	
	for i=0 to ubound(vid)
	
		strsql="Update UserLevel set LevelName = '"& trim(replace(vLevelName(i),"'","''")) &"' where LevelID="& trim(vid(i))
		
        'response.write strsql
	    conn.execute strsql 
	next
	
	'for i=0 to ubound(vid)
	
		'strsql="Update UserLevel set LevelNumber = "& trim(replace(vLevel(i),"'","''")) &" where LevelID="& trim(vid(i))
		
	    conn.execute strsql 
	'next
	
	
end if


if trim(request("action_button")) = "Delete" then

	vid = split(trim(request("xid")),",")

	for i=0 to ubound(vid)
	
		strsql="Delete * From userlevel where Levelid="& trim(vid(i))
		
	    conn.execute strsql 
	next
	
end if


' sorting the menu

if trim(request("action_button")) = "sort" then

	vOrder = split(trim(request("OrderID")),",")
	
	vid = split(trim(request("mid")),",")
	
	for i=0 to ubound(vid)
	
		strsql="Update UserLevel set LevelNumber = "& trim(replace(vOrder(i),"'","''")) &" where Levelid="& trim(vid(i))
		
	    conn.execute strsql 
	next
	
end if


'************************
Userlevel = trim(request(replace("Userlevel","'","''")))

' sorting the menu

if trim(request("action_button")) = "match" then

	vid = split(trim(request("nid")),",")
	
		sql1 = "Delete From MenuUserLevelTable Where UserLevel = "&UserLevel
		
		Conn.Execute sql1
	
	
	for i=0 to ubound(vid)
	
		sql2 = "Insert into MenuUserLevelTable (MenuID, Userlevel) Values ("&  vid(i)  &","& UserLevel &")"
		
	    conn.execute sql2 
	next
	
end if

%>

<!--#include file="include/SQLconn.inc.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title><% = Title %></title>
<link rel="stylesheet" type="text/css" href="include/uob.css" />
<SCRIPT language=JavaScript>
<!--

function doadd()
{
	if((document.fm1.LevelName1.value)=="")
		alert("Please enter the name of the user level!");
	else
		{
		document.fm1.action_button.value="add";
		document.fm1.submit();
		}
}


function doModify()
{
	
		{
		document.fm1.action_button.value="Modify";
		document.fm1.submit();
		}
}

function doOrder()
{
	var i,strmsg,strval,j;
	strmsg="";
	j=0;
	if(document.fm1.OrderID != null)
	{
		for(i=0;i<document.fm1.OrderID.length;i++)
			{
				strval=document.fm1.OrderID[i].value;
				if(strval=="")
				{
					strmsg="The sorting field cannot be null.";
					j=1;
					break;
				}
			}
		if (j==0 && i==0)
	    {
			if(document.fm1.OrderID.value=="")
				strmsg="The sorting field cannot be null.";
		}
	}
	if (strmsg!="")
		alert(strmsg);
	else
	{
		document.fm1.action_button.value="sort";
		document.fm1.submit();
	}
}


function doSelect(){
k=0;
document.fm1.action="UserLevel.asp?sid=<%=SessionID%>&FunctionID=<%=FunctionID%>&UserLevel=<%=UserLevel%>";
	if (document.fm1.nid!=null)
	{
		for(i=0;i<document.fm1.nid.length;i++)
		{
			if(document.fm1.nid[i].checked)
			  {
			  k=1;
			  i=1;
			  break;
			  }
		}
		if(i==0)
		{
			if (!document.fm1.nid.checked)
               k=0;
			else
               k=1;
		}
	}

if (k==0)
  alert("You must select at least one record!");	
else if (k==1)
 {
  var msg = "Are you sure ?";
  if (confirm(msg)==true)
   {
    document.fm1.action_button.value="match";
    document.fm1.submit();
   }
 }

}

function delcheck(){
k=0;
document.fm1.action="UserLevel.asp?sid=<%=SessionID%>&FunctionID=<%=FunctionID%>";
	if (document.fm1.xid!=null)
	{
		for(i=0;i<document.fm1.xid.length;i++)
		{
			if(document.fm1.xid[i].checked)
			  {
			  k=1;
			  i=1;
			  break;
			  }
		}
		if(i==0)
		{
			if (!document.fm1.xid.checked)
               k=0;
			else
               k=1;
		}
	}

if (k==0)
  alert("You must select at least one record!");	
else if (k==1)
 {
  var msg = "Are you sure ?";
  if (confirm(msg)==true)
   {
    document.fm1.action_button.value="Delete";
    document.fm1.submit();
   }
 }

}

//-->
</SCRIPT>
</head>
<body leftmargin="0" topmargin="0">


<!-- #include file ="include/Master.inc.asp" -->


<div id="Content">
<%
'-----------------------------------------------------------------------------
'
'      Main Content of the page is inserted here
'
'-----------------------------------------------------------------------------

%>
 <table width="80%" border="0" class="normal">

    <tr> 
      <td  class="BlueClr" <% If FunctionID <> 1 Then %>bgcolor="#C0C0C0"<% End If %> width="30%" >
      <a href="UserLevel.asp?sid=<% =SessionID %>&functionid=1">Add User Level
      </a>
      </td> 
      <td  class="BlueClr" width="30%" <% If FunctionID <> 2 Then %>bgcolor="#C0C0C0"<% End If %>>
      <a href="UserLevel.asp?sid=<% =SessionID %>&functionid=2">Edit/Delete User Level
      </a></td> 
      <td  class="BlueClr" <% If FunctionID <> 3 Then %>bgcolor="#C0C0C0"<% End If %>>
      <a href="UserLevel.asp?sid=<% =SessionID %>&functionid=3">Manage User level</a></td>
       <td  class="BlueClr" <% If FunctionID <> 4 Then %>bgcolor="#C0C0C0"<% End If %>>
      <a href="UserLevel.asp?sid=<% =SessionID %>&functionid=4&UserLevel=7">Access Right</a></td>  
    </tr>

   </table>

<br>

<form name="fm1" method="post" action="">
<% If FunctionID <> 1 Then %>		
  		
  <div style="display:none" align=center>
  		
<% End If %>
  

<table width="80%" border="0" class="normal">
  <tr> 
      <td colspan="2"  align="right">
<font color="red">*</font>	 Denotes a mandatory field</td>
    </tr>
  <tr> 
      <td colspan="2"  align="right">
�@</td>
    </tr>
    
  
    
 <tr> 
      <td width="27%">User Level Name</td> 
      <td width="69%">
      	     
				<Input name="LevelName1" type=text value="" size="50"></td>
    </tr>
    
 <tr> 
      <td width="27%">
Level</td> 
      <td width="69%">
      <Input name="Level1" type=text value="" size="2"></td>
    </tr>
    
 <tr> 
      <td width="27%">
�@</td> 
      <td width="69%">
      	<input type="button" value="    Add   " onClick="javascript:doadd();">
        <input type="hidden" name="action_button" value="">     
�@</td>
    </tr>
    
 <tr> 
      <td colspan="2" align =center><font color="red"><% = Message %></font></td> 
      <td >
�@</td>
    </tr>
    
<%
	sql = " SELECT * from UserLevel order by LevelNumber Asc"
				set acres = nothing
				set acres = createobject("adodb.recordset")
				acres.cursortype = 1
				acres.locktype = 1
				acres.open sql,conn
				
				recount = acres.recordcount
			
				if pageno="" then
				   pageno=1
			    end if 
			    
				 if pageno>1 then
				   i = (pageno-1) * 10
				   acres.move i
				 end if
				 
				   
	
%>
<tr> 
      <td width="27%"><% call showpageno(pageno) %></td> 
      <td width="69%">
      �@</td>
    </tr>

     <tr> 
      <td width="27%" bgcolor="#FFFFCC">Level Name</td> 
      <td width="69%" bgcolor="#FFFFCC">Level
      �@</td>
    </tr>
<% 
	if not acres.eof then
	  	do while not acres.eof
		   if j> 9 then exit do
%>

     <tr> 
      <td width="27%"><% = acres("LevelName") %>�@</td> 
      <td width="69%"><% = acres("LevelNumber") %></td>
    </tr>
<%
	acres.movenext  
	j=j+1 
	loop 
%>
<% end if  %>
    <tr bgcolor="#ffffff"> 
      <td colspan="2">
        <div align="right"> 
          <%call showpageno(pageno)%>
          </div>
      </td>
    </tr>
</table>
</div>

<!---- End of Add User Level ------>

<!---- Start of Edit/Detete User Level ---->

  <% If FunctionID <> 2 Then %>		
  		
  <div style="display:none" align=center>
  		
  <% End If %>


<table width="80%" border="0" class="normal">

     <tr> 
      <td width="85%" colspan="2">�@</td> 
      <td width="13%">

      �@</td>
    </tr>
    
    <%
	sql = " SELECT * from UserLevel order by LevelNumber Asc"
				
	Set acres = Conn.Execute(sql)
				 
	
%>
     <tr> 
      
      <td width="43%" bgcolor="#FFFFCC">Level Name</td>
      
      <td width="42%" bgcolor="#FFFFCC">Level</td>
       <td width="13%" bgcolor="#FFFFCC">Select</td>
    </tr>
<% 
	if not acres.eof then
	  	do while not acres.eof
		 
%>

     <tr> 
      <td width="43%">     	     
				<Input name="LevelName2" type=text value="<% = acres("LevelName") %>" size="50"></td>
      <td width="42%">     	     
				<% = acres("LevelNumber") %></td>
				      <td width="13%"><input type="checkbox" name="xid" value="<% = acres("LevelID") %>"></td> 

    </tr>
<%
	acres.movenext  
	
	loop 
%>
<% end if  %>

     <tr> 
      <td width="85%" colspan="2">
      <input type="button" value="Modify" onClick="doModify();">
        �@</td>
              <td width="13%">
        <input type="button" value="Delete" onClick="delcheck();"></td> 

    </tr>
</table>
</div>

<!---- End of Edit/Delete Menu ------>
<!---- Start of Sorting Menu ---->

  <% If FunctionID <> 3 Then %>		
  		
  <div style="display:none" align=center>
  		
  <% End If %>


<table width="80%" border="0" class="normal">



	<tr> 
			<td colspan="3">�@</td> 
			
			 
			�@</td>
	</tr>
	    
    <%
	sql = " SELECT * from UserLevel order by levelNumber Asc"
	
	Set acres = Conn.Execute(sql)

				
%>
     <tr> 
      <td width="50%" bgcolor="#FFFFCC">Level Name
      �@</td>
       <td width="30%" bgcolor="#FFFFCC">Level
      �@</td>
    </tr>
<% 
	if not acres.eof then
	  	do while not acres.eof
		   
%>

     <tr> 
       
      <td>     	     
				<% = acres("LevelName") %></td>
				<td>     	     
				<Input name="OrderID" type=text value="<% = acres("LevelNumber") %>" size="2" maxlength="2">
				<input type="hidden" name="mid" value="<% = acres("LevelID") %>">
				</td>

    </tr>
<%
	acres.movenext  
	
	loop 
%>
<% end if  %>

     <tr> 
      <td>
    
              <td width="13%">
        <input type="button" value="  Sort" onClick="doOrder();"></td> 

    </tr>
    </table>
</div>

<!---- End of Sorting Menu ------>
<!---- Start of Acess Right Menu ---->

  <% If FunctionID <> 4 Then %>		
  		
  <div style="display:none" align=center>
  		
  <% End If %>
  

<table width="80%" border="0" class="normal">
<tr> 

<%
        
		sql2 = " Select * From UserLevel Order By LevelNumber Desc"
		
		Set Rs2 = Conn.Execute(sql2)
		
		If Not Rs2.EoF Then
		
			Do While Not Rs2.EoF
			
			
		
%>
	
			<td <% If Userlevel <>  Trim(Rs2("LevelNumber"))  Then %>bgcolor="#FFFFCC"<% End If %>>
			<a href="MenuSetup.asp?FunctionID=4&Userlevel=<% = Rs2("LevelNumber") %>&sid=<% =SessionID %>"><% = Rs2("LevelName") %></a>
			</td>
<%
		Rs2.MoveNext
			Loop
			
End If

%>

	</tr>
	
</table>
<br>
<table width="80%" border="0" class="normal">

	<tr> 
			<td width="21%" bgcolor="#FFFFCC">User Right</td> 
			<td width="77%" bgcolor="#FFFFCC">
			
			�@</td>
	</tr>
</table> 	


<table width="80%" border="0" class="normal">
 
 
	<tr> 
			<td width="39%"></td> 
			<td width="60%">
			�@</td>
	</tr>
	
<%
		sql = " SELECT * from Menu order by OrderID"
		
		set acres = Conn.Execute(sql)
	
    
	if not acres.eof then
	  	do while not acres.eof
		  
%>
	  <tr> 
      <td width="39%"></td> 
      <td width="60%">
      
<% 
	  Sql2 = "Select * From MenuUserLevelTable where MenuID = " & acres("ID") & " and UserLevel = " & Userlevel 
	  'Sql2 = "Select * From MenuUserLevelTable where  UserLevel = " & UserLevel & " and MenuID = " & acres("ID") 
	  
	  'response.write sql2
          'response.end
	  
	  Set Rs2 = Conn.Execute(Sql2)
	  
	  If Not Rs2.EoF Then
	  	
	  		SelectFlag = 1
	  		
	  End If
	  
%>
      
      <input type="checkbox" name="nid" value="<% = acres("ID") %>" <% If SelectFlag = 1 Then%>Checked<% End If %>>&nbsp;
      <% = acres("MenuName") %>
           �@</td>
    </tr>
	
<%
	acres.movenext 
	
		SelectFlag = 0 
	
	loop 

 
	End If
%>
	
	
      <tr> 
      <td width="39%"></td> 
      <td width="60%">
      <input type="button" value="Submit" onClick="doSelect();"></td>
    </tr>
</table> 
</div>
</div>
	
	
                 
</table>
</form>

<%


'-----------------------------------------------------------------------------
'
'      End of the main Content 
'
'-----------------------------------------------------------------------------
%>
</div>
            
              </body>

              </html>
<% 
function showpageno(pageno)
	if recount>10 then
		lastpage=recount\10
		if recount mod 10 >0 then
			lastpage=lastpage+1
		end if
		if pageno>10 then
		     response.write "<a href='UserLevel.asp?sid="&SessionID&"&functionid="&FunctionID&"&pageno=1'> The First Page</a>&nbsp;&nbsp;"
			response.write "<a href='UserLevel.asp?sid="&SessionID&"&functionid="&FunctionID&"&pageno="&(pageno-9-(pageno  mod 10) )&"'>Previous 10</a>&nbsp;&nbsp;"
		end if
		strtemp=pageno
		if (pageno Mod 10 )=0 then
		   strtemp=strtemp-10
		end if
	 for i=(strtemp-(strtemp mod 10)+1) to (strtemp+10-(strtemp mod 10))
	         if lastpage<i then  exit for			 
            if i- pageno=0 then
				response.write cstr(i)&"&nbsp;&nbsp;"
			else
				response.write "<a href='UserLevel.asp?sid="&SessionID&"&functionid="&FunctionID&"&Pageno="&i&"'>"&cstr(i)&"</a>&nbsp;&nbsp;"
			end if	
		next
		if (pageno\10)<(lastpage\10) then
		        response.write "<a href='UserLevel.asp?sid="&SessionID&"&functionid="&FunctionID&"&Pageno="&(pageno+1-(pageno mod 10)) &"'>Next 10</a>&nbsp;&nbsp;"
			   response.write "<a href='UserLevel.asp?sid="&SessionID&"&functionid="&FunctionID&"&Pageno="& (lastpage) &"'>Last Page</a>&nbsp;&nbsp;"
		end if
		
 end if
end function
%>