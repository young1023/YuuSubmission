
<!--#include file="include/SessionHandler.inc.asp" -->
<%
if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if
%>


<% 
Title = "Member Management" 

'Check User Right
if session("shell_power")="" then
  response.redirect "default.asp"
end if

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title><% = Title %></title>
<link rel="stylesheet" type="text/css" href="include/uob.css" />
<SCRIPT language=JavaScript>
<!--

<% 'START password generator  
%>

function getRandomNum(lbound, ubound) {
return (Math.floor(Math.random() * (ubound - lbound)) + lbound);
}
function getRandomChar(number, lower, upper, other, extra) {
var numberChars = "0123456789";
var lowerChars = "abcdefghijklmnopqrstuvwxyz";
var upperChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
var otherChars = "`~!@#$%^&*()-_=+[{]}\\|;:'\",<.>/? ";
var charSet = extra;
if (number == true)
charSet += numberChars;
if (lower == true)
charSet += lowerChars;
if (upper == true)
charSet += upperChars;
if (other == true)
charSet += otherChars;
return charSet.charAt(getRandomNum(0, charSet.length));
}
function getPassword(length, extraChars, firstNumber, firstLower, firstUpper, firstOther,
latterNumber, latterLower, latterUpper, latterOther) {
var rc = "";
if (length > 0)
rc = rc + getRandomChar(firstNumber, firstLower, firstUpper, firstOther, extraChars);
for (var idx = 1; idx < length; ++idx) {
rc = rc + getRandomChar(latterNumber, latterLower, latterUpper, latterOther, extraChars);
}
return rc;
}
<% 'END password generator 
%>
function dosubmit(){
 document.addm.action="execute.asp?sid=<%=SessionID%>";
 if (document.addm.LoginName.value == "")
  {
   alert("Please enter the login name.");
   document.addm.LoginName.focus();
   return false;
  }
   if (document.addm.UserName.value == "")
  {
   alert("Please enter the user name.");
   document.addm.UserName.focus();
   return false;
  }

 if (document.addm.Password.value == "")
  {
   alert("Please enter the password.");
   document.addm.Password.focus();
   return false;
  }
  



document.addm.submit();
}

function NewMember(){
 document.ok.action="NewMember.asp?sid=<%=SessionID%>";
 document.ok.submit();
}

function doConvert(){
window.open("ConvertMember.asp?sid=<%=sessionid%>"); 

}


//-->
</SCRIPT>

<script language="JavaScript">
function disableCtrlKeyCombination(e)
{
        //list all CTRL + key combinations you want to disable
        var forbiddenKeya = 'a';
        var forbiddenKeyc = 'c';
        var forbiddenKeyx = 'x';


        var key;
        var isCtrl;

        if(window.event)
        {
                key = window.event.keyCode;     //IE
                if(window.event.ctrlKey)
                        isCtrl = true;
                else
                        isCtrl = false;
        }
        else
        {
                key = e.which;     //firefox
                if(e.ctrlKey)
                        isCtrl = true;
                else
                        isCtrl = false;
        }

        //if ctrl is pressed check if other key is in forbidenKeys array
        if(isCtrl)
        {
            
                {
                        //case-insensitive comparation
                        if(forbiddenKeya.toLowerCase() == String.fromCharCode(key).toLowerCase())
                        {
                                return false;
                        }
                        if(forbiddenKeyc.toLowerCase() == String.fromCharCode(key).toLowerCase())
                        {
                                return false;
                        }

						if(forbiddenKeyx.toLowerCase() == String.fromCharCode(key).toLowerCase())
                        {
                                return false;
                        }

                }
        }
        return true;
}
</script>

<script language="JavaScript">
<!--
// disable right click
var message="Sorry, The right click function is disable."; // Message for the alert box

function click(e) {
if (document.all) {
if (event.button == 2) {
alert(message);
return false;
}
}
if (document.layers) {
if (e.which == 3) {
alert(message);
return false;
}
}
}
if (document.layers) {
document.captureEvents(Event.MOUSEDOWN);
}
document.onmousedown=click;
// --> 
</script>


</head>

<body leftmargin="0" topmargin="0" onkeypress="return disableCtrlKeyCombination(event);" onkeydown="return disableCtrlKeyCombination(event);" >

<!-- #include file ="include/Master.inc.asp" -->


<div id="Content">
<%
'-----------------------------------------------------------------------------
'
'      Main Content of the page is inserted here
'
'-----------------------------------------------------------------------------
%>

<DIV align=center>

  <TABLE border=0 cellPadding=0 cellSpacing=0 height=100% width=99%>
    <TBODY> 
    <TR>
      <TD vAlign=top>
        <table width="100%" border="0" cellpadding="1" cellspacing="1" height="100%">
          <tr> 
            <td bgcolor="#000000">
              <table width="100%" border="0" cellpadding=0 cellspacing="0" bgcolor="#FFFFFF" height="100%">
                <tr>
                  <td valign="top" align="center" bgcolor="#E6EBEF">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#E6EBEF" class="normal">
                        <tr> 
                          <td valign="top" align="center">
                            <form name=ok method=post>
                            <table width="98%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" class="normal">
                              <tr> 
                                <td height="28"> 
                                  <%

   
        ExpiredAge = 365 'Rs1("SettingValue")

		pageid=trim(request.form("pageid"))
		if pageid="" then
		  pageid=1
		end if
        findnum=replace(trim(request.form("findnum")),"%","�H")
        
        findnum=replace(findnum,"'","''")
        
        SortingID = request("SortingID")
        FilterID = request("FilterID")
        
         If SortingID = "" Then
         
         SortingID = "MemberID"
         
         Elseif SortingID = "MemberName" Then
         
         SortingID = "m.Name"
         
         Elseif SortingID = "Branch" Then
         
         SortingID = "u.Name"
         
         End If
         
         
        fsql = "Select * From Member m Left Join UserLevel l on m.UserLevel = l.LevelNumber "
	    
	    'fsql  =  fsql  &  " Left Join UserGroup u on m.GroupID = u.GroupID and u.sharing=0" 
	    
        If findnum <> "" Then
        
        fsql = fsql & "Where (m.memberName like '%"&findnum&"%') "
        
        End If
        
        fsql = fsql & " Order by "&SortingID 
        
        If FilterID = 1 Then
        
		fsql  =  fsql  &  " desc" 
        
        End If
	

		
		'response.write fsql
        set frs=createobject("adodb.recordset")
		frs.cursortype=3
		frs.locktype=1
        frs.open fsql,conn

       if frs.RecordCount=0 then
%>
      <font color=red>No Record <a href="sa_member.asp?sid=<%=SessionID%>">[ Return ]</a></font>

<%
       else
          findrecord=frs.recordcount
          response.write "Total <font color=red>"&findrecord&"</font> Records ; Total <font color=blue>"

         frs.PageSize = 1000
         call countpage(frs.PageCount,pageid)
     %>
	     <input type="text" name="findnum" size="20" value="<% =findnum %>" class="common">
		 <input type="button" value=" Search by Name " onClick="findenum();" class="common">
	<%   end if
	  %>
                                </td>
                              </tr>
                              <tr> 
                                <td valign="top" height="28"> 
                                  <div align="center">
                                  <table width="98%" border="0" cellspacing="1" class="normal" bgcolor="#C0C0C0" cellpadding="6">
                                    <tr bgcolor = "#FFFFFF"> 
									<td><font color="#FFFFFF"><a href="sa_member.asp?SortingID=MemberID&sid=<%=SessionID%>" style="text-decoration: none">Member ID</a></font></td>
                                    <td height="23" width="15%">
                                    &nbsp;User Name<a href="sa_member.asp?SortingID=MemberName&sid=<%=SessionID%>"><br/></a><a href="sa_member.asp?SortingID=MemberName&FilterID=1&sid=<%=SessionID%>"></a>
                                    </td>
									<td>
									Login Name<br/>												
									<a href="sa_member.asp?SortingID=LoginName&sid=<%=SessionID%>">
									</a><a href="sa_member.asp?SortingID=LoginName&FilterID=1&sid=<%=SessionID%>"></a>
									</td>
																		<td>
										  Password<a href="sa_member.asp?SortingID=Dept&sid=<%=SessionID%>"><br/></a><a href="sa_member.asp?SortingID=Dept&FilterID=1&sid=<%=SessionID%>"></a></td>
																		<td>
										  Email<br/><a href="sa_member.asp?SortingID=Email&sid=<%=SessionID%>"></a><a href="sa_member.asp?SortingID=Email&FilterID=1&sid=<%=SessionID%>"></a></td>
																		<td width="20%">
										  User Right<br/><a href="sa_member.asp?SortingID=levelName&sid=<%=SessionID%>"></a><a href="sa_member.asp?SortingID=levelName&FilterID=1&sid=<%=SessionID%>"></a></td>

 																		<td width="10%">
																		Lock<br/><a href="sa_member.asp?SortingID=Lock&sid=<%=SessionID%>"></a><a href="sa_member.asp?SortingID=Lock&FilterID=1&sid=<%=SessionID%>"></a></td>                                      
                                    <td width="7%">Delete</td>
                               </tr>
                               
                                    <%
                                    
                                
 i=0
 if frs.recordcount>0 then
  frs.AbsolutePage = pageid
  do while (frs.PageSize-i)
   if frs.eof then exit do
   i=i+1
%>
   <tr bgcolor = "#FFFFFF">
   <%
    response.write "<td>"&frs("memberid")&"</td>"
   response.write "<td>"
   %>
   

    <a href="MemberDetail.asp?sid=<%=SessionID%>&MemberID=<%=frs("MemberID")%>"><% =trim(frs("MemberName"))%></a>

 
   <%
   	response.write "</td>"
    response.write "<td>"&frs("Loginname")&"</td>"
    response.write "<td>"&frs("Pwd")&"</td>"
    response.write "<td>"&frs("email")&"</td>"
    response.write "<td>"&frs("LevelName")&"</td>"
   response.write "<td align=center>"
   If frs("Lock") > 3 Then
   response.write "Locked"
   End If
   response.write "</td>"
   response.write "<td align=center>"
   response.write "<input type=checkbox name=mid value="&frs("MemberID")&">"
   response.write "</td>"
   response.write "</tr>"
   flage=not flage
   frs.movenext
  loop
 end if
  %>
                                  </table>
                                	</div>
                                </td>
                              </tr>
                              <tr> 
                                <td align="right" height="28"> 
                                  <script language=JavaScript>
<!--
function delcheck(){
k=0;
document.ok.action="delid.asp?sid=<%=SessionID%>";
	if (document.ok.mid!=null)
	{
		for(i=0;i<document.ok.mid.length;i++)
		{
			if(document.ok.mid[i].checked)
			  {
			  k=1;
			  i=1;
			  break;
			  }
		}
		if(i==0)
		{
			if (!document.ok.mid.checked)
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
    document.ok.whatdo.value="delmember";
    document.ok.submit();
   }
 }

}

function gtpage(what)
{
document.ok.pageid.value=what;
document.ok.action="sa_member.asp?sid=<%=SessionID%>";
document.ok.submit();
}

function findenum()
{
document.ok.pageid.value=1;
document.ok.action="sa_member.asp?sid=<%=SessionID%>";
document.ok.submit();
}
//-->
                                  </script>
                                  <%
	 if frs.recordcount>0 then
             call countpage(frs.PageCount,pageid)
			 response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			 if Clng(pageid)<>1 then
                 response.write " <a href=javascript:gtpage('1') style='cursor:hand' >First</a> "
                 response.write " <a href=javascript:gtpage('"&(pageid-1)&"') style='cursor:hand' >Previous</a> "
			 else
                 response.write " First "
                 response.write " Previous "
			 end if
	         if Clng(pageid)<>Clng(frs.PageCount) then
                 response.write " <a href=javascript:gtpage('"&(pageid+1)&"') style='cursor:hand' >Next</a> "
                 response.write " <a href=javascript:gtpage('"&frs.PageCount&"') style='cursor:hand' >Last</a> "
             else
                 response.write " Next "
                 response.write " Last "
			 end if
	         response.write "&nbsp;&nbsp;"
	 end if
%>
                                </td>
                              </tr>
                              <tr> 
                                <td height="28" align="center"> 
                                  <%
if frs.recordcount>0 then
  response.write "<input type='button' value='   New Member   ' onClick='NewMember();' class='common'>&nbsp;"
  response.write "<input type='button' value='   Delete   ' onClick='delcheck();' class='common'>&nbsp;"
  'response.write "<input type='button' value='   Excel   ' onClick='doConvert();' class='common'>&nbsp;"
  response.write "<input type=hidden value='' name=whatdo>"
  response.write "<input type=hidden value="&pageid&" name=pageid>"
end if
			  frs.close
			  set frs=nothing
			  'conn.close
			  'set conn=nothing
%>
                                </td>
                              </tr>
                              <tr> 
                                <td valign="top"></td>
                              </tr>
                            </table>
                          </form>


                          </td>
                        </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <TR>
      <TD colSpan=5 height=11 align=center>
        <script language=JavaScript src="img/copyright.js"></script>
      </TD>
    </TR>
</TABLE>
</DIV>
<%

  ' function
  Sub countpage(PageCount,pageid)
  response.write pagecount&"</font> Pages "
	   if PageCount>=1 and PageCount<=10 then
		 for i=1 to PageCount
		   if (pageid-i =0) then
              response.write "<font color=green> "&i&"</font> "
		   else
             response.write " <a href=javascript:gtpage('"&i&"') style='cursor:hand' >"&i&"</a>"
		   end if
		 next
	   elseif PageCount>11 then
	      if pageid<=5 then
		     for i=1 to 10
		       if (pageid-i =0) then
                 response.write "<font color=green> "&i&"</font> "
		       else
                 response.write " <a href=javascript:gtpage('"&i&"') style='cursor:hand' >"&i&"</a>"
		       end if
		     next
		  else
		    for i=(pageid-4) to (pageid+5)
		       if (pageid-i =0) then
                 response.write "<font color=green> "&i&"</font> "
		       elseif i=<pagecount then
                 response.write " <a href=javascript:gtpage('"&i&"') style='cursor:hand' >"&i&"</a>"
		       end if
			next
		  end if
	   end if
  end sub
%>



























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