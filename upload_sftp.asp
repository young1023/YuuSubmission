<!--#include file="include/SessionHandler.inc.asp" -->


<%

Dim wShell

Set wShell = Server.CreateObject("WScript.Shell") 


on error resume next

wShell.Run "C:\TheClub\Upload_TheClub_Files.bat", 1, false


%>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" type="text/css" href="include/uob.css" />
<TITLE>Upload File</TITLE>


</head>
<body leftmargin="0" topmargin="0">


<!-- #include file ="include/Master.inc.asp" -->


<div id="Content">


      <table width="96%" height="400" border="0" cellspacing="0" cellpadding="2" class="Normal">
        <tr>
          <td align="middle">


<%

    If Err.Number <> 0 Then

             Response.Write "Upload suceesfully!"


    Else


            Response.Write "Upload problem. Please try again."


    End If


%>

          </td>

        </tr>
       
 


<tr>
<td align="middle">



</td>
   </tr>
     </table>

</div>
 </body>
    </html>