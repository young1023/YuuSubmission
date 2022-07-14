<!--#include file="include/SessionHandler.inc.asp" -->
<!-- #include file="ShadowUpload.asp" -->
<%
      sFolder = "batch\Source"

dim fs, fo, ts, f

set fs=Server.CreateObject("Scripting.FileSystemObject")

%>

<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" type="text/css" href="include/uob.css" />
<TITLE>Upload File</TITLE>

</head>

<body leftmargin="0" topmargin="0">


<!-- #include file ="include/Master.inc.asp" -->


<div id="Content">


      <table width="98%" height="400" border="1" cellspacing="0" cellpadding="2" class="Normal">
        <tr>
          <td align="middle" height="50">

Upload file to Server

          </td>

        </tr>
        <tr>
          <td align="middle">

<%

Dim objUpload


    Set objUpload = New ShadowUpload

    ' Check if there is error'
    If objUpload.GetError<>"" Then

        Response.Write("sorry, could not upload: "&objUpload.GetError)

    Elseif objUpload.File(x).Size > 15000000 Then

        Response.Write("sorry, the file exceeded 15M.")

    Else


        For x=0 To objUpload.FileCount-1

            Response.Write("file name: "&objUpload.File(x).FileName&"<br /><br />")

            Response.Write("file size: "&objUpload.File(x).Size&"<br /><br />")

         ' Check if distinction file exists, delete the file if exists

         If fs.FileExists(Server.MapPath(sFolder) & objUpload.File(x).FileName)  Then

              fs.DeleteFile(Server.MapPath(sFolder) & objUpload.File(x).FileName)
 
         end if
  
    'Start uploading'
    Call objUpload.File(x).SaveToDisk(Server.MapPath(sFolder) , "")


        
    'Open log file and write log'
    Set objFile = fs.OpenTextFile(Server.MapPath("\batch\log\log.txt"),8)
    objFile.Write objUpload.File(x).FileName & " was uploaded on " & FormatDateTime(now()) & "." & Chr(13) & Chr(10)
    objFile.Close


    ' redirect to main page after uploading'
    response.redirect "FileOnServer.asp?sid="&sessionid
      

        Next

    End If

    'reset the variable'
    Set objUpload = Nothing


      

%>


</td>
  </tr>
 

     </table>

</div>

 </body>
    </html>