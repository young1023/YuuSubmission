

<!--#include file="include/SessionHandler.inc.asp" -->
<%
if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if


whatdo = trim(request.form("action_button"))

pageid=trim(request("pageid"))

If request.form("action_button") = "deleteFile" then
 

     FileName = Request.form("deletefile")
 
  ' create the file system objects

   strPhysicalPath = Server.MapPath("\batch\Source")

   Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

   Set objFolder = objFSO.GetFolder(strPhysicalPath)
  
  
    'Delete the file'
    objFSO.DeleteFile(objFolder&"\"&FileName)


    'Open log file and write log'
    Set objFile = objFSO.OpenTextFile(Server.MapPath("\batch\log\log.txt"),8)
    objFile.Write FileName & " was deleted on " & FormatDateTime(now()) & "." & Chr(13) & Chr(10)
    objFile.Close


End IF

%>

<html>

<head>

<title>Delete File</title>

<meta http-equiv="Content-Type" content="text/html; charset=utf-8">

<meta http-equiv="Refresh" content="2; url=FileOnServer.asp">

<link rel="stylesheet" type="text/css" href="include/uob.css" />

</head>

<body leftmargin="0" topmargin="0">


<!-- #include file ="include/Master.inc.asp" -->


<table border=0 cellpadding=3 cellspacing=2 class="normal" width="90%" height="100%">

  <tr> 
    <td align="center">

    File <%=FileName%> was deleted.

    <p>

      <p>

        <p>

  <a href="FileOnServer.asp">Return</a>

    </td>
  
  </tr>
  
</table>

            
</body>

</html>