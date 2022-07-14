<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="include/SessionHandler.inc.asp" -->
<%
if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if

MemberID = Session("MemberID")

Title = "Report From Yuu"
   
   'On Error Resume Next

   ' declare variables
   Dim objFSO, objFolder
   Dim objCollection, objItem

   Dim strPhysicalPath, strTitle, strServerName
   Dim strPath, strTemp
   Dim strName, strFile, strExt, strAttr
   Dim intSizeB, intSizeK, intAttr, dtmDate


   ' create the file system objects
   strPhysicalPath = "C:\inetpub\wwwroot\sftp\report"  
   Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
   Set objFolder = objFSO.GetFolder(strPhysicalPath)


   Set outputLines = CreateObject("System.Collections.ArrayList")
      for each f in objFSO.GetFolder(strPhysicalPath).files
         outputLines.Add f.Name
	 
      next

   outputLines.Reverse()
   



%>

<!--#include file="include/SQLconn.inc.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><% = Title %></title>
<link rel="stylesheet" type="text/css" href="include/uob.css" />
<SCRIPT language=JavaScript>
<!--

function doDelete(){
document.fm1.action="GetAMReport.asp";
document.fm1.action_button.value="deleteFile";
document.fm1.submit();
}

function doGet(){
document.fm1.action="GetAMReport.asp";
document.fm1.action_button.value="GetFile";
document.fm1.submit();
}

//-->
</SCRIPT>
</head>
<body leftmargin="0" topmargin="0">


<!-- #include file ="include/Master.inc.asp" -->

<div id="Content">

 
<table width="90%" border="0" cellspacing="0" cellpadding="4"  class="normal" >




<tr>
   <th align="left" width="40%">Name</th>
   <th align="left" width="30%">B</th>
   
   <th align="left">Date</th>
   
</tr>



<%
   ''''''''''''''''''''''''''''''''''''''''
   ' output the file list
   ''''''''''''''''''''''''''''''''''''''''

   Set objCollection = objFolder.Files

    For Each outputLine in outputLines

  set file = objFSO.GetFolder(strPhysicalPath).files.item (outputLine&"")

  strFile = file.name 

  intSizeB = File.Size

 

%>


<tr>

   

   <td align="left">
 
   <a href="/report/<%=strFile%>" download>
   <%=strFile%></a>


</td>

   <td align="left"><%=intSizeB%></td>
  
   <td align="left"><%= File.DateLastModified %></td>

</tr>

 

<% Next %>

</table>

<%
   Set objFSO = Nothing
   Set objFolder = Nothing
%>


</div>
            
              </body>

              </html>