<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="include/SessionHandler.inc.asp" -->
<%
if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if

MemberID = Session("MemberID")
   
On Error Resume Next

   ' declare variables
   Dim objFSO, objFolder, outputLines
   Dim strPhysicalPath, objCollection
   Dim strName, strFile, strExt, strAttr
   Dim intSizeB, intSizeK, intAttr, dtmDate


   ' create the file system objects
   strPhysicalPath = Server.MapPath("\batch\Source")
   Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
   Set objFolder = objFSO.GetFolder(strPhysicalPath)


   Set outputLines = CreateObject("System.Collections.ArrayList")
      for each f in objFSO.GetFolder(strPhysicalPath).files
         outputLines.Add f.Name
      next

   'outputLines.Reverse()
   
   
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>Files on Server</title>
<link rel="stylesheet" type="text/css" href="include/uob.css" />

<SCRIPT language=JavaScript>
<!--

function doDelete(){
document.fm1.action="deletefile.asp";

if((document.fm1.deletefile.value)=="")
  {
    alert("Please select file to delete!");
    return false;

  }
    
  else
  
document.fm1.action_button.value="deleteFile";
document.fm1.submit();
  
}

//-->
</SCRIPT>
</head>

<body leftmargin="0" topmargin="0">

<!-- #include file ="include/Master.inc.asp" -->

<div id="Content">

  <form name=fm1 method=post>

 <table width="90%" border="0" cellspacing="0" cellpadding="4"  class="normal" >

<tr>
   <th align="left" width="60%">Name </th>

   <th align="left" width="10%">KB</th>
   
   <th align="left">Date</th>

   <th align="left" width="10%">Delete</th>
   
</tr>



<%
   ''''''''''''''''''''''''''''''''''''''''
   ' output the file list
   ''''''''''''''''''''''''''''''''''''''''

   Set objCollection = objFolder.Files

   ' Set file number'
   i = 1

    For Each outputLine in outputLines

  set file = objFSO.GetFolder(strPhysicalPath).files.item (outputLine&"")

  strFile = file.name 

  intSizeB = File.Size

  intSizeK = Int((intSizeB/1024) + .5)

%>


<tr>
   <td align="left">
 
   <% =i & " .   " %>  <%=strFile%>

</td>

   <td align="left"><%=intSizeK%>K</td>
  
   <td align="left"><%= File.DateLastModified %></td>

   <td align="left"><input type="radio" name="deletefile" value="<%=strFile%>" <% If i = 1 Then%>checked<%End If%>></td>

</tr>

<% i = i + 1 %>

<% Next %>


<tr>

<td colspan="4" align="center">

        
               <input type="Button" value=" Delete " onClick="doDelete();" class="Normal">

                  <input type="hidden" name="action_button" value="">  

</td>

  </tr>

</table>

</form>

<%
   Set objFSO = Nothing
   Set objFolder = Nothing
%>


</div>
            
              </body>

              </html>