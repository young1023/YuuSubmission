<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="include/SessionHandler.inc.asp" -->
<%
if session("shell_power")="" then
  response.redirect "logout.asp?r=1"
end if

MemberID = Session("MemberID")

Title = "Report From Yuu"
   
Const adVarChar = 200
Const adInteger = 3
Const adDate = 7

sPath = Server.MapPath("\report\")

Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
Set oFolder = objFSO.GetFolder(sPath)

Dim objFSO, oFolder, objRS



Sub ListFolderContents(sPath, sTypeToShow)
    Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
    Set oFolder = objFSO.GetFolder(sPath)

    'create a disconnected recordset
    Set objRS = Server.CreateObject("ADODB.Recordset")

    'append proper fields
    objRS.Fields.Append "Name", adVarChar, 255
    'objRS.Fields.Append "Url", adVarChar, 255
    objRS.Fields.Append "Size", adVarChar, 255
    objRS.Fields.Append "DateLastModified", adDate
    objRS.Open

    'extract all files in given path:
    'Call ExtractAllFiles(oFolder)  

    For Each oFile in oFolder.Files
        objRS.AddNew
        objRS.Fields("Name").Value = oFile.Name
        'objRS.Fields("Url").Value = MapURL(oFile.Path)
        objRS.Fields("Size") = oFile.size
        objRS.Fields("DateLastModified").Value = oFile.DateLastModified
       response.write objRS("Name") & "<br/>"
       response.write "<b>" & objRS("DateLastModified") & "</b><br/>"
    Next  

    'sort and apply:

    
    If Not(objRS.EOF) Then 
        objRS.Sort = "DateLastModified DESC"
        objRS.MoveFirst
    End If

    i = 1

    'loop through all the records:

    Do Until objRS.EOF

    %>

       

                <%=i%><a href="/report/<%=objRS("Name")%> download><%=objRS("Name")%></a>


                <%=objRS("Size")%>
                
           

                <%=objRS("DateLastModified")%><br/>
                    
           

<%

        objRS.MoveNext

        i = i + 1

    Loop

    

    
End Sub
   



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
   <th align="left" width="30%">size</th>
   
   <th align="left">Date</th>
   
</tr>

<%

 Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
    Set oFolder = objFSO.GetFolder(sPath)

    'create a disconnected recordset
    Set objRS = Server.CreateObject("ADODB.Recordset")

    'append proper fields
    objRS.Fields.Append "Name", adVarChar, 255
    'objRS.Fields.Append "Url", adVarChar, 255
    objRS.Fields.Append "Size", adVarChar, 255
    objRS.Fields.Append "DateLastModified", adDate
    objRS.Open

    'extract all files in given path:
    'Call ExtractAllFiles(oFolder)  

    For Each oFile in oFolder.Files
        objRS.AddNew
        objRS.Fields("Name").Value = oFile.Name
        'objRS.Fields("Url").Value = MapURL(oFile.Path)
        objRS.Fields("Size") = oFile.size
        objRS.Fields("DateLastModified").Value = oFile.DateLastModified
       'response.write objRS("Name") & "<br/>"
       'response.write "<b>" & objRS("DateLastModified") & "</b><br/>"
    Next  

    'sort and apply:

    
    If Not(objRS.EOF) Then 
        objRS.Sort = "DateLastModified DESC"
        objRS.MoveFirst
    End If

    i = 1

    'loop through all the records:

    Do Until objRS.EOF

    %>

    <tr>

            <td>
       
                <%=i%>.&nbsp;<a href="/report/<%=objRS("Name")%>" download><%=objRS("Name")%></a>

            </td>

            <td>

                <%=objRS("Size")%>&nbsp;KB
                
            </td>

            <td>

                <%=objRS("DateLastModified")%>

            </td>

    </tr>        
                    
           

<%
        'response.write i

        objRS.MoveNext

        i = i + 1

    Loop


    'clean up resources
    Set oFolder = Nothing
    Set objFSO = Nothing
    objRS.Close
    Set objRS = Nothing


%>

</table>
    </div>
            
              </body>

              </html>