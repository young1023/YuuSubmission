
<%
'Set user level to 1'
session("shell_power") = 1

    FileName = "log.txt"
 
  ' create the file system objects

   strPhysicalPath = Server.MapPath("\batch\Log") 

   Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

   Set objFolder = objFSO.GetFolder(strPhysicalPath)

   Set objFile = objFSO.OpenTextFile(objFolder&"\"&FileName, 1)
   

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>Log Information</title>
<link rel="stylesheet" type="text/css" href="include/uob.css" />
</head>
<body leftmargin="0" topmargin="0">

<!-- #include file ="include/Master.inc.asp" -->

<div id="Content">
	
               <table width="98%" border="0" cellpadding="2" cellspacing="2" class="normal">
         
                    <tr>
                        <td>
                          
 <% 

   Dim arrFileLines()
   i = 0

   'Initiate line number
    lineNo = 1

    Do Until objFile.AtEndOfStream

    Redim Preserve arrFileLines(i)
     arrFileLines(i) = objFile.ReadLine
     i = i + 1

    
    Loop

    objFile.Close

    For l = Ubound(arrFileLines) to LBound(arrFileLines) Step -1

      Response.write lineNo & ". " &  arrFileLines(l) & "<Br/>"

      lineNo = lineNo + 1

    Next

%>         

                        </td>
                      </tr>

                    </table>
            
</div>            
  
            
  </body>

</html>
