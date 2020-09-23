VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "ASP Gen"
   ClientHeight    =   1275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   ScaleHeight     =   1275
   ScaleWidth      =   5475
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtProjectPath 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Text            =   "C:\ASPGen"
      Top             =   360
      Width           =   5415
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select Database"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Project Folder Path:"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrProperties(14) As String
Dim TableBGColor As String
Dim FirstHalfSQL As String
Dim SecondHalfSQL As String
Dim EndSQL As String

Public Function CreateProject(idbPath, iUserRootPath, iProjectName) As Boolean

    Dim objDAO As New DAO.DBEngine
    Dim objDatabase As DAO.Database
    Dim objTable As DAO.TableDef
    Dim objField As DAO.Field
    Dim objProperty As DAO.Property
    Dim intProjectID As Integer
    Dim intTableID As Integer
    Dim intColumnID As Integer
    Dim UserName As String
    Dim iProjectRootPath As String
    Dim iProjectTablePath As String
    Dim fsoObject As New FileSystemObject
    Dim bDoesFolderExist As Boolean
    Dim xCount As Integer
    Dim iProjectID As String
    Dim CurrentPrimaryKey As String
    Dim objIndex As DAO.Index
    Dim objIndexField As DAO.Field
    Dim tfolders As Variant
    Dim driveRoot As String
    Dim zCount As Integer
    Dim CurrentPath As String
    
    If CheckGoodDatabase(idbPath) = False Then
        CreateProject = False
        Exit Function
    End If
    
    tfolders = Split(iUserRootPath, "\", , vbTextCompare)
    driveRoot = tfolders(0)
    For xCount = 1 To UBound(tfolders)
        CurrentPath = driveRoot
        For zCount = 1 To xCount
            CurrentPath = CurrentPath & "\" & tfolders(zCount)
        Next zCount
        If Not fsoObject.FolderExists(CurrentPath) Then
            Call fsoObject.CreateFolder(CurrentPath)
        End If
    Next
    
    xCount = 1
    iProjectRootPath = iUserRootPath & "\Project" & xCount
    
    If fsoObject.FolderExists(iProjectRootPath) Then
        bDoesFolderExist = True
        While bDoesFolderExist = True
            xCount = xCount + 1
            iProjectRootPath = iUserRootPath & "\Project" & xCount
            bDoesFolderExist = fsoObject.FolderExists(iProjectRootPath)
        Wend
        iProjectRootPath = iUserRootPath & "\Project" & xCount
    End If
    iProjectID = xCount
    
    Call AddProjectFolder(idbPath, iUserRootPath, iProjectID, iProjectName, iProjectRootPath)
    
    Set objDatabase = objDAO.OpenDatabase(idbPath)
    
    For Each objTable In objDatabase.TableDefs
    
        If Mid(objTable.Name, 1, 4) <> "MSys" Then
            iProjectTablePath = iProjectRootPath & "\" & objTable.Name
            Call AddTableFolder(iUserRootPath, iProjectRootPath, iProjectTablePath, iProjectID, objTable.Name)
            
            For Each objIndex In objTable.Indexes
                For Each objIndexField In objIndex.Fields
                    If objIndex.Primary = True Then
                        'primary key located for table
                        CurrentPrimaryKey = objIndexField.Name
                    End If
                Next
            Next
            
            For Each objField In objTable.Fields
                
                Erase arrProperties
                For Each objProperty In objField.Properties
                    
                    Select Case objProperty.Name
                        Case "Attributes"
                            arrProperties(1) = objProperty.Value
                        Case "Type"
                            arrProperties(2) = objProperty.Value
                            If objProperty.Value = 11 Then
                                objField.Properties("Required").Value = False
                                arrProperties(8) = False
                            End If
                        Case "OrdinalPosition"
                            arrProperties(3) = objProperty.Value
                        Case "Size"
                            arrProperties(4) = objProperty.Value
                        Case "DefaultValue"
                            arrProperties(5) = objProperty.Value
                        Case "ValidationRule"
                            arrProperties(6) = objProperty.Value
                        Case "ValidationText"
                            arrProperties(7) = objProperty.Value
                        Case "Required"
                            arrProperties(8) = objProperty.Value
                        Case "AllowZeroLength"
                            arrProperties(9) = objProperty.Value
                        Case "DecimalPlaces"
                            arrProperties(10) = objProperty.Value
                        Case "Format"
                            arrProperties(11) = objProperty.Value
                        Case "Name"
                            arrProperties(12) = objProperty.Value
                            If CurrentPrimaryKey = objProperty.Value Then
                                arrProperties(13) = True
                            Else
                                arrProperties(13) = False
                            End If
                        Case Else
                            
                    End Select
                Next
                
                Call CreateAddTablePostHTML(iProjectRootPath, iProjectID, objTable.Name)
                Call CreateAddTableASP(iProjectRootPath, iProjectID, objTable.Name)
                Call CreateViewTable(iProjectRootPath, iProjectID, objTable.Name)
                Call CreateGetTableData(iProjectRootPath, iProjectID, objTable.Name)
                Call CreateDeleteTable(iProjectRootPath, iProjectID, objTable.Name)
                Call CreateUpdateTablePostHTML(iProjectRootPath, iProjectID, objTable.Name)
                Call CreateUpdateTableASP(iProjectRootPath, iProjectID, objTable.Name)
                
            Next
        End If
    Next
     
    CreateProject = True
    
    Set objDAO = Nothing
    Set objDatabase = Nothing
    Set objTable = Nothing
    Set objField = Nothing
    Set objProperty = Nothing
    
End Function

Private Function AddProjectFolder(ByRef idbPath, iUserRootPath, iProjectID, iProjectName, iProjectRootPath) As Boolean

    Dim fsoObject As New FileSystemObject
    Dim fsoTextStream As TextStream
    Dim fsoTextStreamTemp As TextStream
    Dim CurrentLine As String
    Dim fsoTextStreamTable As TextStream
    Dim BGColor As String
    
    Call fsoObject.CreateFolder(iProjectRootPath)
    Call fsoObject.CopyFile(idbPath, iProjectRootPath & "\database.mdb")
    
    idbPath = iProjectRootPath & "\database.mdb"

    Set fsoTextStreamTable = fsoObject.CreateTextFile(iProjectRootPath & "\default.asp")
    
    Call fsoTextStreamTable.WriteLine("<%@ Language=VBScript %>")
    Call fsoTextStreamTable.WriteLine("<HTML>")
    Call fsoTextStreamTable.WriteLine("<HEAD>")
    Call fsoTextStreamTable.WriteLine("<TITLE>" & iProjectName & " table directory</TITLE>")
    Call fsoTextStreamTable.WriteLine("</HEAD>")
    Call fsoTextStreamTable.WriteLine("<BODY>")
    Call fsoTextStreamTable.WriteLine("<table border=""0"" width=""100%"" cellpadding=""2"" bgcolor=""#000080"">")
    Call fsoTextStreamTable.WriteLine("<tr>")
    Call fsoTextStreamTable.WriteLine("<td><font color=""#FFFFFF"">Table Directory</font></td>")
    Call fsoTextStreamTable.WriteLine("</tr>")
    Call fsoTextStreamTable.WriteLine("</table>")
    Call fsoTextStreamTable.WriteLine("<br>")
    Call fsoTextStreamTable.WriteLine("<table border=""0"" width=""450"">")
    Call fsoTextStreamTable.WriteLine("<!-- LinkStarter. Do not modyify -->")
    Call fsoTextStreamTable.WriteLine("</table>")
    Call fsoTextStreamTable.WriteLine("<p><a href=""../"">Return to project list</a></p>")
    Call fsoTextStreamTable.WriteLine("</BODY>")
    Call fsoTextStreamTable.WriteLine("</HTML>")
    Call fsoTextStreamTable.Close
    
    Set fsoTextStreamTable = Nothing
    
End Function

Private Function AddTableFolder(iUserRootPath, iProjectRootPath, iProjectTablePath, iProjectID, iTableName) As Boolean
        
    Dim fsoObject As New FileSystemObject
    Dim fsoTextStream As TextStream
    Dim fsoTextStreamTemp As TextStream
    Dim fsoTextStreamAddPage As TextStream
    Dim CurrentLine As String
    
    Call fsoObject.CreateFolder(iProjectRootPath & "\" & iTableName)
    
    Set fsoTextStream = fsoObject.OpenTextFile(iProjectRootPath & "\default.asp", ForReading)
    Set fsoTextStreamTemp = fsoObject.CreateTextFile(iProjectRootPath & "\defaultTemp.asp", True)
    
    While Not fsoTextStream.AtEndOfStream
        
        CurrentLine = fsoTextStream.ReadLine
        fsoTextStreamTemp.WriteLine CurrentLine
        
        If CurrentLine = "<!-- LinkStarter. Do not modyify -->" Then
        
            If TableBGColor = "#C0C0C0" Then
                TableBGColor = "#F3F3DC"
            Else
                TableBGColor = "#C0C0C0"
            End If
        
            Call fsoTextStreamTemp.WriteLine("<tr>")
            Call fsoTextStreamTemp.WriteLine("<td width=""200"" bgcolor=""" & TableBGColor & """>" & iTableName & "</td>")
            Call fsoTextStreamTemp.WriteLine("<td width=""125"" bgcolor=""" & TableBGColor & """><center><a href=""" & iTableName & "/Add.asp"">Add New Record</a></center></td>")
            Call fsoTextStreamTemp.WriteLine("<td width=""125"" bgcolor=""" & TableBGColor & """><center><a href=""" & iTableName & "/View.asp"">View Data</a></center></td>")
            Call fsoTextStreamTemp.WriteLine("</tr>")
        End If
    Wend
    
    Call fsoTextStream.Close
    Call fsoTextStreamTemp.Close
    
    Call fsoObject.DeleteFile(iProjectRootPath & "\default.asp", True)
    Call fsoObject.MoveFile(iProjectRootPath & "\defaultTemp.asp", iProjectRootPath & "\default.asp")

    Set fsoTextStreamAddPage = fsoObject.CreateTextFile(iProjectTablePath & "\add.asp")
    
    Call fsoTextStreamAddPage.WriteLine("<%@ Language=VBScript %>")
    Call fsoTextStreamAddPage.WriteLine("<%on error resume next%>")
    Call fsoTextStreamAddPage.WriteLine("<HTML>")
    Call fsoTextStreamAddPage.WriteLine("<HEAD>")
    Call fsoTextStreamAddPage.WriteLine("<TITLE>" & iTableName & " Insert Data Page</TITLE>")
    Call fsoTextStreamAddPage.WriteLine("</HEAD>")
    Call fsoTextStreamAddPage.WriteLine("<BODY>")
    Call fsoTextStreamAddPage.WriteLine("<table border=""0"" width=""100%"" cellpadding=""2"" bgcolor=""#000080"">")
    Call fsoTextStreamAddPage.WriteLine("<tr>")
    Call fsoTextStreamAddPage.WriteLine("<td><font color=""#FFFFFF"">Add Information To Table " & iTableName & "</font></td>")
    Call fsoTextStreamAddPage.WriteLine("</tr>")
    Call fsoTextStreamAddPage.WriteLine("</table>")
    Call fsoTextStreamAddPage.WriteLine("<br>")
    Call fsoTextStreamAddPage.WriteLine("<% if request.querystring(""fieldempty"") <> """" then")
    Call fsoTextStreamAddPage.WriteLine("response.write(""Please fill all required fields. Field "" & request.querystring(""fieldempty"") & "" was left empty."")")
    Call fsoTextStreamAddPage.WriteLine("end if ")
    Call fsoTextStreamAddPage.WriteLine("if request.querystring(""duplicatedata"") = ""true"" then")
    Call fsoTextStreamAddPage.WriteLine("response.write(""Duplicate data entered in primary key field."")")
    Call fsoTextStreamAddPage.WriteLine("end if ")
    Call fsoTextStreamAddPage.WriteLine("if request.querystring(""invaliddata"") <> """" then")
    Call fsoTextStreamAddPage.WriteLine("response.write(""Invalid data entered in field "" & request.querystring(""invaliddata"") & ""."")")
    Call fsoTextStreamAddPage.WriteLine("end if ")
    Call fsoTextStreamAddPage.WriteLine("if request.querystring(""successful"") = ""true"" then")
    Call fsoTextStreamAddPage.WriteLine("response.write(""Record added successfully."")")
    Call fsoTextStreamAddPage.WriteLine("end if")
    Call fsoTextStreamAddPage.WriteLine("if request.querystring(""nodata"") = ""true"" then")
    Call fsoTextStreamAddPage.WriteLine("response.write(""Unable to add data. No data submitted."")")
    Call fsoTextStreamAddPage.WriteLine("end if %>")
    Call fsoTextStreamAddPage.WriteLine("<p><font color=""#008000"">*</font>Primary Key / ")
    Call fsoTextStreamAddPage.WriteLine("<font color=""#FF0000"">*</font>Required Field</p>")
    Call fsoTextStreamAddPage.WriteLine("<form method=""POST"" action=""ASPAdd.asp"">")
    Call fsoTextStreamAddPage.WriteLine("<!-- LinkStarter. Do not modyify -->")
    Call fsoTextStreamAddPage.WriteLine("<p><input type=""submit"" value=""Submit"" name=""B1"">   <input type=""reset"" value=""Reset"" name=""B2""></p>")
    Call fsoTextStreamAddPage.WriteLine("</form>")
    Call fsoTextStreamAddPage.WriteLine("<p><a href=""../"">Return to table list</a></p>")
    Call fsoTextStreamAddPage.WriteLine("</BODY>")
    Call fsoTextStreamAddPage.WriteLine("</HTML>")
    Call fsoTextStreamAddPage.Close
    
    Set fsoTextStreamAddPage = fsoObject.CreateTextFile(iProjectTablePath & "\ASPAdd.asp")
    
    Call fsoTextStreamAddPage.WriteLine("<%@ Language=VBScript %>")
    Call fsoTextStreamAddPage.WriteLine("<%on error resume next%>")
    Call fsoTextStreamAddPage.WriteLine("<HTML>")
    Call fsoTextStreamAddPage.WriteLine("<HEAD>")
    Call fsoTextStreamAddPage.WriteLine("<TITLE>" & iTableName & " Insert Data ASP Page</TITLE>")
    Call fsoTextStreamAddPage.WriteLine("</HEAD>")
    Call fsoTextStreamAddPage.WriteLine("<BODY>")
    Call fsoTextStreamAddPage.WriteLine("<% Set adoConnection = server.CreateObject(""ADODB.Connection"")")
    Call fsoTextStreamAddPage.WriteLine("Set adoRecordset = server.CreateObject(""ADODB.Recordset"")")
    Call fsoTextStreamAddPage.WriteLine("adoConnection.Provider = ""Microsoft.Jet.OLEDB.4.0""")
    Call fsoTextStreamAddPage.WriteLine("Dim strLocation, iLength")
    Call fsoTextStreamAddPage.WriteLine("strLocation = Request.ServerVariables(""PATH_TRANSLATED"")")
    Call fsoTextStreamAddPage.WriteLine("iLength = Len(strLocation)")
    Call fsoTextStreamAddPage.WriteLine("iLength = iLength - 10")
    Call fsoTextStreamAddPage.WriteLine("strLocation = Left(strLocation, iLength)")
    Call fsoTextStreamAddPage.WriteLine("strLocation = strLocation & ""../Database.mdb""")
    Call fsoTextStreamAddPage.WriteLine("adoConnection.Open (""Data Source="" & strLocation)")
    Call fsoTextStreamAddPage.WriteLine("iFieldCount = 0")
    Call fsoTextStreamAddPage.WriteLine("FirstHalfSQL = ""insert into [" & iTableName & "] (""")
    Call fsoTextStreamAddPage.WriteLine("SecondHalfSQL = "") Values (""")
    Call fsoTextStreamAddPage.WriteLine("EndSQL = "")""%>")
    Call fsoTextStreamAddPage.WriteLine("<!-- LinkStarter. Do not modyify -->")
    Call fsoTextStreamAddPage.WriteLine("<% adoRecordset.ActiveConnection = AdoConnection")
    Call fsoTextStreamAddPage.WriteLine("SQLInsert = FirstHalfSQL & SecondHalfSQL & EndSQL")
    Call fsoTextStreamAddPage.WriteLine("if SQLInsert <> ""insert into [" & iTableName & "] () Values ()"" then")
    Call fsoTextStreamAddPage.WriteLine("on error resume next")
    Call fsoTextStreamAddPage.WriteLine("call adoRecordset.Open(SQLInsert)")
    Call fsoTextStreamAddPage.WriteLine("if err.number = -2147467259 then")
    Call fsoTextStreamAddPage.WriteLine("response.redirect(""add.asp?duplicatedata=true"")")
    Call fsoTextStreamAddPage.WriteLine("end if")
    Call fsoTextStreamAddPage.WriteLine("on error goto 0")
    Call fsoTextStreamAddPage.WriteLine("else")
    Call fsoTextStreamAddPage.WriteLine("response.redirect(""add.asp?nodata=true"")")
    Call fsoTextStreamAddPage.WriteLine("end if")
    Call fsoTextStreamAddPage.WriteLine("response.redirect(""add.asp?successful=true"") %>")
    Call fsoTextStreamAddPage.WriteLine("</BODY>")
    Call fsoTextStreamAddPage.WriteLine("</HTML>")
    Call fsoTextStreamAddPage.Close
    
    Set fsoTextStreamAddPage = fsoObject.CreateTextFile(iProjectTablePath & "\view.asp")

    Call fsoTextStreamAddPage.WriteLine("<%@ Language=VBScript %>")
    Call fsoTextStreamAddPage.WriteLine("<%on error resume next%>")
    Call fsoTextStreamAddPage.WriteLine("<HTML>")
    Call fsoTextStreamAddPage.WriteLine("<HEAD>")
    Call fsoTextStreamAddPage.WriteLine("<title>Select Data</title>")
    Call fsoTextStreamAddPage.WriteLine("</head>")
    Call fsoTextStreamAddPage.WriteLine("<body>")
    Call fsoTextStreamAddPage.WriteLine("<table border=""0"" width=""100%"" cellpadding=""2"" bgcolor=""#000080"">")
    Call fsoTextStreamAddPage.WriteLine("<tr>")
    Call fsoTextStreamAddPage.WriteLine("<td><font color=""#FFFFFF"">Select fields to retrieve from " & iTableName & ":</font></td>")
    Call fsoTextStreamAddPage.WriteLine("</tr>")
    Call fsoTextStreamAddPage.WriteLine("</table>")
    Call fsoTextStreamAddPage.WriteLine("<br>")
    Call fsoTextStreamAddPage.WriteLine("<form method=""GET"" action=""getdata.asp"">")
    Call fsoTextStreamAddPage.WriteLine("<!-- LinkStarter. Do not modyify -->")
    Call fsoTextStreamAddPage.WriteLine("<p><input type=""submit"" value=""Submit"" name=""B1"">&nbsp;&nbsp;&nbsp;<input type=""reset"" value=""Reset"" name=""B2""></p>")
    Call fsoTextStreamAddPage.WriteLine("</form>")
    Call fsoTextStreamAddPage.WriteLine("<p><a href=""../"">Return to table list</a></p>")
    Call fsoTextStreamAddPage.WriteLine("</body>")
    Call fsoTextStreamAddPage.WriteLine("</html>")
    Call fsoTextStreamAddPage.Close
    
    Set fsoTextStreamAddPage = fsoObject.CreateTextFile(iProjectTablePath & "\getdata.asp")

    Call fsoTextStreamAddPage.WriteLine("<%@ Language=VBScript %>")
    Call fsoTextStreamAddPage.WriteLine("<%on error resume next%>")
    Call fsoTextStreamAddPage.WriteLine("<%response.buffer = false%>")
    Call fsoTextStreamAddPage.WriteLine("<HTML>")
    Call fsoTextStreamAddPage.WriteLine("<HEAD>")
    Call fsoTextStreamAddPage.WriteLine("<TITLE>" & iTableName & " Data</TITLE>")
    Call fsoTextStreamAddPage.WriteLine("</HEAD>")
    Call fsoTextStreamAddPage.WriteLine("<BODY>")
    Call fsoTextStreamAddPage.WriteLine("<%Dim arrField()")
    Call fsoTextStreamAddPage.WriteLine("Dim totFieldCount")
    Call fsoTextStreamAddPage.WriteLine("If Request.QueryString(""NAV"") = """" Then")
    Call fsoTextStreamAddPage.WriteLine("intPage = 1")
    Call fsoTextStreamAddPage.WriteLine("Else")
    Call fsoTextStreamAddPage.WriteLine("intPage = Request.QueryString(""NAV"")")
    Call fsoTextStreamAddPage.WriteLine("End If")
    Call fsoTextStreamAddPage.WriteLine("totFieldCount = 0%>")
    Call fsoTextStreamAddPage.WriteLine("<!-- LinkStarter. Do not modyify -->")
    Call fsoTextStreamAddPage.WriteLine("<%if totfieldcount <> 0 then")
    Call fsoTextStreamAddPage.WriteLine("Set adoConnection = server.CreateObject(""ADODB.Connection"")")
    Call fsoTextStreamAddPage.WriteLine("Set adoRecordset = server.CreateObject(""ADODB.Recordset"")")
    Call fsoTextStreamAddPage.WriteLine("adoConnection.Provider = ""Microsoft.Jet.OLEDB.4.0""")
    Call fsoTextStreamAddPage.WriteLine("Dim strLocation, iLength")
    Call fsoTextStreamAddPage.WriteLine("strLocation = Request.ServerVariables(""PATH_TRANSLATED"")")
    Call fsoTextStreamAddPage.WriteLine("iLength = Len(strLocation)")
    Call fsoTextStreamAddPage.WriteLine("iLength = iLength - 11")
    Call fsoTextStreamAddPage.WriteLine("strLocation = Left(strLocation, iLength)")
    Call fsoTextStreamAddPage.WriteLine("strLocation = strLocation & ""../Database.mdb""")
    Call fsoTextStreamAddPage.WriteLine("adoConnection.Open (""Data Source="" & strLocation)")
    Call fsoTextStreamAddPage.WriteLine("adoRecordset.ActiveConnection = AdoConnection")
    Call fsoTextStreamAddPage.WriteLine("SqlSelect = ""select * from [" & iTableName & "]""")
    Call fsoTextStreamAddPage.WriteLine("adoRecordset.CursorLocation = 3")
    Call fsoTextStreamAddPage.WriteLine("adoRecordset.CursorType = 3")
    Call fsoTextStreamAddPage.WriteLine("call adoRecordset.Open(SQLSelect)")
    Call fsoTextStreamAddPage.WriteLine("adoRecordset.PageSize = 10")
    Call fsoTextStreamAddPage.WriteLine("adoRecordset.CacheSize = adoRecordset.PageSize")
    Call fsoTextStreamAddPage.WriteLine("intPageCount = adoRecordset.PageCount")
    Call fsoTextStreamAddPage.WriteLine("intRecordCount = adoRecordset.RecordCount")
    Call fsoTextStreamAddPage.WriteLine("If CInt(intPage) > CInt(intPageCount) Then intPage = intPageCount")
    Call fsoTextStreamAddPage.WriteLine("If CInt(intPage) <= 0 Then intPage = 1")
    Call fsoTextStreamAddPage.WriteLine("If intRecordCount > 0 Then")
    Call fsoTextStreamAddPage.WriteLine("adoRecordset.AbsolutePage = intPage")
    Call fsoTextStreamAddPage.WriteLine("intStart = adoRecordset.AbsolutePosition")
    Call fsoTextStreamAddPage.WriteLine("If CInt(intPage) = CInt(intPageCount) Then")
    Call fsoTextStreamAddPage.WriteLine("intFinish = intRecordCount")
    Call fsoTextStreamAddPage.WriteLine("Else")
    Call fsoTextStreamAddPage.WriteLine("intFinish = intStart + (adoRecordset.PageSize - 1)")
    Call fsoTextStreamAddPage.WriteLine("End If%>")
    Call fsoTextStreamAddPage.WriteLine("<h4>Records")
    Call fsoTextStreamAddPage.WriteLine("<%=intStart%> through <%=intFinish%> out of <%=intRecordCount%>.</h4>")
    Call fsoTextStreamAddPage.WriteLine("<table border=""0"">")
    Call fsoTextStreamAddPage.WriteLine("<tr>")
    Call fsoTextStreamAddPage.WriteLine("<%fieldcount = 0")
    Call fsoTextStreamAddPage.WriteLine("for each tempField in arrField %>")
    Call fsoTextStreamAddPage.WriteLine("<td bgcolor=""#000080""><font color=""#FFFFFF""><%=tempField%></font>&nbsp;&nbsp;</td>")
    Call fsoTextStreamAddPage.WriteLine("<%fieldcount = fieldcount + 1")
    Call fsoTextStreamAddPage.WriteLine("next%>")
    Call fsoTextStreamAddPage.WriteLine("</tr>")
    Call fsoTextStreamAddPage.WriteLine("<tr>")
    Call fsoTextStreamAddPage.WriteLine("<%for xcount = 1 to fieldcount%>")
    Call fsoTextStreamAddPage.WriteLine("<td>&nbsp;</td>")
    Call fsoTextStreamAddPage.WriteLine("<%next%>")
    Call fsoTextStreamAddPage.WriteLine("</tr>")
    Call fsoTextStreamAddPage.WriteLine("<%bcolor = ""#COCOCO""%>")
    Call fsoTextStreamAddPage.WriteLine("<%For intRecord = 1 To adoRecordset.PageSize%>")
    Call fsoTextStreamAddPage.WriteLine("<tr>")
    Call fsoTextStreamAddPage.WriteLine("<% qString = """" %>")
    Call fsoTextStreamAddPage.WriteLine("<% for each temparrfield in arrfield %>")
    Call fsoTextStreamAddPage.WriteLine("<td bgcolor=""<%=bcolor%>"">&nbsp;<%=adorecordset(temparrfield)%></td>")
    Call fsoTextStreamAddPage.WriteLine("<%next%>")
    Call fsoTextStreamAddPage.WriteLine("<% for each tempField in adoRecordset.fields %>")
    Call fsoTextStreamAddPage.WriteLine("<% if not isnull(adorecordset(tempfield.name)) then")
    Call fsoTextStreamAddPage.WriteLine(" encodeField = server.urlencode(adorecordset(tempfield.name))")
    Call fsoTextStreamAddPage.WriteLine(" else")
    Call fsoTextStreamAddPage.WriteLine(" encodeField = """"")
    Call fsoTextStreamAddPage.WriteLine(" end if")
    Call fsoTextStreamAddPage.WriteLine(" tempFieldName = Replace(tempfield.name, "" "", """")")
    Call fsoTextStreamAddPage.WriteLine(" qString = qString & ""&"" & ""dat"" & tempFieldName & ""="" & encodeField%>")
    Call fsoTextStreamAddPage.WriteLine("<%next%>")
    Call fsoTextStreamAddPage.WriteLine("<td bgcolor=""<%=bcolor%>"">&nbsp;<a href=""delete.asp?<%=request.querystring & qString%>"">delete</a></td>")
    Call fsoTextStreamAddPage.WriteLine("<td bgcolor=""<%=bcolor%>"">&nbsp;<a href=""update.asp?<%=request.querystring & qString%>"">update</a></td>")
    Call fsoTextStreamAddPage.WriteLine("<%adorecordset.MoveNext")
    Call fsoTextStreamAddPage.WriteLine("If bcolor = ""#COCOCO"" Then")
    Call fsoTextStreamAddPage.WriteLine("bcolor = ""#F3F3DC""")
    Call fsoTextStreamAddPage.WriteLine("Else")
    Call fsoTextStreamAddPage.WriteLine("bcolor = ""#COCOCO""")
    Call fsoTextStreamAddPage.WriteLine("End If")
    Call fsoTextStreamAddPage.WriteLine("If adorecordset.EOF Then Exit For")
    Call fsoTextStreamAddPage.WriteLine("Next%>")
    Call fsoTextStreamAddPage.WriteLine("</tr>")
    Call fsoTextStreamAddPage.WriteLine("<%else")
    Call fsoTextStreamAddPage.WriteLine("response.write(""<i>No Data Available</i>"")")
    Call fsoTextStreamAddPage.WriteLine("end if")
    Call fsoTextStreamAddPage.WriteLine("else")
    Call fsoTextStreamAddPage.WriteLine("response.redirect(""view.asp?nodata=true"")")
    Call fsoTextStreamAddPage.WriteLine("end if%>")
    Call fsoTextStreamAddPage.WriteLine("</table>")
    Call fsoTextStreamAddPage.WriteLine("<%if intRecordCount > 0 then%>")
    Call fsoTextStreamAddPage.WriteLine("<%tempstring = request.querystring")
    Call fsoTextStreamAddPage.WriteLine("foundvalue = InStrRev(tempstring, ""&NAV="", Len(tempstring), vbTextCompare)")
    Call fsoTextStreamAddPage.WriteLine("if foundvalue <> 0 then")
    Call fsoTextStreamAddPage.WriteLine("tempstring = Mid(tempstring, 1, foundvalue - 1)")
    Call fsoTextStreamAddPage.WriteLine("end if%>")
    Call fsoTextStreamAddPage.WriteLine("<br>")
    Call fsoTextStreamAddPage.WriteLine("<%If CInt(intPage) > 1 Then%>")
    Call fsoTextStreamAddPage.WriteLine("<a href=""getdata.asp?<%=tempstring%>&NAV=<%=intPage - 1%>""><< Prev</a>")
    Call fsoTextStreamAddPage.WriteLine("<%else%>")
    Call fsoTextStreamAddPage.WriteLine("<< Prev")
    Call fsoTextStreamAddPage.WriteLine("<%End IF")
    Call fsoTextStreamAddPage.WriteLine("If CInt(intPage) < CInt(intPageCount) Then%>")
    Call fsoTextStreamAddPage.WriteLine("<a href=""getdata.asp?<%=tempstring%>&NAV=<%=intPage + 1%>"">Next >></a>")
    Call fsoTextStreamAddPage.WriteLine("<%else%>")
    Call fsoTextStreamAddPage.WriteLine("Next >>")
    Call fsoTextStreamAddPage.WriteLine("<%End If%>")
    Call fsoTextStreamAddPage.WriteLine("<%End If%>")
    Call fsoTextStreamAddPage.WriteLine("<p><a href=""view.asp"">Return to selection page</a></p>")
    Call fsoTextStreamAddPage.WriteLine("</BODY>")
    Call fsoTextStreamAddPage.WriteLine("</HTML>")
    Call fsoTextStreamAddPage.Close
    
    Set fsoTextStreamAddPage = fsoObject.CreateTextFile(iProjectTablePath & "\delete.asp")

    Call fsoTextStreamAddPage.WriteLine("<%@ Language=VBScript %>")
    Call fsoTextStreamAddPage.WriteLine("<%on error resume next%>")
    Call fsoTextStreamAddPage.WriteLine("<HTML>")
    Call fsoTextStreamAddPage.WriteLine("<HEAD>")
    Call fsoTextStreamAddPage.WriteLine("<title>Delete Data</title>")
    Call fsoTextStreamAddPage.WriteLine("</head>")
    Call fsoTextStreamAddPage.WriteLine("<body>")
    Call fsoTextStreamAddPage.WriteLine("<%totfieldcount = 0")
    Call fsoTextStreamAddPage.WriteLine("qString = """"%>")
    Call fsoTextStreamAddPage.WriteLine("<!-- LinkStarter. Do not modyify -->")
    Call fsoTextStreamAddPage.WriteLine("<%if totfieldcount <> 0 then")
    Call fsoTextStreamAddPage.WriteLine("Set adoConnection = server.CreateObject(""ADODB.Connection"")")
    Call fsoTextStreamAddPage.WriteLine("Set adoRecordset = server.CreateObject(""ADODB.Recordset"")")
    Call fsoTextStreamAddPage.WriteLine("adoConnection.Provider = ""Microsoft.Jet.OLEDB.4.0""")
    Call fsoTextStreamAddPage.WriteLine("Dim strLocation, iLength")
    Call fsoTextStreamAddPage.WriteLine("strLocation = Request.ServerVariables(""PATH_TRANSLATED"")")
    Call fsoTextStreamAddPage.WriteLine("iLength = Len(strLocation)")
    Call fsoTextStreamAddPage.WriteLine("iLength = iLength - 10")
    Call fsoTextStreamAddPage.WriteLine("strLocation = Left(strLocation, iLength)")
    Call fsoTextStreamAddPage.WriteLine("strLocation = strLocation & ""../Database.mdb""")
    Call fsoTextStreamAddPage.WriteLine("adoConnection.Open (""Data Source="" & strLocation)")
    Call fsoTextStreamAddPage.WriteLine("adoRecordset.ActiveConnection = AdoConnection")
    Call fsoTextStreamAddPage.WriteLine("if qString <> """" then")
    Call fsoTextStreamAddPage.WriteLine("response.write(qstring)")
    Call fsoTextStreamAddPage.WriteLine("adoRecordset.open (""delete from [" & iTableName & "] where "" & qString)")
    Call fsoTextStreamAddPage.WriteLine("response.redirect(""getdata.asp?"" & mid(QueryCheckString,1,len(QueryCheckString)-1)) & ""&NAV="" & request.querystring(""NAV"")")
    Call fsoTextStreamAddPage.WriteLine("else")
    Call fsoTextStreamAddPage.WriteLine("response.write(""no data specified"")")
    Call fsoTextStreamAddPage.WriteLine("end if")
    Call fsoTextStreamAddPage.WriteLine("end if%>")
    Call fsoTextStreamAddPage.WriteLine("</body>")
    Call fsoTextStreamAddPage.WriteLine("</html>")
    Call fsoTextStreamAddPage.Close

    Set fsoTextStreamAddPage = fsoObject.CreateTextFile(iProjectTablePath & "\update.asp")
    
    Call fsoTextStreamAddPage.WriteLine("<%@ Language=VBScript %>")
    Call fsoTextStreamAddPage.WriteLine("<%on error resume next%>")
    Call fsoTextStreamAddPage.WriteLine("<HTML>")
    Call fsoTextStreamAddPage.WriteLine("<HEAD>")
    Call fsoTextStreamAddPage.WriteLine("<TITLE>" & iTableName & " Update Data Page</TITLE>")
    Call fsoTextStreamAddPage.WriteLine("</HEAD>")
    Call fsoTextStreamAddPage.WriteLine("<BODY>")
    Call fsoTextStreamAddPage.WriteLine("<table border=""0"" width=""100%"" cellpadding=""2"" bgcolor=""#000080"">")
    Call fsoTextStreamAddPage.WriteLine("<tr>")
    Call fsoTextStreamAddPage.WriteLine("<td><font color=""#FFFFFF"">Update Information In " & iTableName & "</font></td>")
    Call fsoTextStreamAddPage.WriteLine("</tr>")
    Call fsoTextStreamAddPage.WriteLine("</table>")
    Call fsoTextStreamAddPage.WriteLine("<br>")
    Call fsoTextStreamAddPage.WriteLine("<% if request.querystring(""fieldempty"") <> """" then")
    Call fsoTextStreamAddPage.WriteLine("response.write(""Please fill all required fields. Field "" & request.querystring(""fieldempty"") & "" was left empty."")")
    Call fsoTextStreamAddPage.WriteLine("request.querystring(""datHyperlink"")")
    Call fsoTextStreamAddPage.WriteLine("end if ")
    Call fsoTextStreamAddPage.WriteLine("if request.querystring(""duplicatedata"") = ""true"" then")
    Call fsoTextStreamAddPage.WriteLine("response.write(""Duplicate data entered in primary key field."")")
    Call fsoTextStreamAddPage.WriteLine("end if ")
    Call fsoTextStreamAddPage.WriteLine("if request.querystring(""invaliddata"") <> """" then")
    Call fsoTextStreamAddPage.WriteLine("response.write(""Invalid data entered in field "" & request.querystring(""invaliddata"") & ""."")")
    Call fsoTextStreamAddPage.WriteLine("end if ")
    Call fsoTextStreamAddPage.WriteLine("if request.querystring(""successful"") = ""true"" then")
    Call fsoTextStreamAddPage.WriteLine("response.write(""Record added successfully."")")
    Call fsoTextStreamAddPage.WriteLine("end if")
    Call fsoTextStreamAddPage.WriteLine("if request.querystring(""nodata"") = ""true"" then")
    Call fsoTextStreamAddPage.WriteLine("response.write(""Unable to add data. No data submitted."")")
    Call fsoTextStreamAddPage.WriteLine("end if %>")
    Call fsoTextStreamAddPage.WriteLine("<p><font color=""#008000"">*</font>Primary Key / ")
    Call fsoTextStreamAddPage.WriteLine("<font color=""#FF0000"">*</font>Required Field</p>")
    Call fsoTextStreamAddPage.WriteLine("<form method=""POST"" action=""ASPupdate.asp?<%=request.querystring%>"">")
    Call fsoTextStreamAddPage.WriteLine("<!-- LinkStarter. Do not modyify -->")
    Call fsoTextStreamAddPage.WriteLine("<p><input type=""submit"" value=""Update"" name=""B1"">   <input type=""reset"" value=""Reset"" name=""B2""></p>")
    Call fsoTextStreamAddPage.WriteLine("<input type=""hidden"" value=""<%=request.querystring%>"" name=""UpdateQueryString"">")
    Call fsoTextStreamAddPage.WriteLine("<input type=""hidden"" value=""<%=qstring%>"" name=""qString"">")
    Call fsoTextStreamAddPage.WriteLine("<input type=""hidden"" value=""<%=mid(QueryCheckString,1,len(QueryCheckString)-1) & ""&NAV="" & request.querystring(""NAV"")%>"" name=""QueryCheckString"">")
    Call fsoTextStreamAddPage.WriteLine("<p><a href=""getdata.asp?<%=mid(QueryCheckString,1,len(QueryCheckString)-1) & ""&NAV="" & request.querystring(""NAV"")%>"">Return to view data</p>")
    Call fsoTextStreamAddPage.WriteLine("</form>")
    Call fsoTextStreamAddPage.WriteLine("</BODY>")
    Call fsoTextStreamAddPage.WriteLine("</HTML>")
    Call fsoTextStreamAddPage.Close
    
    Set fsoTextStreamAddPage = fsoObject.CreateTextFile(iProjectTablePath & "\ASPupdate.asp")
    
    Call fsoTextStreamAddPage.WriteLine("<%@ Language=VBScript %>")
    Call fsoTextStreamAddPage.WriteLine("<%on error resume next%>")
    Call fsoTextStreamAddPage.WriteLine("<HTML>")
    Call fsoTextStreamAddPage.WriteLine("<HEAD>")
    Call fsoTextStreamAddPage.WriteLine("<TITLE>" & iTableName & " Update Data ASP Page</TITLE>")
    Call fsoTextStreamAddPage.WriteLine("</HEAD>")
    Call fsoTextStreamAddPage.WriteLine("<BODY>")
    Call fsoTextStreamAddPage.WriteLine("<% Set adoConnection = server.CreateObject(""ADODB.Connection"")")
    Call fsoTextStreamAddPage.WriteLine("Set adoRecordset = server.CreateObject(""ADODB.Recordset"")")
    Call fsoTextStreamAddPage.WriteLine("adoConnection.Provider = ""Microsoft.Jet.OLEDB.4.0""")
    Call fsoTextStreamAddPage.WriteLine("Dim strLocation, iLength")
    Call fsoTextStreamAddPage.WriteLine("strLocation = Request.ServerVariables(""PATH_TRANSLATED"")")
    Call fsoTextStreamAddPage.WriteLine("iLength = Len(strLocation)")
    Call fsoTextStreamAddPage.WriteLine("iLength = iLength - 13")
    Call fsoTextStreamAddPage.WriteLine("strLocation = Left(strLocation, iLength)")
    Call fsoTextStreamAddPage.WriteLine("strLocation = strLocation & ""../Database.mdb""")
    Call fsoTextStreamAddPage.WriteLine("adoConnection.Open (""Data Source="" & strLocation)")
    Call fsoTextStreamAddPage.WriteLine("QueryCheckString = request.form(""QueryCheckString"")")
    Call fsoTextStreamAddPage.WriteLine("SecondHalfSQL = request.form(""qString"")")
    Call fsoTextStreamAddPage.WriteLine("iFieldCount = 0")
    Call fsoTextStreamAddPage.WriteLine("FirstHalfSQL = ""update [" & iTableName & "] set ""%>")
    Call fsoTextStreamAddPage.WriteLine("<!-- LinkStarter. Do not modyify -->")
    Call fsoTextStreamAddPage.WriteLine("<% adoRecordset.ActiveConnection = AdoConnection")
    Call fsoTextStreamAddPage.WriteLine("SQLInsert = FirstHalfSQL & "" where "" & SecondHalfSQL")
    Call fsoTextStreamAddPage.WriteLine("call adoRecordset.Open(SQLInsert)")
    Call fsoTextStreamAddPage.WriteLine("response.redirect(""getdata.asp?"" & QueryCheckString)")
    Call fsoTextStreamAddPage.WriteLine("%>")
    Call fsoTextStreamAddPage.WriteLine("</BODY>")
    Call fsoTextStreamAddPage.WriteLine("</HTML>")
    Call fsoTextStreamAddPage.Close
    
    Set fsoTextStreamAddPage = Nothing
    
End Function

Private Function CreateAddTablePostHTML(iProjectRootPath, iProjectID, iTableName) As Boolean

    Dim DataType As Integer
    Dim Attrib As Long
    Dim FieldName As String
    Dim fsoTextStream As TextStream
    Dim fsoTextStreamTemp As TextStream
    Dim fsoObject As New FileSystemObject
    Dim CurrentLine As String
    Dim Required As String
    Dim StarRequired As String
    Dim PrimaryKey As Boolean
    Dim StarPrimary As String
    
    Attrib = arrProperties(1)
    DataType = arrProperties(2)
    FieldName = arrProperties(12)
    Required = arrProperties(8)
    PrimaryKey = arrProperties(13)
    
    Set fsoTextStream = fsoObject.OpenTextFile(iProjectRootPath & "\" & iTableName & "\add.asp", ForReading)
    Set fsoTextStreamTemp = fsoObject.CreateTextFile(iProjectRootPath & "\" & iTableName & "\addTemp.asp", True)
    
    While Not fsoTextStream.AtEndOfStream
        
        CurrentLine = fsoTextStream.ReadLine
        fsoTextStreamTemp.WriteLine CurrentLine
        
        If CurrentLine = "<!-- LinkStarter. Do not modyify -->" Then
            
            If Required = "True" Then
                StarRequired = "<font color=""#FF0000"">*</font>"
            Else
                StarRequired = ""
            End If
            
            If PrimaryKey = True Then
                StarPrimary = "<font color=""#008000"">*</font>"
            Else
                StarPrimary = ""
            End If
            
            Select Case DataType
                Case 10            '     Text
                    fsoTextStreamTemp.WriteLine ("<p>" & StarRequired & StarPrimary & FieldName & ":<br>")
                    Call fsoTextStreamTemp.WriteLine("<input type=""text"" name=""" & FieldName & """ size=""20""></p>")
                Case 12            '     Memo
                    Select Case Attrib
                        Case 2     '     memo
                            fsoTextStreamTemp.WriteLine ("<p>" & StarRequired & StarPrimary & FieldName & ":<br>")
                            Call fsoTextStreamTemp.WriteLine("<textarea rows=""4"" name=""" & FieldName & """ cols=""40""></textarea></p>")
                        Case 32770 'hyperlink
                            fsoTextStreamTemp.WriteLine ("<p>" & StarRequired & StarPrimary & FieldName & ":<br>")
                            Call fsoTextStreamTemp.WriteLine("<input type=""text"" name=""" & FieldName & """ size=""20""></p>")
                    End Select
                Case 3, 4, 2, 6, 7, 15, 20
                    Select Case Attrib
                        Case 17    'Autonumber
                                   'do nothing
                        Case Else     '    Number
                            fsoTextStreamTemp.WriteLine ("<p>" & StarRequired & StarPrimary & FieldName & ":<br>")
                            Call fsoTextStreamTemp.WriteLine("<input type=""text"" name=""" & FieldName & """ size=""20""></p>")
                    End Select
                Case 8             '  DateTime
                    fsoTextStreamTemp.WriteLine ("<p>" & StarRequired & StarPrimary & FieldName & ":<br>")
                    Call fsoTextStreamTemp.WriteLine("<input type=""text"" name=""" & FieldName & """ size=""20""></p>")
                Case 5             '  Currency
                    fsoTextStreamTemp.WriteLine ("<p>" & StarRequired & StarPrimary & FieldName & ":<br>")
                    Call fsoTextStreamTemp.WriteLine("<input type=""text"" name=""" & FieldName & """ size=""20""></p>")
                Case 1             '     YesNo
                    fsoTextStreamTemp.WriteLine ("<p>" & StarRequired & StarPrimary & FieldName & ":  ")
                    Call fsoTextStreamTemp.WriteLine("yes <input type=""radio"" value=""Yes"" checked name=""" & FieldName & """> no <input type=""radio"" name=""" & FieldName & """ value=""No""></p>")
                Case 11            'OLEOBject
                                   'is not supported
            End Select
        End If
    Wend
    
    Call fsoTextStream.Close
    Call fsoTextStreamTemp.Close
    
    Call fsoObject.DeleteFile(iProjectRootPath & "\" & iTableName & "\add.asp", True)
    Call fsoObject.MoveFile(iProjectRootPath & "\" & iTableName & "\addTemp.asp", iProjectRootPath & "\" & iTableName & "\add.asp")
    
    
End Function

Private Function CreateAddTableASP(iProjectRootPath, iProjectID, iTableName) As Boolean

    Dim DataType As Integer
    Dim Attrib As Long
    Dim FieldName As String
    Dim fsoTextStream As TextStream
    Dim fsoTextStreamTemp As TextStream
    Dim fsoObject As New FileSystemObject
    Dim CurrentLine As String
    Dim Required As String
    Dim Star As String
    Dim FieldNameVariable As String
    Dim PrimaryKey As Boolean
    
    Attrib = arrProperties(1)
    DataType = arrProperties(2)
    FieldName = arrProperties(12)
    Required = arrProperties(8)
    PrimaryKey = arrProperties(13)
    
    Set fsoTextStream = fsoObject.OpenTextFile(iProjectRootPath & "\" & iTableName & "\ASPadd.asp", ForReading)
    Set fsoTextStreamTemp = fsoObject.CreateTextFile(iProjectRootPath & "\" & iTableName & "\ASPaddTemp.asp", True)
    
    While Not fsoTextStream.AtEndOfStream
        
        CurrentLine = fsoTextStream.ReadLine
        
        fsoTextStreamTemp.WriteLine CurrentLine
        
        If CurrentLine = "<!-- LinkStarter. Do not modyify -->" And Attrib <> 17 And DataType <> 11 Then
            
            fsoTextStreamTemp.WriteLine ("<%")
            
            FieldNameVariable = Replace(FieldName, " ", "")
            FieldNameVariable = "var" & FieldNameVariable
            
            If Required = "True" Or PrimaryKey = True Then
                fsoTextStreamTemp.WriteLine ("If Request.Form(""" & FieldName & """) = """" Then")
                fsoTextStreamTemp.WriteLine ("Response.Redirect (""add.asp?fieldempty=" & FieldName & """)")
                fsoTextStreamTemp.WriteLine ("End If")
            End If
            
            Select Case DataType
                Case 10            '     Text
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                    fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                    fsoTextStreamTemp.WriteLine ("iFieldCount = iFieldCount + 1")
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = Replace(" & FieldNameVariable & ", ""'"", ""''"")")
                    
                    fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & ""[" & FieldName & "]""")
                    fsoTextStreamTemp.WriteLine ("Else")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "",[" & FieldName & "]""")
                    fsoTextStreamTemp.WriteLine ("End If")
                    
                    fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                    fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & ""'"" & " & FieldNameVariable & " & ""'""")
                    fsoTextStreamTemp.WriteLine ("Else")
                    fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & "",'"" & " & FieldNameVariable & " & ""'""")
                    fsoTextStreamTemp.WriteLine ("End If")
                    fsoTextStreamTemp.WriteLine ("End If")
                Case 12            '     Memo
                    Select Case Attrib
                        Case 2     '     memo
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                            fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                            fsoTextStreamTemp.WriteLine ("iFieldCount = iFieldCount + 1")
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = Replace(" & FieldNameVariable & ", ""'"", ""''"")")
                    
                            fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                            fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & ""[" & FieldName & "]""")
                            fsoTextStreamTemp.WriteLine ("Else")
                            fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "",[" & FieldName & "]""")
                            fsoTextStreamTemp.WriteLine ("End If")
                    
                            fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                            fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & ""'"" & " & FieldNameVariable & " & ""'""")
                            fsoTextStreamTemp.WriteLine ("Else")
                            fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & "",'"" & " & FieldNameVariable & " & ""'""")
                            fsoTextStreamTemp.WriteLine ("End If")
                            fsoTextStreamTemp.WriteLine ("End If")
                        
                        Case 32770 'hyperlink
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                            fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                            fsoTextStreamTemp.WriteLine ("iFieldCount = iFieldCount + 1")
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = Replace(" & FieldNameVariable & ", ""'"", ""''"")")
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = " & FieldNameVariable & " & ""#"" & " & FieldNameVariable & " & ""#""")
                            
                            fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                            fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & ""[" & FieldName & "]""")
                            fsoTextStreamTemp.WriteLine ("Else")
                            fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "",[" & FieldName & "]""")
                            fsoTextStreamTemp.WriteLine ("End If")
                    
                            fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                            fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & ""'"" & " & FieldNameVariable & " & ""'""")
                            fsoTextStreamTemp.WriteLine ("Else")
                            fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & "",'"" & " & FieldNameVariable & " & ""'""")
                            fsoTextStreamTemp.WriteLine ("End If")
                            fsoTextStreamTemp.WriteLine ("End If")
                    End Select
                Case 3, 4, 2, 6, 7, 15, 20
                    Select Case Attrib
                        Case 17    'Autonumber
                                   'do nothing
                        Case Else     '    Number
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                            fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                            fsoTextStreamTemp.WriteLine ("iFieldCount = iFieldCount + 1")
                            fsoTextStreamTemp.WriteLine ("On Error Resume Next")
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = CDbl(" & FieldNameVariable & ")")
                            fsoTextStreamTemp.WriteLine ("If Err.Number = 13 Then")
                            fsoTextStreamTemp.WriteLine ("response.redirect (""add.asp?invaliddata=" & FieldName & """)")
                            fsoTextStreamTemp.WriteLine ("End If")
                            fsoTextStreamTemp.WriteLine ("Err.Clear")
                            fsoTextStreamTemp.WriteLine ("On Error GoTo 0")

                            fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                            fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & ""[" & FieldName & "]""")
                            fsoTextStreamTemp.WriteLine ("Else")
                            fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "",[" & FieldName & "]""")
                            fsoTextStreamTemp.WriteLine ("End If")
                    
                            fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                            fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & " & FieldNameVariable)
                            fsoTextStreamTemp.WriteLine ("Else")
                            fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & "","" & " & FieldNameVariable)
                            fsoTextStreamTemp.WriteLine ("End If")
                            fsoTextStreamTemp.WriteLine ("End If")
                    End Select
                Case 8             '  DateTime
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                    fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                    fsoTextStreamTemp.WriteLine ("iFieldCount = iFieldCount + 1")
                    fsoTextStreamTemp.WriteLine ("On Error Resume Next")
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = CDate(" & FieldNameVariable & ")")
                    fsoTextStreamTemp.WriteLine ("If Err.Number = 13 Then")
                    fsoTextStreamTemp.WriteLine ("response.redirect (""add.asp?invaliddata=" & FieldName & """)")
                    fsoTextStreamTemp.WriteLine ("End If")
                    fsoTextStreamTemp.WriteLine ("Err.Clear")
                    fsoTextStreamTemp.WriteLine ("On Error GoTo 0")
                    
                    fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & ""[" & FieldName & "]""")
                    fsoTextStreamTemp.WriteLine ("Else")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "",[" & FieldName & "]""")
                    fsoTextStreamTemp.WriteLine ("End If")
                    
                    fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                    fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & ""'"" & " & FieldNameVariable & " & ""'""")
                    fsoTextStreamTemp.WriteLine ("Else")
                    fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & "",'"" & " & FieldNameVariable & " & ""'""")
                    fsoTextStreamTemp.WriteLine ("End If")
                    fsoTextStreamTemp.WriteLine ("End If")
                Case 5             '  Currency
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                    fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                    fsoTextStreamTemp.WriteLine ("iFieldCount = iFieldCount + 1")
                    fsoTextStreamTemp.WriteLine ("On Error Resume Next")
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = CCur(" & FieldNameVariable & ")")
                    fsoTextStreamTemp.WriteLine ("If Err.Number = 13 Then")
                    fsoTextStreamTemp.WriteLine ("response.redirect (""add.asp?invaliddata=" & FieldName & """)")
                    fsoTextStreamTemp.WriteLine ("End If")
                    fsoTextStreamTemp.WriteLine ("Err.Clear")
                    fsoTextStreamTemp.WriteLine ("On Error GoTo 0")
                    
                    fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & ""[" & FieldName & "]""")
                    fsoTextStreamTemp.WriteLine ("Else")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "",[" & FieldName & "]""")
                    fsoTextStreamTemp.WriteLine ("End If")
                    
                    fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                    fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & " & FieldNameVariable)
                    fsoTextStreamTemp.WriteLine ("Else")
                    fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & "","" & " & FieldNameVariable)
                    fsoTextStreamTemp.WriteLine ("End If")
                    fsoTextStreamTemp.WriteLine ("End If")
                Case 1             '     YesNo
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                    fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                    fsoTextStreamTemp.WriteLine ("iFieldCount = iFieldCount + 1")
                    fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " = ""Yes"" then")
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = True")
                    fsoTextStreamTemp.WriteLine ("Else")
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = False")
                    fsoTextStreamTemp.WriteLine ("End If")
                    
                    fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & ""[" & FieldName & "]""")
                    fsoTextStreamTemp.WriteLine ("Else")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "",[" & FieldName & "]""")
                    fsoTextStreamTemp.WriteLine ("End If")
                    
                    fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                    fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & " & FieldNameVariable)
                    fsoTextStreamTemp.WriteLine ("Else")
                    fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & "","" & " & FieldNameVariable)
                    fsoTextStreamTemp.WriteLine ("End If")
                    fsoTextStreamTemp.WriteLine ("End If")
                Case 11            'OLEOBject
                                   'is not supported
            End Select
            
            fsoTextStreamTemp.WriteLine ("%>")
        End If
    Wend
    
    Call fsoTextStream.Close
    Call fsoTextStreamTemp.Close
    
    Call fsoObject.DeleteFile(iProjectRootPath & "\" & iTableName & "\ASPadd.asp", True)
    Call fsoObject.MoveFile(iProjectRootPath & "\" & iTableName & "\ASPaddTemp.asp", iProjectRootPath & "\" & iTableName & "\ASPadd.asp")
    
End Function

Private Function CreateViewTable(iProjectRootPath, iProjectID, iTableName) As Boolean

    Dim DataType As Integer
    Dim Attrib As Long
    Dim FieldName As String
    Dim fsoTextStream As TextStream
    Dim fsoTextStreamTemp As TextStream
    Dim fsoObject As New FileSystemObject
    Dim CurrentLine As String
    Dim Required As String
    Dim Star As String
    
    Attrib = arrProperties(1)
    DataType = arrProperties(2)
    FieldName = arrProperties(12)
    Required = arrProperties(8)
    
    Set fsoTextStream = fsoObject.OpenTextFile(iProjectRootPath & "\" & iTableName & "\view.asp", ForReading)
    Set fsoTextStreamTemp = fsoObject.CreateTextFile(iProjectRootPath & "\" & iTableName & "\viewtemp.asp", True)
    
    While Not fsoTextStream.AtEndOfStream
        
        CurrentLine = fsoTextStream.ReadLine
        fsoTextStreamTemp.WriteLine CurrentLine
        
        If CurrentLine = "<!-- LinkStarter. Do not modyify -->" Then
            If DataType <> 11 Then
                fsoTextStreamTemp.WriteLine ("<p><input type=""checkbox"" name=""" & FieldName & """ value=""ON""> " & FieldName & "</p>")
            End If
        End If
    Wend
    
    Call fsoTextStream.Close
    Call fsoTextStreamTemp.Close
    
    Call fsoObject.DeleteFile(iProjectRootPath & "\" & iTableName & "\view.asp", True)
    Call fsoObject.MoveFile(iProjectRootPath & "\" & iTableName & "\viewtemp.asp", iProjectRootPath & "\" & iTableName & "\view.asp")
    
End Function

Private Function CreateGetTableData(iProjectRootPath, iProjectID, iTableName) As Boolean

    Dim DataType As Integer
    Dim Attrib As Long
    Dim FieldName As String
    Dim fsoTextStream As TextStream
    Dim fsoTextStreamTemp As TextStream
    Dim fsoObject As New FileSystemObject
    Dim CurrentLine As String
    Dim Required As String
    Dim Star As String
    Dim FieldNameVariable As String
    
    Attrib = arrProperties(1)
    DataType = arrProperties(2)
    FieldName = arrProperties(12)
    Required = arrProperties(8)
    
    Set fsoTextStream = fsoObject.OpenTextFile(iProjectRootPath & "\" & iTableName & "\getdata.asp", ForReading)
    Set fsoTextStreamTemp = fsoObject.CreateTextFile(iProjectRootPath & "\" & iTableName & "\getdatatemp.asp", True)
    
    While Not fsoTextStream.AtEndOfStream
        
        CurrentLine = fsoTextStream.ReadLine
        fsoTextStreamTemp.WriteLine CurrentLine
        
        If CurrentLine = "<!-- LinkStarter. Do not modyify -->" Then
            
            If DataType <> 11 Then
                FieldNameVariable = Replace(FieldName, " ", "")
                FieldNameVariable = "var" & FieldNameVariable
                
                fsoTextStreamTemp.WriteLine ("<%")
                fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.querystring(""" & FieldName & """)")
                fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " = ""ON"" then")
                fsoTextStreamTemp.WriteLine ("ReDim Preserve arrField(totFieldCount)")
                fsoTextStreamTemp.WriteLine ("arrField(totFieldCount) = """ & FieldName & """")
                fsoTextStreamTemp.WriteLine ("totFieldCount = totFieldCount + 1")
                fsoTextStreamTemp.WriteLine ("End If")
                fsoTextStreamTemp.WriteLine ("%>")
            End If
        End If
    Wend
    
    Call fsoTextStream.Close
    Call fsoTextStreamTemp.Close
    
    Call fsoObject.DeleteFile(iProjectRootPath & "\" & iTableName & "\getdata.asp", True)
    Call fsoObject.MoveFile(iProjectRootPath & "\" & iTableName & "\getdatatemp.asp", iProjectRootPath & "\" & iTableName & "\getdata.asp")
    
End Function

Private Function CreateDeleteTable(iProjectRootPath, iProjectID, iTableName) As Boolean

    Dim DataType As Integer
    Dim Attrib As Long
    Dim FieldName As String
    Dim fsoTextStream As TextStream
    Dim fsoTextStreamTemp As TextStream
    Dim fsoObject As New FileSystemObject
    Dim CurrentLine As String
    Dim Required As String
    Dim Star As String
    Dim FieldNameVariable As String
    Dim QueryCheck As String
    
    Attrib = arrProperties(1)
    DataType = arrProperties(2)
    FieldName = arrProperties(12)
    Required = arrProperties(8)
    
    Set fsoTextStream = fsoObject.OpenTextFile(iProjectRootPath & "\" & iTableName & "\delete.asp", ForReading)
    Set fsoTextStreamTemp = fsoObject.CreateTextFile(iProjectRootPath & "\" & iTableName & "\deletetemp.asp", True)
    
    While Not fsoTextStream.AtEndOfStream
        
        CurrentLine = fsoTextStream.ReadLine
        fsoTextStreamTemp.WriteLine CurrentLine
        
        If CurrentLine = "<!-- LinkStarter. Do not modyify -->" Then
            
            FieldNameVariable = Replace(FieldName, " ", "")
            FieldNameVariable = "var" & FieldNameVariable
                    
            QueryCheck = "Q" & FieldNameVariable
            
            fsoTextStreamTemp.WriteLine ("<%")
            
            fsoTextStreamTemp.WriteLine (QueryCheck & " = request.querystring(""" & FieldName & """)")
            fsoTextStreamTemp.WriteLine ("If " & QueryCheck & " = ""ON"" then")
            fsoTextStreamTemp.WriteLine ("QueryCheckString = QueryCheckString & """ & FieldName & """ & ""=ON&""")
            fsoTextStreamTemp.WriteLine ("end if")
            
            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.querystring(""dat" & FieldName & """)")
            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = replace(" & FieldNameVariable & ", ""'"", ""''"")")
            fsoTextStreamTemp.WriteLine ("ReDim Preserve arrField(totFieldCount)")
            fsoTextStreamTemp.WriteLine ("arrField(totFieldCount) = """ & FieldName & """")
            fsoTextStreamTemp.WriteLine ("if " & FieldNameVariable & " <> """" then")
            Select Case DataType
                Case 10, 12        'Text, Memo, Hyperlink
                    
                    fsoTextStreamTemp.WriteLine ("if totfieldcount <> 0 then qString = qString & "" and""")
                    fsoTextStreamTemp.WriteLine ("qString = qString & "" [" & FieldName & "] = '"" & " & FieldNameVariable & " & ""'""")
                    fsoTextStreamTemp.WriteLine ("totFieldCount = totFieldCount + 1")
                Case 3, 4, 2, 6, 7, 15, 20, 5, 1       'Number, Currency, YesNo
                    fsoTextStreamTemp.WriteLine ("if totfieldcount <> 0 then qString = qString & "" and""")
                    fsoTextStreamTemp.WriteLine ("qString = qString & "" [" & FieldName & "] = "" & " & FieldNameVariable & "")
                    fsoTextStreamTemp.WriteLine ("totFieldCount = totFieldCount + 1")
                Case 8             'Date
                    fsoTextStreamTemp.WriteLine ("if totfieldcount <> 0 then qString = qString & "" and""")
                    fsoTextStreamTemp.WriteLine ("qString = qString & "" [" & FieldName & "] = cdate('"" & " & FieldNameVariable & " & ""')""")
                    fsoTextStreamTemp.WriteLine ("totFieldCount = totFieldCount + 1")

                Case 11            'OLEOBject
                                   'is not supported
            End Select
            fsoTextStreamTemp.WriteLine ("end if")
            
            fsoTextStreamTemp.WriteLine ("%>")

        End If
    Wend
    
    Call fsoTextStream.Close
    Call fsoTextStreamTemp.Close
    
    Call fsoObject.DeleteFile(iProjectRootPath & "\" & iTableName & "\delete.asp", True)
    Call fsoObject.MoveFile(iProjectRootPath & "\" & iTableName & "\deletetemp.asp", iProjectRootPath & "\" & iTableName & "\delete.asp")
    
End Function

Private Function CreateUpdateTablePostHTML(iProjectRootPath, iProjectID, iTableName) As Boolean

    Dim DataType As Integer
    Dim Attrib As Long
    Dim FieldName As String
    Dim fsoTextStream As TextStream
    Dim fsoTextStreamTemp As TextStream
    Dim fsoObject As New FileSystemObject
    Dim CurrentLine As String
    Dim Required As String
    Dim StarRequired As String
    Dim PrimaryKey As Boolean
    Dim StarPrimary As String
    Dim valueFieldName As String
    Dim QueryCheck As String
    Dim FieldNameVariable As String
    
    Attrib = arrProperties(1)
    DataType = arrProperties(2)
    FieldName = arrProperties(12)
    Required = arrProperties(8)
    PrimaryKey = arrProperties(13)
    
    Set fsoTextStream = fsoObject.OpenTextFile(iProjectRootPath & "\" & iTableName & "\update.asp", ForReading)
    Set fsoTextStreamTemp = fsoObject.CreateTextFile(iProjectRootPath & "\" & iTableName & "\updateTemp.asp", True)
    
    While Not fsoTextStream.AtEndOfStream
        
        CurrentLine = fsoTextStream.ReadLine
        fsoTextStreamTemp.WriteLine CurrentLine
        
        
        If CurrentLine = "<!-- LinkStarter. Do not modyify -->" Then
            
            valueFieldName = Replace(FieldName, " ", "")
            valueFieldName = "dat" & valueFieldName
            
            FieldNameVariable = Replace(FieldName, " ", "")
            FieldNameVariable = "var" & FieldNameVariable
            QueryCheck = "Q" & valueFieldName
            
            If Required = "True" Then
                StarRequired = "<font color=""#FF0000"">*</font>"
            Else
                StarRequired = ""
            End If
            
            If PrimaryKey = True Then
                StarPrimary = "<font color=""#008000"">*</font>"
            Else
                StarPrimary = ""
            End If
             
            fsoTextStreamTemp.WriteLine ("<%")
            fsoTextStreamTemp.WriteLine (QueryCheck & " = request.querystring(""" & FieldName & """)")
            fsoTextStreamTemp.WriteLine ("If " & QueryCheck & " = ""ON"" then")
            fsoTextStreamTemp.WriteLine ("QueryCheckString = QueryCheckString & """ & FieldName & """ & ""=ON&""")
            fsoTextStreamTemp.WriteLine ("end if")
            fsoTextStreamTemp.WriteLine ("%>")
            
            Select Case DataType
                Case 10, 12        'Text, Memo, Hyperlink
                    fsoTextStreamTemp.WriteLine ("<%if request.querystring (""" & valueFieldName & """) <> """" then ")
                    fsoTextStreamTemp.WriteLine ("if qString <> """" then qString = qString & "" and""")
                    fsoTextStreamTemp.WriteLine ("qString = qString & "" [" & FieldName & "] = '"" & request.querystring(""" & valueFieldName & """) & ""'""")
                    fsoTextStreamTemp.WriteLine ("end if%>")
                Case 3, 4, 2, 6, 7, 15, 20, 5, 1       'Number, Currency, YesNo
                    fsoTextStreamTemp.WriteLine ("<%if request.querystring (""" & valueFieldName & """) <> """" then ")
                    fsoTextStreamTemp.WriteLine ("if qString <> """" then qString = qString & "" and""")
                    fsoTextStreamTemp.WriteLine ("qString = qString & "" [" & FieldName & "] = "" & request.querystring(""" & valueFieldName & """)")
                    fsoTextStreamTemp.WriteLine ("end if%>")
                Case 8             'Date
                    fsoTextStreamTemp.WriteLine ("<%if request.querystring (""" & valueFieldName & """) <> """" then ")
                    fsoTextStreamTemp.WriteLine ("if qString <> """" then qString = qString & "" and""")
                    fsoTextStreamTemp.WriteLine ("qString = qString & "" [" & FieldName & "] = cdate('"" & request.querystring(""" & valueFieldName & """) & ""')""")
                    fsoTextStreamTemp.WriteLine ("end if%>")
                Case 11            'OLEOBject
                                   'is not supported
            End Select
            
            Select Case DataType
                Case 10            '     Text
                    Call fsoTextStreamTemp.WriteLine("<p>" & StarRequired & StarPrimary & FieldName & ":<br>")
                    Call fsoTextStreamTemp.WriteLine("<input type=""text"" name=""" & FieldName & """ size=""20"" value=""<%=request.querystring(""" & valueFieldName & """)%>""></p>")
                Case 12            '     Memo
                    Select Case Attrib
                        Case 2     '     memo
                            Call fsoTextStreamTemp.WriteLine("<p>" & StarRequired & StarPrimary & FieldName & ":<br>")
                            Call fsoTextStreamTemp.WriteLine("<textarea rows=""4"" name=""" & FieldName & """ cols=""40""><%=request.querystring(""" & valueFieldName & """)%></textarea></p>")
                        Case 32770 'hyperlink
                            Call fsoTextStreamTemp.WriteLine("<p>" & StarRequired & StarPrimary & FieldName & ":<br>")
                            Call fsoTextStreamTemp.WriteLine("<%if request.querystring(""" & valueFieldName & """) <> """" then%>")
                            Call fsoTextStreamTemp.WriteLine("<input type=""text"" name=""" & FieldName & """ size=""20"" value=""<%=mid(request.querystring(""" & valueFieldName & """),1,(len(request.querystring(""" & valueFieldName & """))/2)-1)%>""></p>")
                            Call fsoTextStreamTemp.WriteLine("<%else%>")
                            Call fsoTextStreamTemp.WriteLine("<input type=""text"" name=""" & FieldName & """ size=""20""></p>")
                            Call fsoTextStreamTemp.WriteLine("<%end if%>")
                    End Select
                Case 3, 4, 2, 6, 7, 15, 20
                    Select Case Attrib
                        Case 17    'Autonumber
                                   'do nothing
                        Case Else     '    Number
                            Call fsoTextStreamTemp.WriteLine("<p>" & StarRequired & StarPrimary & FieldName & ":<br>")
                            Call fsoTextStreamTemp.WriteLine("<input type=""text"" name=""" & FieldName & """ size=""20"" value=""<%=request.querystring(""" & valueFieldName & """)%>""></p>")
                    End Select
                Case 8             '  DateTime
                    Call fsoTextStreamTemp.WriteLine("<p>" & StarRequired & StarPrimary & FieldName & ":<br>")
                    Call fsoTextStreamTemp.WriteLine("<input type=""text"" name=""" & FieldName & """ size=""20"" value=""<%=request.querystring(""" & valueFieldName & """)%>""></p>")
                Case 5             '  Currency
                    Call fsoTextStreamTemp.WriteLine("<p>" & StarRequired & StarPrimary & FieldName & ":<br>")
                    Call fsoTextStreamTemp.WriteLine("<input type=""text"" name=""" & FieldName & """ size=""20"" value=""<%=request.querystring(""" & valueFieldName & """)%>""></p>")
                Case 1             '     YesNo
                    
                    Call fsoTextStreamTemp.WriteLine("<p>" & StarRequired & StarPrimary & FieldName & ":  ")
                    Call fsoTextStreamTemp.WriteLine("<%if request.querystring(""" & valueFieldName & """) <> ""False"" then%>")
                    Call fsoTextStreamTemp.WriteLine("yes <input type=""radio"" value=""Yes"" checked name=""" & FieldName & """> no <input type=""radio"" name=""" & FieldName & """ value=""No""></p>")
                    Call fsoTextStreamTemp.WriteLine("<%else%>")
                    Call fsoTextStreamTemp.WriteLine("yes <input type=""radio"" value=""Yes"" name=""" & FieldName & """> no <input type=""radio"" name=""" & FieldName & """ value=""No"" checked></p>")
                    Call fsoTextStreamTemp.WriteLine("<%end if%>")
                Case 11            'OLEOBject
                                   'is not supported
            End Select
        End If
    Wend
    
    Call fsoTextStream.Close
    Call fsoTextStreamTemp.Close
    
    Call fsoObject.DeleteFile(iProjectRootPath & "\" & iTableName & "\update.asp", True)
    Call fsoObject.MoveFile(iProjectRootPath & "\" & iTableName & "\updateTemp.asp", iProjectRootPath & "\" & iTableName & "\update.asp")
    
    
End Function

Private Function CreateUpdateTableASP(iProjectRootPath, iProjectID, iTableName) As Boolean

    Dim DataType As Integer
    Dim Attrib As Long
    Dim FieldName As String
    Dim fsoTextStream As TextStream
    Dim fsoTextStreamTemp As TextStream
    Dim fsoObject As New FileSystemObject
    Dim CurrentLine As String
    Dim Required As String
    Dim Star As String
    Dim FieldNameVariable As String
    Dim PrimaryKey As Boolean
    
    Attrib = arrProperties(1)
    DataType = arrProperties(2)
    FieldName = arrProperties(12)
    Required = arrProperties(8)
    PrimaryKey = arrProperties(13)
    
    Set fsoTextStream = fsoObject.OpenTextFile(iProjectRootPath & "\" & iTableName & "\ASPupdate.asp", ForReading)
    Set fsoTextStreamTemp = fsoObject.CreateTextFile(iProjectRootPath & "\" & iTableName & "\ASPupdateTemp.asp", True)
    
    While Not fsoTextStream.AtEndOfStream
        
        CurrentLine = fsoTextStream.ReadLine
        
        fsoTextStreamTemp.WriteLine CurrentLine
        
        If CurrentLine = "<!-- LinkStarter. Do not modyify -->" And Attrib <> 17 And DataType <> 11 Then
            
            FieldNameVariable = Replace(FieldName, " ", "")
            FieldNameVariable = "var" & FieldNameVariable
            
            fsoTextStreamTemp.WriteLine ("<%")
            
            If Required = "True" Or PrimaryKey = True Then
                fsoTextStreamTemp.WriteLine ("If Request.Form(""" & FieldName & """) = """" Then")
                fsoTextStreamTemp.WriteLine ("Response.Redirect (""update.asp?"" & request.querystring & ""&fieldempty=" & FieldName & """)")
                fsoTextStreamTemp.WriteLine ("End If")
            End If
            
            Select Case DataType
                Case 10            '     Text
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = replace(" & FieldNameVariable & ", ""'"", ""''"")")
                    fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = Replace(" & FieldNameVariable & ", ""'"", ""''"")")
                    
                    fsoTextStreamTemp.WriteLine ("if FirstHalfSQL = ""update [" & iTableName & "] set "" then")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "" [" & FieldName & "] = '"" & " & FieldNameVariable & " & ""'""")
                    fsoTextStreamTemp.WriteLine ("else")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "", [" & FieldName & "] = '"" & " & FieldNameVariable & " & ""'""")
                    fsoTextStreamTemp.WriteLine ("end if")
                    
                    fsoTextStreamTemp.WriteLine ("end if")
                Case 12            '     Memo
                    Select Case Attrib
                        Case 2     '     memo
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                            fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = Replace(" & FieldNameVariable & ", ""'"", ""''"")")
                            
                            fsoTextStreamTemp.WriteLine ("if FirstHalfSQL = ""update [" & iTableName & "] set "" then")
                            fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "" [" & FieldName & "] = '"" & " & FieldNameVariable & " & ""'""")
                            fsoTextStreamTemp.WriteLine ("else")
                            fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "", [" & FieldName & "] = '"" & " & FieldNameVariable & " & ""'""")
                            fsoTextStreamTemp.WriteLine ("end if")
                            
                            fsoTextStreamTemp.WriteLine ("end if")
                        Case 32770 'hyperlink
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                            fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = Replace(" & FieldNameVariable & ", ""'"", ""''"")")
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = " & FieldNameVariable & " & ""#"" & " & FieldNameVariable & " & ""#""")
                            
                            fsoTextStreamTemp.WriteLine ("if FirstHalfSQL = ""update [" & iTableName & "] set "" then")
                            fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "" [" & FieldName & "] = '"" & " & FieldNameVariable & " & ""'""")
                            fsoTextStreamTemp.WriteLine ("else")
                            fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "", [" & FieldName & "] = '"" & " & FieldNameVariable & " & ""'""")
                            fsoTextStreamTemp.WriteLine ("end if")
                            
                            fsoTextStreamTemp.WriteLine ("end if")
                    End Select
                Case 3, 4, 2, 6, 7, 15, 20
                    Select Case Attrib
                        Case 17    'Autonumber
                                   'do nothing
                        Case Else     '    Number
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                            fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                            fsoTextStreamTemp.WriteLine ("On Error Resume Next")
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = CDbl(" & FieldNameVariable & ")")
                            fsoTextStreamTemp.WriteLine ("If Err.Number = 13 Then")
                            fsoTextStreamTemp.WriteLine ("response.redirect (""getdata.asp?"" & QueryCheckString)")
                            fsoTextStreamTemp.WriteLine ("End If")
                            fsoTextStreamTemp.WriteLine ("Err.Clear")
                            fsoTextStreamTemp.WriteLine ("On Error GoTo 0")
                            
                            fsoTextStreamTemp.WriteLine ("if FirstHalfSQL = ""update [" & iTableName & "] set "" then")
                            fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "" [" & FieldName & "] = "" & " & FieldNameVariable)
                            fsoTextStreamTemp.WriteLine ("else")
                            fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "", [" & FieldName & "] = "" & " & FieldNameVariable)
                            fsoTextStreamTemp.WriteLine ("end if")
                            
                            fsoTextStreamTemp.WriteLine ("end if")
                    End Select
                Case 8             '  DateTime
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                    fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                    fsoTextStreamTemp.WriteLine ("On Error Resume Next")
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = CDate(" & FieldNameVariable & ")")
                    fsoTextStreamTemp.WriteLine ("If Err.Number = 13 Then")
                    fsoTextStreamTemp.WriteLine ("response.redirect (""getdata.asp?"" & QueryCheckString)")
                    fsoTextStreamTemp.WriteLine ("End If")
                    fsoTextStreamTemp.WriteLine ("Err.Clear")
                    fsoTextStreamTemp.WriteLine ("On Error GoTo 0")
                    
                    fsoTextStreamTemp.WriteLine ("if FirstHalfSQL = ""update [" & iTableName & "] set "" then")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "" [" & FieldName & "] = cdate('"" & " & FieldNameVariable & " & ""')""")
                    fsoTextStreamTemp.WriteLine ("else")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "", [" & FieldName & "] = cdate('"" & " & FieldNameVariable & " & ""')""")
                    fsoTextStreamTemp.WriteLine ("end if")
                    
                    fsoTextStreamTemp.WriteLine ("end if")
                    
                Case 5             '  Currency
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                    fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                    fsoTextStreamTemp.WriteLine ("On Error Resume Next")
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = CCur(" & FieldNameVariable & ")")
                    fsoTextStreamTemp.WriteLine ("If Err.Number = 13 Then")
                    fsoTextStreamTemp.WriteLine ("response.redirect (""getdata.asp?"" & QueryCheckString)")
                    fsoTextStreamTemp.WriteLine ("End If")
                    fsoTextStreamTemp.WriteLine ("Err.Clear")
                    fsoTextStreamTemp.WriteLine ("On Error GoTo 0")
                    
                    fsoTextStreamTemp.WriteLine ("if FirstHalfSQL = ""update [" & iTableName & "] set "" then")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "" [" & FieldName & "] = "" & " & FieldNameVariable)
                    fsoTextStreamTemp.WriteLine ("else")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "", [" & FieldName & "] = "" & " & FieldNameVariable)
                    fsoTextStreamTemp.WriteLine ("end if")
                    
                    fsoTextStreamTemp.WriteLine ("end if")
                Case 1             '     YesNo
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                    fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " = ""Yes"" then")
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = True")
                    fsoTextStreamTemp.WriteLine ("Else")
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = False")
                    fsoTextStreamTemp.WriteLine ("End If")
                     
                    fsoTextStreamTemp.WriteLine ("if FirstHalfSQL = ""update [" & iTableName & "] set "" then")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "" [" & FieldName & "] = "" & " & FieldNameVariable)
                    fsoTextStreamTemp.WriteLine ("else")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "", [" & FieldName & "] = "" & " & FieldNameVariable)
                    fsoTextStreamTemp.WriteLine ("end if")
                
                Case 11            'OLEOBject
                                   'is not supported
            End Select
            
            fsoTextStreamTemp.WriteLine ("%>")
        End If
    Wend
    
    Call fsoTextStream.Close
    Call fsoTextStreamTemp.Close
    
    Call fsoObject.DeleteFile(iProjectRootPath & "\" & iTableName & "\ASPupdate.asp", True)
    Call fsoObject.MoveFile(iProjectRootPath & "\" & iTableName & "\ASPupdateTemp.asp", iProjectRootPath & "\" & iTableName & "\ASPupdate.asp")
    
End Function

Private Function CheckGoodDatabase(idbPath) As Boolean

    'function checks to see if database is valid by
    'making a connection to it and then disconnecting
    
    Dim objDAO As New DAO.DBEngine
    Dim objDatabase As DAO.Database
    Dim DBPassword As String
    Dim wrkJet As DAO.Workspace
    
    On Error GoTo errorhandler:
    
    Set objDatabase = objDAO.OpenDatabase(idbPath)
    
    CheckGoodDatabase = True
    
    Exit Function
    
errorhandler:
    CheckGoodDatabase = False
    
End Function

Private Sub Command1_Click()
    
    On Error GoTo errorhandler
    
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "Access Database (*.mdb)|*.mdb"
    CommonDialog1.ShowOpen
    Command1.Enabled = False
    DoEvents
    If CreateProject(CommonDialog1.FileName, txtProjectPath, "test") = True Then
        MsgBox ("Project created successfully")
    Else
        MsgBox ("Unable to create project")
    End If
    Command1.Enabled = True

errorhandler:

End Sub

