'-- MSI database class. This class is a VBScript class for working with MSI database files.
   '-- datatypes: As written this database class can save strings to 255 characters and integers.
   
   '-- ====================================
   '-- NOTE: This class includes error trapping and error information feedback for common errors,
   '-  BUT IT DOES NOT DEAL WITH THE ERROR OF NO LOADED DATABASE. YOU MUST GET A
   '-- VALID RETURN OF 0 FROM LoadDB FUNCTION BEFOR CALLING ANY OTHER FUNCTIONS.
   '-- ===================================================================
   
'-- To create class object:  Set MB = New MBase

'-- See help file (MBase Functions.html) for detailed information.
'-- When functions fail use ErrorString property to return information about what went wrong.

'------------------------------------------------
'-Function LoadDB(MSIFilePath)
'-Sub ReLoadDB
'-Sub UnLoadDB
'-Function MakeNewDB(Path)
'-Sub ClearSummary
'-Sub SetSummaryValues(sTitle, sSubject, sAuthor, sNotes)
'-Sub ClearAllTables
'-Sub AddTable(TableName, ColumnString)
'-Sub AddTableCustom(TableName, ColumnString)
'-Sub ImportTable(Path)
'-Sub RemoveTable(TableName)
'-Function AddColumn(TableName, ColName, IfNumeric)
'-Function AddRecord(TableName, Values)
'-Function RemoveRecord(TableName, ColName, Value)
'-Function GetSelectValues(TableName, ColumnName, ColumnToCompare, ValueToCompare)
'-Function GetAllValues(TableName, ColumnName)
'-Function SetSelectValues(TableName, ColumnToSet, ValueToSet, ColumnToCompare, ValueToCompare)
'-Function SetAllValues(TableName, ColumnToSet, ValueToSet)
'-Function GetTableContent(TableName)
'-Function GetTableChart
'-Function ExportAllTables(FolderPath)
'-Function ExportAllValidTables(FolderPath)
'-Function LoadFindData()
'-function Find1(StringTofind)
'-Property ErrorString
'-Function GetTableNames(ReturnArray)
'-Function GetTableColumns(TableName, ReturnArray)
'-Function GetTableColumnTypes(TableName, ReturnArray)

'-- new 5-07:
'-Function GetTableColumnSpecificTypes(TableName, ReturnArray)
'-Function GetColumnDataType(TableName, ColumnName, ReturnArray)

  '-- new functions 5-27-06 -----------------
'-Function GetCABList() - return array. array(0) = CStr(number of cabs) . array(1 to (number of cabs)) are cab names.
'-Function ExtractCAB(CabName) - return 0 on success, or error code.
'-Function InsertCAB(CabName, NewCABPath) - return 0 on success. CabName is name in _streams table. NewCABPath is path of CAB being inserted.

'--===============================================================   
  
'-- 8888888888888888888888888888888888888888888888888888888888888888888888888888888888

Class MBase
  Private MSIFile  '-- path of currently loaded file.
  Private MSIFolder  '-- path of loaded file's parent folder.
  Private TabFol  '-- path of folder in TEMP folder for exported tables.
  Private BooLoaded
  Private cApos, cCar
  Private DB, WI, View, Rec, sErr, DicTables, DicTables2, DicTableCType, ATables
  Private DicFind, FindReady   '-- dictionary to hold text of tables. FindReady is boolean.

'------------------------------- Load an MSI database file ------------------------------------
'-- Return: 0-success. 1-invalid path. Return error code if WI fails or if Read/Write access fails.
Public Function LoadDB(sPath)
  Dim FSO1, Boo1, sTable, i3, i4, ATab(), ATab2(), ATabT(), iCnt, s2, iTyp, BooValid, Pt1
    sErr = ""  '-- clear ErrorString.
    UnLoadDB  '-- clear any objects left over from last DB load.
    LoadDB = 1
  Set FSO1 = CreateObject("Scripting.FileSystemObject")
     Boo1 = FSO1.FileExists(sPath) 
  Set FSO1 = Nothing
  
     If (Boo1 = False) Then 
        sErr = "LoadDB: Invalid file path."
        Exit Function  '-- error 1. invalid file path.
     End If    
   Err.Clear
   On Error Resume Next
  Set WI = CreateObject("WindowsInstaller.Installer")
  Set DB = WI.OpenDatabase(sPath, 1)
     If (Err.Number <> 0) Then
        sErr = "LoadDB: " & Err.Description
        LoadDB = Err.Number
        Exit Function     
     End If
  
  '-- two dictionaries to keep track of DB layout. DicTables contains keys with table names. Items are
  '-- arrays with Column names for that table, starting at 1. So ubound is number of columns.
  '-- DicTables2 is also keyed by table name, but instead of Column name, it contains Column data
  '-- type indicator. 0-integer. 1-string(255). indexes in array correspond to Column number of columns.
  
  Set DicTables = CreateObject("Scripting.Dictionary")
  Set DicTables2 = CreateObject("Scripting.Dictionary")
  Set DicTableCType = CreateObject("Scripting.Dictionary")

     Set View = DB.OpenView("SELECT `Name` FROM _Tables")
    View.Execute 
      Do
        Set Rec = View.Fetch
                If Rec Is Nothing Then Exit Do
          sTable = Rec.StringData(1) 
           If (Left(sTable, 1) <> "_") Then   '-- don't load _validation table. it's just a copy of all tables combined.
              DicTables.Add sTable, ATab   '-- dic holds array of table Column names.
              DicTables2.Add sTable, ATab  '-- holds array of table Column types.
              DicTableCType.Add sTable, ATab  '-- holds array of table Column types as numeric value.
           End If  
      Loop
      Set Rec = Nothing
   View.Close
  Set View = Nothing
   '---------------------------- sort out tables and columns. ----------------------------------  
   '-- Having put table names into a Dic, this loops through Dic key names,
   '-- getting the columns for that table from the _Columns table. An array
   '-- is created with each array position holding the respective Column name.
   '-- Once this is done it only needs data types to finish.
   ATables = DicTables.Keys
     For i3 = 0 to ubound(ATables)
        sTable = ATables(i3)
        ReDim ATab(30)  '-- reinitialize ATab array. up to 30 columns.
        ReDim ATab2(30) '-- type array. (string or numeric)
        ReDim ATabT(30) '-- specific type array.

            '-- for each table name get all of its columns from the Column table.
            '-- getting Type is a doozy. Microsoft has deliberately hidden the Type
            '-- field of the _columns table. Iteration of this table is the
            '-- best that one can do for discovering table columns, as clunky as it is.
            '--   But it's worse than that. Discovering the data type of the Column
            '-- is a secret, apparently available only to Microsoft buddies. The type
            '-- Column is a string that represents a number(!), which is a bit mask.
            '-- HD00 (3328) seems to indicate a string, while H200 indicates the string
            '-- is a property or constant to be translated. This script is not concerned with
            '-- working out all properties of each Column. It just looks for strings and calls
            '-- the rest integers, because it only uses 2 type: CHAR(255) and INT.
            '-- If you want to work out all MSI types, just rewrite this to add Type directly
            '--to ATab2 array and then print out TableChart. Then drop a normal MSI
            '-- file onto it. That will give you a text file of all bitmask values for all the
            '-- columns in all the 80-odd MSI tables.
           Set View = DB.OpenView("SELECT `Name`, `Type`, `Number` FROM `_Columns` WHERE `Table` ='" & sTable & "'")
            View.Execute
              Do
                Set Rec = View.Fetch
                     If Rec Is Nothing Then Exit Do
                iTyp = CLng(Rec.StringData(2))  '-- secret Type field.
                iCnt = Rec.IntegerData(3)  '-- Column position.
                ATab(iCnt) = Rec.StringData(1) '--Column name.
                  If (iTyp And 3328) = 3328 Then  '-- string
                     ATab2(iCnt) = 1  '-- string.
                   Else
                     ATab2(iCnt) = 0 '-- integer.
                  End If    
                    ATabT(iCnt) = iTyp '-- add specific type.
              Loop
           Set Rec = Nothing
           View.Close
           Set View = Nothing
               '-- snip off unused array elements:
             For iCnt = 1 to 30
                If Len(ATab(iCnt)) = 0 Then
                    ReDim preserve ATab(iCnt - 1)
                    ReDim preserve ATab2(iCnt - 1)
                    ReDim preserve ATabT(iCnt - 1)
                    Exit For
                End If
             Next
           '-- before actually setting up this table as a Dic item, make sure it's not empty.
           '-- (this test had to wait until there was a Column name available to test.)  
           BooValid = ValidTable(sTable, ATab(1))
             If (BooValid = False) Then  '-- invalid table or no content. remove it.
                DicTables.Remove sTable
                DicTables2.Remove sTable
                DicTableCType.Remove sTable

             Else '-- table has at least one row. finish loading Column names and types into table dict.
               '-- assign to dictionaries:      
                DicTables.Item(sTable) = ATab
                DicTables2.Item(sTable) = ATab2
                DicTableCType.Item(sTable) = ATabT

             End If   
      Next   
  MSIFile = sPath 
  Pt1 = InStrRev(MSIFile, "\")
  MSIFolder = Left(MSIFile, Pt1)   '--   c:\folder1\   - includes backslash.
  BooLoaded = True
    TempFolOps 2  '-- make a fresh temp tables folder.
    i4 = ExportAllValidTables(TabFol)  '-- export any tables with content for Find function searches.
  LoadDB = 0     
       '-- Now the database is loaded with MSIEXEC and the 2 dictionaries 
       '-- hold the table organization. DicTables holds an array of Column names
       '-- for each table and DicTables2 holds an array of Column values for each table.
End Function
'------------------------------------- unload database ------------------------------------
Public Sub UnLoadDB()
    On Error Resume Next
  MSIFile = ""
  MSIFolder = ""
  Set Rec = Nothing
  Set View = Nothing
  Set DB = Nothing
  Set WI = Nothing
  Set DicTables = Nothing
  Set DicTables2 = Nothing
  Set DicTableCType = Nothing
  Set DicFind = Nothing
  FindReady = False
  BooLoaded = False
End Sub

Public Function ReLoadDB()
  Dim Ret1
   UnLoadDB
   Ret1 = LoadDB(MSIFile)
   ReLoadDB = Ret1
End Function

Public Function MakeNewDB(sPath)
   Dim DB2
       sErr = ""  '-- clear ErrorString.
      Err.Clear
      On Error Resume Next
    Set WI = CreateObject("WindowsInstaller.Installer")
    Set DB2 = WI.OpenDatabase(sPath, 3)
    DB2.Commit
     If (Err.Number <> 0) Then
        sErr = "MakeNewDB: " & Err.Description
        MakeNewDB = Err.Number
     Else
        MakeNewDB = 0   
     End If
   Set DB2 = Nothing  
   Set WI = Nothing
End Function

  '-- load tables in temp folder into dicfind for searching. returns false if folder doesn't exist or if no tables.
  '-- on success, dicfind is loaded with: key=table name. item = array of rows.
  
Public Function LoadFindData()
  Dim FSO6, TS, s6, oFils, oFol, oFil, sName, Pt1, Pt2, i6, AFil
   LoadFindData = False
   sErr = ""  '-- clear ErrorString.
     On Error Resume Next
      Set FSO6 = CreateObject("Scripting.FileSystemObject")
       If (FSO6.FolderExists(TabFol) = False) Then 
           sErr = "LoadFindData: TEMP subfolder does not exist."
           Set FSO6 = Nothing
           Exit Function
       End If
    Set DicFind = Nothing   
    Set oFol = FSO6.GetFolder(TabFol)  
    Set oFils = oFol.Files
      If (oFils.count > 0) Then
            Set DicFind = CreateObject("Scripting.Dictionary")
          For Each oFil in oFils
             Set TS = oFil.OpenAsTextStream(1)
                 s6 = TS.ReadAll
                 TS.Close
             Set TS = Nothing
               sName = oFil.Name
                Pt1 = InStrRev(sName, ".")
                   If (Pt1 > 0) Then sName = Left(sName, (Pt1 - 1))
              If (sName <> "Summary Info") Then
                     Pt2 = 1  '-- snip top 3 lines from table txt file and save.
                     i6 = 0
                 Do
                     Pt1 = InStr(Pt2, s6, vbCrLf)                     
                       If (Pt1 > 0) Then Pt2 = Pt1 + 1
                    i6 = i6 + 1
                    If i6 = 3 Then Exit Do 
                 Loop  
                   s6 = Right(s6, (len(s6) - Pt2)) 
                   AFil = Split(s6, vbCrLf)
                  DicFind.Add sName, AFil
                  LoadFindData = True 
              End If
          Next    
      Else
         sErr = "LoadFindData: No table files in TEMP subfolder."        
      End If  
             
     Set oFil = Nothing
     Set oFils = Nothing
     Set oFol = Nothing
     Set FSO6 = Nothing 
End Function

'-- find function. searches tables for a given string. return "" if not found. otherwise return: tablename col# vbcrlf tablename col# vbcrlf etc.
Public Function Find1(sToFind)
  Dim s3, sTab, i7, AKeys, UBKeys, AFil, i8, Pt1, s4
   Find1 = ""
     If DicFind is Nothing Then Exit Function
     If DicFind.count = 0 Then Exit Function
     
   s3 = ""
   AKeys = DicFind.Keys
   UBKeys = UBound(AKeys)
      For i7 = 0 to UBKeys
         sTab = AKeys(i7)
         AFil = DicFind.Item(sTab)  '-- get table array.
         For i8 = 0 to UBound(AFil)
            If InStr(1, AFil(i8), sToFind, 1) > 0 Then
               s4 = AFil(i8)
               Pt1 = InStr(1, s4, chr(9))
               If (Pt1 > 0) Then 
                   s4 = Left(s4, (Pt1 - 1))
               Else
                  s4 = ""
               End If
               s3 = s3 & sTab & " - " & s4 & " |" & vbCrLf    '-- like:  Registry 20 vbcrlf
            End If
         Next
      Next
        Find1 = s3      
End Function

'--------------------------- clear the database. WARNING!!! All data will be lost in this MSI file. ------------
Public Sub ClearAllTables()
  Dim sTName, A1(), i4, SIO, ViewC, RecC, RetC
     On Error Resume Next
       RetC = MsgBox("CAUTION: This will remove all data from the file - " & vbCrLf & MSIFile & vbCrLf & "Proceed?", 65, "MSI-Base")
       If RetC = 2 Then Exit Sub
     
 Set ViewC = DB.OpenView("DELETE FROM `_Streams`")
     ViewC.Execute
     DB.Commit
    ViewC.Close
 Set ViewC = Nothing
 Set ViewC = DB.OpenView("DELETE FROM `_Storages`")
     ViewC.Execute
     DB.Commit
    ViewC.Close
 Set ViewC = Nothing
     i4 = -1
  Set ViewC = DB.OpenView("SELECT `Name` FROM _Tables")
    ViewC.Execute
      Do  
         Set RecC = ViewC.Fetch
                  If RecC is Nothing Then Exit Do
             sTName = RecC.StringData(1)
             i4 = i4 + 1
             ReDim preserve A1(i4)
             A1(i4) = sTName
       Loop       
     Set RecC = Nothing
     ViewC.Close
   Set ViewC = Nothing
   
  If i4 > -1 Then
     For i4 = 0 to ubound(A1)
       sTName = A1(i4)
       RemoveTable sTName
     Next
  End If  
    DB.Commit
     DicTables.RemoveAll
     DicTables2.RemoveAll
     DicTableCType.RemoveAll

End Sub

'----------------------------- Remove a database table. ------------------------------------
Public Sub RemoveTable(sTab)
  Dim ViewR
     On Error Resume Next
  Set ViewR = DB.OpenView("DROP TABLE `" & sTab & "`")
    ViewR.Execute
    DB.Commit
    ViewR.Close
  Set ViewR = Nothing  
      DicTables.Remove sTab
      DicTables2.Remove sTab
      DicTableCType.Remove sTab

End Sub

'--------------------------- Add a table. ---------------------------------------------
'-- "col1 i, col2 s, col3 i"  Column name followed by data type indicator. must be i or s. must be followed by comma.  
Public Function AddTable(sTabName, ColumnString)
  Dim ViewT, ATab(), ATab2(), s2, ACols, sCol1, K1, i4, UB1, sType
    Err.Clear
  On Error Resume Next
     sErr = ""  '-- clear ErrorString.
    ACols = Split(ColumnString, ",")
      UB1 = UBound(ACols)
      ReDim ATab(UB1 + 1)
      ReDim ATab2(UB1 + 1)
        For i4 = 0 to UB1
          sCol1 = Trim(ACols(i4))
          sType = UCase(Right(sCol1, 1))  
          sCol1 = Left(sCol1, (len(sCol1) - 2))
                If i4 = 0 Then K1 = sCol1  '-- get primary key.
          s2 = s2 & cCar &  sCol1
          ATab(i4 + 1) = sCol1  '-- build dictionary array while building sql string.
           If sType = "I" Then
              If (i4 = 0) Then
                 s2 = s2 & "` LONG NOT NULL, "
              Else
                 s2 = s2 & "` LONG, "
              End If    
              ATab2(i4 + 1) = 0        
           Else
              If (i4 = 0) Then
                  s2 = s2 & "` CHAR(255) NOT NULL, "  
              Else
                  s2 = s2 & "` CHAR(255), "   
              End If    
              ATab2(i4 + 1) = 1
           End If 
       Next       
        '-- string is now:  `name` LONG, `name2` LONG,       etc.
        s2 = Left(s2, (len(s2) - 2))
        s2 = "CREATE TABLE `" & sTabName & "` (" & s2
        s2 = s2 &  " PRIMARY KEY `" & K1 & "` )"
    Set ViewT = DB.OpenView(s2)
         ViewT.Execute
       DB.Commit
      ViewT.Close
  Set ViewT = Nothing  
      '-- update dictionaries.
   
  If (Err.Number <> 0) Then
        sErr = "AddTable: " & Err.Description
        AddTable = Err.Number
  Else
        DicTables.Add sTabName, ATab
        DicTables2.Add sTabName, ATab2
        AddTable = 0
  End If       
End Function

  '-------------------------------------------------------- AddTableCustom - adds new table with data type options. ------------
Public Function AddTableCustom(sTabName, ColumnString)
    Dim ViewT, ATab(), ATab2(), s2, ACols, sCol1, K1, i4, UB1, sType, Pt1
    Err.Clear
  On Error Resume Next
     sErr = ""  '-- clear ErrorString.
    ACols = Split(ColumnString, ",")
      UB1 = UBound(ACols)
      ReDim ATab(UB1 + 1)  ' -- store Column names: "Col1"
      ReDim ATab2(UB1 + 1)  '-- sotre Column data type and proprties strings: "CHAR(70) NOT NULL"
        For i4 = 0 to UB1
          sCol1 = Trim(ACols(i4))
          Pt1 = InStr(1, sCol1, " ")
            If (Pt1 > 1) Then
                  sType = Right(sCol1, (len(sCol1) - (Pt1 - 1)))   '-- get right side of string from first " " onward.
                  sCol1 = Left(sCol1, (Pt1 - 1))
                        If i4 = 0 Then 
                              K1 = sCol1  '-- get primary key.
                              If InStr(1, sType, "NOT NULL") = 0 Then sType = sType & " NOT NULL"
                        End If      
                  s2 = s2 & cCar &  sCol1 & cCar & sType & ", "   '--:  s2 & `colname` CHAR(70) NOT NULL,  
                  ATab(i4 + 1) = sCol1  '-- build dictionary array while building sql string.
                   If InStr(1, sType, "CHAR") > 0 Then  '-- save as string or int.
                       ATab2(i4 + 1) = 1       
                   Else
                       ATab2(i4 + 1) = 0
                   End If 
            End If        
        Next       
        '-- string is now:  `name` LONG, `name2` LONG,       etc.
        s2 = Left(s2, (len(s2) - 2))
        s2 = "CREATE TABLE `" & sTabName & "` (" & s2 &  " PRIMARY KEY `" & K1 & "` )"
     Set ViewT = DB.OpenView(s2)
         ViewT.Execute
       DB.Commit
     ViewT.Close
  Set ViewT = Nothing  
      '-- update dictionaries.
   
  If (Err.Number <> 0) Then
        sErr = "AddTableCustom: " & Err.Description
        AddTable = Err.Number
  Else
        DicTables.Add sTabName, ATab
        DicTables2.Add sTabName, ATab2
        AddTableCustom = 0
  End If       

End Function
  '------------------------------------- import table --------------------------
Public Sub ImportTable(sPath)
   Dim sFol, sFil, Pt1
     On Error Resume Next
       Pt1 = InStrRev(sPath, "\")
       sFol = Left(sPath, (Pt1 - 1))
       sFil = Right(sPath, (len(sPath) - Pt1))
         DB.Import sFol, sFil
       DB.Commit
End Sub

  '----------------------------------------Add  a Column ----------------------------------------------
Public Function AddColumn(sTabName, sColName, BooNumeric)
  Dim ViewT, ATab(), ATabD, i5, i6
             Err.Clear
            On Error Resume Next
          sErr = ""  '-- clear ErrorString.
                If (DicTables.Exists(sTabName) = False) Then
                    sErr = "AddColumn: Invalid table name."
                    AddColumn = 1
                    Exit Function
                End If
                    
     If (BooNumeric = False) Then
        Set ViewT = DB.OpenView("ALTER TABLE `" & sTabName & "` ADD `" & sColName & "` CHAR(255)")
     Else
        Set ViewT = DB.OpenView("ALTER TABLE `" & sTabName & "` ADD `" & sColName & "` LONG")
     End If    
    ViewT.Execute
    DB.Commit
    ViewT.Close
  Set ViewT = Nothing  
   
  If (Err.Number <> 0) Then
        sErr = "AddColumn: " & Err.Description
        AddColumn = Err.Number
  Else
         '-- update table dics.
         ATabD = DicTables.item(sTabName)
         i5 = UBound(ATabD)
         ReDim ATab(i5 + 1)
           For i6 = 1 to i5
              ATab(i6) = ATabD(i6)
           Next
             ATab(i5 + 1) = sColName
             DicTables.item(sTabName) = ATab
          
          ATabD = DicTables2.item(sTabName)
          ReDim ATab(i5 + 1)
              For i6 = 1 to i5
                  ATab(i6) = ATabD(i6)
              Next
           If (BooNumeric = True) Then
               ATab(i5 + 1) = 0
           Else
              ATab(i5 + 1) = 1
           End If      
             DicTables2.item(sTabName) = ATab
    
           AddColumn = 0
  End If

End Function

Public Sub WriteToDisk()
  DB.Commit
End Sub

 '------------------------------------------ Add a record -------------------------------------
'-- AddRecord("Table1", "rose, red, 24, -, no")  
'-- AddRecord will check for proper string.

Public Function AddRecord(sTabName, sVals)
  Dim ViewAR, AFld, AType, AVals, iUB, i3, sf, sv
     On Error Resume Next     
       sErr = ""  '-- clear ErrorString.   
          If (DicTables.Exists(sTabName) = False) Then 
             sErr = "AddRecord: Invalid table name."
             AddRecord = 1 '-- no such table.
             Exit Function
          End If      
    AFld = DicTables.Item(sTabName)
    AType = DicTables2.Item(sTabName)
     sVals = Replace(sVals, ", ", ",")  '-- remove spaces between items.
     AVals = Split(sVals, ",")
     iUB = UBound(AFld)
          If UBound(AVals) <> (iUB - 1) Then 
             sErr = "AddRecord: Invalid number of values."
             AddRecord = 2  '-- wrong number of data values.
             Exit Function  '-- wrong number of items.
          End If   
          
    For i3 = 1 to iUB
      If (AVals(i3 - 1) <> "-") Then   '-- if no entry for this field it's "-", so Skip it.
         sf = sf & cCar & AFld(i3) & cCar & ", "
           If AType(i3) = 0 Then
                If (IsNumeric(AVals(i3 - 1)) = False) Then
                   sErr = "AddRecord: " & AFld(i3) & " value is numeric."
                   AddRecord = 3  '-- sent string for integer value.
                   Exit Function
                Else
                   sv = sv & AVals(i3 - 1) & ", " 
                End If
           Else   '-- string value.
               sv = sv & cApos & AVals(i3 - 1) & cApos & ", " 
           End If    
      End If
   Next
      sf = Left(sf, (len(sf) - 2))   '--   `flower`, `color`, `height`, `evergreen`   (note that Column 4 is not here because it was "-". evergreen is Column 5.)
      sv = Left(sv, (len(sv) - 2))  '--   'rose', 'red', 24, 'no'      
      
       Err.Clear
       On Error Resume Next
    Set ViewAR = DB.OpenView("INSERT INTO `" & sTabName & "` (" & sf & ") VALUES (" & sv & ")")
          ViewAR.Execute 
          DB.Commit
        ViewAR.Close
    Set ViewAR = Nothing

  If (Err.Number <> 0) Then
     sErr = "AddRecord: " & Err.Description
     AddRecord = Err.Number
  Else
     AddRecord = 0
  End If
End Function

  '----------------------------------------------------------- Remove record ----------------------
Public Function RemoveRecord(sTabName, sColumn, CValue)
 Dim ViewRR, ColType
    Err.Clear
  On Error Resume Next
   sErr = ""  '-- clear ErrorString.
    ColType = ValType(sTabName, sColumn, "RemoveRecord")
       If ColType = -1 Then   '-- table or Column not valid.
          sErr = "RemoveRecord: Table or column name invalid."
           RemoveRecord = 1 
           Exit Function
        End If   
    If (ColType = 0) Then       
        Set ViewRR = DB.OpenView("DELETE FROM `" & sTabName & "` WHERE `" & sColumn & "` =" & CValue)
    Else
        Set ViewRR = DB.OpenView("DELETE FROM `" & sTabName & "` WHERE `" & sColumn & "` =" & cApos & CValue & cApos)
    End If
          ViewRR.Execute 
          DB.Commit
        ViewRR.Close
    Set ViewRR = Nothing
  
   If (Err.Number <> 0) Then
     sErr = "RemoveRecord: " & Err.Description
     RemoveRecord = Err.Number
  Else
     RemoveRecord = 0
  End If

End Function

  '-- given a table name and first Column name, test for a value to see if table exists and has content.
  '--  ValidTable returns true only if table exists AND has at leat one row of data.
  
 Public Function ValidTable(sTabName, sColumnFrom)    '-- boolean.
   Dim ViewV, RecV
    ValidTable = False
      Err.Clear
    On Error Resume Next
                   '-- get Column data types and make sure that table, Column names are valid.   
   '  ColType = ValType(sTabName, sColumnFrom, "")
      '    If (ColType = -1) Then  Exit Function
      '-- try to retrieve one row:
    Set ViewV = DB.OpenView("SELECT `" & sColumnFrom & "` FROM `" & sTabName & "`")
      ViewV.Execute
         If Err.Number <> 0 Then  '-- column or table name not valid.
             Set ViewV = Nothing
             Exit Function
         End If
          '-- edited 4-27-07      
      If LoadEmptyTables = False Then
          Set RecV = ViewV.Fetch
               ValidTable = Not (RecV is Nothing)  '-- if RecV not nothing then at least one row exists.
          Set RecV = Nothing
      Else
          ValidTable = True  
      End If    
    Set ViewV = Nothing   
End Function

  '--------------------------------------------- Get all values from a Column. ----------------------------
  '-- (Tablename, ColumnName)  returns array of all values in given Column. 
  '-- On success Array(0) is number of items returned. If failed Array(0) = -1
   
Public Function GetAllValues(sTabName, sColumnFrom)
  Dim ACols(), ViewV, RecV, i3, iCols, ColType, sCol1
      Err.Clear
    On Error Resume Next
     sErr = ""  '-- clear ErrorString.
                   '-- get Column data types and make sure that table, Column names are valid.   
     ColType = ValType(sTabName, sColumnFrom, "GetAllValues")
        If (ColType = -1) Then
            ReDim ACols(0)
            ACols(0) = -1
            Exit Function  '-- ValType function records error.
        End If

      sCol1 = GetColumn1(sTabName)  '-- get name of first column for sorting.
      ReDim ACols(50)
      i3 = 50
      iCols = 0
    Set ViewV = DB.OpenView("SELECT `" & sColumnFrom & "` FROM `" & sTabName & "`")  ' ORDER BY `" & sCol1 & "`")
      ViewV.Execute
       Do  
          Set RecV = ViewV.Fetch
                 If RecV is Nothing Then Exit Do
             iCols = iCols + 1
                    If iCols > i3 Then
                        i3 = i3 + 50
                        ReDim preserve ACols(i3)
                    End If 
              If (ColType = 0) Then        
                 ACols(iCols) = RecV.IntegerData(1)
              Else
                 ACols(iCols) = RecV.StringData(1)
              End If          
        Loop       
     Set RecV = Nothing
     ViewV.Close
     Set ViewV = Nothing
     
   ReDim preserve ACols(iCols)
   
    If Err.Number <> 0 Then
       sErr = "GetAllValues: " & Err.Description
       iCols = -1
    End If     
          '-- send back array of values with number of values returned in Array(0).
       If (ColType = 0) Then  
          ACols(0) = iCols
       Else
          ACols(0) = CStr(iCols)
       End If 
   GetAllValues = ACols    
End Function

   '-- get all values for columnFrom where ColumnCompare value = ValueCompare.
   '-- (tablename, columnName, ColumnCompare, ValueCompare) returns all Column values in
   '-- ColumnName Column where ColumnCompare value matches ValueCompare.
   '-- On success Array(0) is number of items returned. If failed Array(0) = -1
Public Function GetSelectValues(sTabName, sColumnFrom, sColumnCompare, ValueCompare)
     Dim ACols(), ViewV, RecV, i3, iCols, FromType, CompareType
      Err.Clear
      On Error Resume Next
        sErr = ""  '-- clear ErrorString.
        '-- get Column data types and make sure that table, Column names are valid.   
          FromType = ValType(sTabName, sColumnFrom, "GetSelectValues")
          CompareType = ValType(sTabName, sColumnCompare, "GetSelectValues")
        If (FromType = -1) Or (CompareType = -1) Then
            ReDim ACols(0)
            ACols(0) = -1
            Exit Function  '-- ValType function records error.
        End If
        If (CompareType = 0) And (IsNumeric(CompareType) = False) Then
           sErr = "GetSelectValues: Data type of ValueCompare parameter should be numeric."
           ReDim ACols(0)
            ACols(0) = -1
            Exit Function
        End If
        
      ReDim ACols(50)
      i3 = 50
      iCols = 0
         If (CompareType = 1) Then ValueCompare = cApos & ValueCompare & cApos
    Set ViewV = DB.OpenView("SELECT `" & sColumnFrom & "` FROM `" & sTabName & "` WHERE `" & sColumnCompare & "` =" & ValueCompare)
     
 ViewV.Execute
       Do  
          Set RecV = ViewV.Fetch
                 If RecV is Nothing Then Exit Do
             iCols = iCols + 1
                    If iCols > i3 Then
                        i3 = i3 + 50
                        ReDim preserve ACols(i3)
                    End If 
              If (FromType = 1) Then        
                 ACols(iCols) = RecV.StringData(1)
              Else
                 ACols(iCols) = RecV.IntegerData(1)
              End If          
        Loop       
     Set RecV = Nothing
     ViewV.Close
     Set ViewV = Nothing
     
   ReDim preserve ACols(iCols)
   
    If Err.Number <> 0 Then
       sErr = "GetSelectValues: " & Err.Description
       iCols = -1
    End If     
       If (FromType = 0) Then  
          ACols(0) = iCols
       Else
          ACols(0) = CStr(iCols)
       End If 
   GetSelectValues = ACols    

End Function

     '-- set all values in sColumnSet Column to ValSet, where value in sColCompare Column is ValCompare. -----------------
Public Function SetSelectValues(sTabName, sColumnSet, ValSet, sColCompare, ValCompare)
  Dim ViewS, SetType, CompType, sCom
     Err.Clear
    On Error Resume Next
       sErr = ""  '-- clear ErrorString.
         SetType = ValType(sTabName, sColumnSet, "SetSelectValues")
         CompType = ValType(sTabName, sColCompare, "SetSelectValues")
           If (SetType = -1) Or (CompType = -1) Then
               SetSelectValues = 1  '-- invalid Column or table name.
               Exit Function   '-- ValType will have recorded sErr string.
           End If
            If (SetType = 0 And IsNumeric(ValSet) = False) Or (CompType = 0 And IsNumeric(ValCompare) = False) Then
                sErr = "SetSelectValues: Wrong data type. Value must be numeric."
                SetSelectValues = 2  '-- one of the values is supposed to be numeric but isn't.
                Exit Function
             End If
           If (sColumnSet = sColCompare) Then
              sErr = "SetSelectValues: Cannot set key value. Record must be deleted and a new record created."
              SetSelectValues = 3  '-- trying to change key value.
              Exit Function
           End If

        If SetType = 1 Then ValSet = cApos & ValSet & cApos
        If CompType = 1 Then ValCompare = cApos & ValCompare & cApos
  
       Set ViewS = DB.OpenView("UPDATE `" & sTabName & "` SET `" & sTabName & "`.`" & sColumnSet & "` =" & ValSet & " WHERE `" & sTabName & "`.`" & sColCompare & "` =" & ValCompare)
          ViewS.Execute
          DB.Commit
       Set ViewS = Nothing
     
       If Err.Number <> 0 Then
          sErr = "SetSelectValues: " & Err.Description
          SetSelectValues = Err.Number
       Else
          SetSelectValues = 0
       End If     

End Function
 '----------------------------------------------------------- Set all values ------------------------
Public Function SetAllValues(sTabName, sColName, ValSet)
  Dim ViewS, SetType
     Err.Clear
     On Error Resume Next
       sErr = ""  '-- clear ErrorString.
         SetType = ValType(sTabName, sColumnSet, "SetAllValues")
         If (SetType = -1) Then
               SetAllValues = 1  '-- invalid Column or table name.
               Exit Function  '-- ValType will have recorded error string.
         End If
         If (SetType = 0 And IsNumeric(ValSet) = False) Then
               sErr = "SetAllValues: Value to set is numeric but the 3rd parameter sent to the function was not numeric."
               SetAllValues = 2  '-- value is supposed to be numeric but isn't.
               Exit Function
          End If
        
        If SetType = 1 Then ValSet = cApos & ValSet & cApos
    Set ViewS = DB.OpenView("UPDATE `" & sTabName & "` SET `" & sColumn & "` =" & ValSet)
        ViewS.Execute
         DB.Commit
         ViewS.Close
     Set ViewS = Nothing
     
       If Err.Number <> 0 Then
          sErr = "SetAllValues: " & Err.Description
          SetAllValues = Err.Number
       Else
          SetAllValues = 0
       End If     


End Function

'------------------------------------- Table Chart -------------------------------
Public Function GetTableChart()
  Dim s2, A2, A3, A4, A5, i5, i6, sTabName
     A2 = DicTables.Keys
     A3 = DicTables2.Keys
     
   If (DicTables.Count > 0) Then
       For i5 = 0 to UBound(A2)
          sTabName = A2(i5) '-- table name.
          s2 = s2 & sTabName & vbCrLf  '--table name.
          A4 = DicTables.Item(sTabName)   '-- names array.
          A5 = DicTables2.Item(sTabName)  '-- types array.
            For i6 = 1 to UBound(A4)
                s2 = s2 & chr(9) & A4(i6) & " " & CStr(A5(i6)) & vbCrLf  '-- returns each Column name with type: name 0 is int. name 1 is string.
            Next         
       Next
  Else
       s2 = ""     
  End If     
    
  GetTableChart = s2
End Function

  '------------------ return entire table contents as string.
Public Function GetTableContent(sTabName)
   Dim ViewT, cTab, RecT, sRet, ACols, ATypes, i3, iUB, AList(), ALine(), sCol, iCnt, iCols, ColType, i4
        sErr = ""  '-- clear ErrorString.
   If (DicTables.Exists(sTabName) = False) Then
      sErr = "GetTableContent: Invalid table name."
      GetTableContent = ""
      Exit Function
   End If
      On Error Resume Next
  cTab = chr(9)
  ACols = DicTables.Item(sTabName)
  ATypes = DicTables2.Item(sTabName)
   sRet = Join(ACols, cTab)  & vbCrLf & vbCrLf
   iUB = UBound(ACols)
   ReDim ALine(iUB - 1)
      For i3 = 1 to iUB
          sCol = sCol & " " & cCar & ACols(i3) & cCar & ","  '-- build query string.
      Next
        sCol = Left(sCol, (len(sCol) - 1))       
 
             iCols = 0
             iCnt = 50
             ReDim AList(iCnt)
         Set ViewT = DB.OpenView("SELECT" & sCol & " FROM `" & sTabName & "`")
           ViewT.Execute
              Do  
                 Set RecT = ViewT.Fetch
                         If RecT is Nothing Then Exit Do
                  iCols = iCols + 1
                    If iCols > iCnt Then
                        iCnt = iCnt + 50
                        ReDim preserve AList(iCnt)
                    End If 
                  For i4 = 1 to iUB   '-- collect values in an array sized to number of columns.
                      ColType = ATypes(i4)
                        If (ColType = 0) Then        
                             ALine(i4 - 1) = CStr(RecT.IntegerData(i4)) 
                         Else
                             ALine(i4 - 1) = Trim(RecT.StringData(i4))
                         End If     
                  Next
                     AList(iCols) = Join(ALine, cTab)  '-- join the array with yabs and add string to AList.
                                              '-- This method is in hopes ofavoiding bogdown due to extreme
                                              '-- numbers of small string concatenations.    
             Loop       
         Set RecT = Nothing
         ViewT.Close
         Set ViewT = Nothing
           GetTableContent = sRet & Join(AList, vbCrLf)
End Function

  '----------------------------- export tables ----------------------------
Public Function ExportAllTables(FolderPath)
 Dim FSO1, Boo1, ViewE, RecE, sTable, oFol
   sErr = ""  '-- clear ErrorString.
 
    Set FSO1 = CreateObject("Scripting.FileSystemObject")
       Boo1 = FSO1.FolderExists(FolderPath) 
       If (Boo1 = False) Then
           ExportAllTables = 1  '-- invalid path.
           sErr = "ExportAllTables: Invalid folder path."
           Set FSO1 = Nothing
           Exit Function
       Else    
          Set oFol = FSO1.CreateFolder(FolderPath & "\export")
          FolderPath = FolderPath & "\export"
          Set oFol = Nothing
       End If
    Set FSO1 = Nothing
       Err.Clear
       On Error Resume Next
   
    Set ViewE = DB.OpenView("SELECT `Name` FROM _Tables")
    ViewE.Execute 
     Do
       Set RecE = ViewE.Fetch
             If RecE Is Nothing Then Exit Do
          sTable = RecE.StringData(1) 
          DB.Export sTable, FolderPath, sTable & ".txt"
     Loop
        Set RecE = Nothing
        ViewE.Close
      Set ViewE = Nothing
    
     If Err.Number <> 0 Then
          sErr = "ExportAllTables: " & Err.Description
          ExportAllTables = Err.Number
     Else
          ExportAllTables = 0
     End If         
End Function

'----------------------------- export only tables With content. ----------------------------
Public Function ExportAllValidTables(FolderPath)
 Dim FSO1, Boo1, ViewE, RecE, sTable, oFol, ATablesEx, i4, UBTab
        sErr = ""  '-- clear ErrorString.
    Set FSO1 = CreateObject("Scripting.FileSystemObject")
       If (FSO1.FolderExists(FolderPath) = False) Then
           ExportAllValidTables = 1  '-- invalid path.
           sErr = "ExportAllValidTables: Invalid folder path."
           Set FSO1 = Nothing
           Exit Function
       End If
           Set FSO1 = Nothing
       Err.Clear
       On Error Resume Next
    ATablesEx = DicTables.Keys  '-- array of valid key names. As long as this is called after file is loaded, array should be valid.
    UBTab = UBound(ATablesEx)
    Set ViewE = DB.OpenView("SELECT `Name` FROM _Tables")
    ViewE.Execute 
     Do
       Set RecE = ViewE.Fetch
             If RecE Is Nothing Then Exit Do
          sTable = RecE.StringData(1) 
          For i4 = 0 to UBTab
             If (sTable = ATablesEx(i4)) Then '-- table name matches an actual used table, so export it.
                  DB.Export sTable, FolderPath, sTable & ".txt"
                  Exit For
             End If  
          Next       
     Loop
        Set RecE = Nothing
        ViewE.Close
      Set ViewE = Nothing
    
     If Err.Number <> 0 Then
          sErr = "ExportAllValidTables: " & Err.Description
          ExportAllValidTables = Err.Number
     Else
          ExportAllValidTables = 0
     End If         
End Function

'--------------------------------- list embedded CABs --------------------------------
    '-- return list of embedded CABs, MergeModule.CABinet if it's an MSM.
    '--  >> Updated October '06. The original version of this function checked
    '-- the _Streams table for entries that included ".cab", but entries
    '-- in _Streams are sometimes labeled with a GUID instead of a filename.
    '-  So this version checks the Media table, Cabinet column for anything
    '-- starting with "#". That will be an embedded CAB.
      '-- NOTE that this function does not list all CABs but only embedded CABs.
      '-- This function is generally not of much use. It's included here only
      '-- for people who may be using this class with an MSI used for
      '-- software installation.
   
Public Function GetCABList()
 Dim ViewC, RecC, sCabList3, s2, ARet(), A2, UBa, i7
   On Error Resume Next
           ReDim ARet(0)
           ARet(0) = "0"   
      Set ViewC = DB.OpenView("SELECT `Cabinet` FROM Media")
      ViewC.execute
       Do
        Set RecC = ViewC.Fetch
              If RecC is Nothing Then Exit Do
           s2 = RecC.stringdata(1)          
             If InStr(1, s2, "MergeModule.", 1) > 0 Then
                 ReDim ARet(1)
                 ARet(0) = "1"
                 ARet(1) = s2
                 GetCABList = ARet
                 Exit Function
             Else
                 If (Left(s2, 1) = "#") Then
                     sCabList3 = sCabList3 & Chr(1) & (Right(s2, (len(s2) - 1)))
                 End If
             End If   
      Loop
         If (Len(sCabList3) > 0) Then
            A2 = Split(sCabList3, Chr(1))
            UBa = UBound(A2)
            ReDim ARet(UBa + 1)
            ARet(0) = CStr(UBa + 1)
              For i7 = 0 to UBa
                  ARet(i7 + 1) = A2(i7)
              Next  
         End If  
            GetCABList = ARet
      Set RecC = Nothing
      Set ViewC = Nothing
End Function
  '-------------------------------------  Extract CAB file from MSM or MSI  -----------------------
Public Function ExtractCAB(sCabName)
 Dim ViewE, RecE, DLen, FSO2, TS2, s2, sFile2
  On Error Resume Next
           sErr = ""  '-- clear ErrorString.
    Set FSO2 = CreateObject("Scripting.FileSystemObject")
     If (FSO2.FolderExists(MSIFolder) = False) Then
       sErr = "ExtractCAB: No file loaded."
       ExtractCAB = 1
       Set FSO2 = Nothing
       Exit Function
    End If
       Err.Clear
       On Error Resume Next
  Set ViewE = DB.OpenView("SELECT `Name`,`Data` FROM _Streams WHERE `Name`= '" & sCabName & "'")
    ViewE.execute
     Set RecE = ViewE.Fetch
          If RecE is Nothing Then
               Set ViewE = Nothing
               sErr = "ExtractCAB: Failed to extract CAB."
                 If (Err.Number <> 0) Then 
                     sErr = sErr & " " & Err.Description
                     ExtractCAB = Err.Number
                 Else        
                     ExtractCAB = 2  '--just make up a non-zero error number if this is somehow not an MSI error.
                 End If  
                    Set FSO2 = Nothing
                    Exit Function
           End If
         DLen = RecE.datasize(2)
         s2 = RecE.ReadStream(2, DLen, 2)
         sFile2 = sCabName

     If (UCase(Right(sFile2, 4)) <> ".CAB") Then sFile2 = sFile2 & ".cab"
   Set TS2 = FSO2.CreateTextFile(MSIFolder & sFile2, True, False)
       TS2.Write s2
       TS2.Close
   Set TS2 = Nothing
Set FSO2 = Nothing
Set Rec = Nothing
Set View = Nothing
     If (Err.Number <> 0) Then sErr = "ExtractCAB: Error - " & Err.Description
  ExtractCAB = Err.Number
End Function

 '------------------------ insert CAB into MSM or MSI. If CAB of same name already exists it will be replaced.
Public Function InsertCAB(sCabName, sNewPath)
  Dim FSO2, ViewI, RecI, BooFile
     On Error Resume Next
          sErr = ""
      Set FSO2 = CreateObject("Scripting.FileSystemObject")
          BooFile =  (FSO2.FileExists(sNewPath))
      Set FSO2 = Nothing   
    If (BooFile = False) Then    
        sErr = "InsertCAB: New CAB file path invalid."
        InsertCAB = 1
       Exit Function
    End If
    
       Err.Clear
       On Error Resume Next
     Set ViewI = DB.OpenView("SELECT `Name`,`Data` FROM _Streams")
        ViewI.execute
        Set RecI = WI.CreateRecord(2)
	RecI.StringData(1) = sCabName
        	RecI.SetStream 2, sNewPath
        	ViewI.Modify 3, RecI
        DB.Commit
     Set RecI = Nothing
    Set ViewI = Nothing   
       If (Err.Number <> 0) Then sErr = "InsertCAB: Error - " & Err.Description
       InsertCAB = Err.Number
End Function
'---------------------------------------------- array of Column names. ---------------
Public Function GetTableColumns(sTabName, ATabList)
  Dim ACols1, iCols
     sErr = ""  '-- clear ErrorString.

   If DicTables.Exists(sTabName) = False Then
      GetTableColumns = -1
      sErr = "GetTableColumns: Invalid table name."
      Exit Function
   End If
       '-- send back as a 1-based array for uniformity.
   ACols1 = DicTables.Item(sTabName)
   iCols = UBound(ACols1)
   ATabList = ACols1
   GetTableColumns = iCols
End Function

'---------------------------------------------- array of Column types. ---------------
Public Function GetTableColumnTypes(sTabName, ATabList)
  Dim ACols1, iCols
         sErr = ""  '-- clear ErrorString.

    If DicTables2.Exists(sTabName) = False Then
      GetTableColumnTypes = -1
      sErr = "GetTableColumnTypes: Invalid table name."
      Exit Function
   End If
         '-- send back as a 1-based array for uniformity.
   ACols1 = DicTables2.Item(sTabName)
   iCols = UBound(ACols1)
      ATabList = ACols1
   GetTableColumnTypes = iCols
End Function

'---------------------------------------------- array of Column specific types. ---------------
Public Function GetTableColumnSpecificTypes(sTabName, ATabList)
  Dim ACols1, iCols
         sErr = ""  '-- clear ErrorString.

    If DicTableCType.Exists(sTabName) = False Then
      GetTableColumnSpecificTypes = -1
      sErr = "GetTableColumnTypes: Invalid table name."
      Exit Function
   End If
         '-- send back as a 1-based array for uniformity.
   ACols1 = DicTableCType.Item(sTabName)
   iCols = UBound(ACols1)
      ATabList = ACols1
   GetTableColumnSpecificTypes = iCols
End Function

'---------------------------------------------- return data type info. for one column. ---------------
  ' return: 0-ok. -1-invalid table name. -2-no column match.
Function GetColumnDataType(sTableName, sColumnName, AType)
  Dim AColsC, AColsT, iCols, iTyp, ARet(3), i3, HiB
         sErr = ""  '-- clear ErrorString.
         iTyp = 0
          On Error Resume Next 
   If DicTableCType.Exists(sTableName) = False Then
      GetColumnDataType = -1
      sErr = "GetTableColumnTypes: Invalid table name."
      Exit Function
   End If
   
    '-- get array of column names and array of column values.
  AColsT = DicTableCType.Item(sTableName)
  AColsC = DicTables.Item(sTableName)
    For iCols = 1 to UBound(AColsC)
            If AColsC(iCols) = sColumnName Then
            iTyp = AColsT(iCols)
            Exit For
        End If
    Next
    
   If (iTyp < 256) Then  '-- would have to be at least H100 to be valid.
     GetColumnDataType = -2
      sErr = "GetTableColumnTypes: Column name " & sColumnName & " not found for table " & sTableName & "."
      Exit Function
   End If
  
'-- Now starts the process of figuring out the data type. The type is stored as a number.
'-- if translated to hex that number will have 3 or 4 places. The two right places denote
'-- max length. The 3rd place fromn the right denotes data type. The left-side place denotes
'-- nullable and/or key. This value is "secret" in that Microsoft has not documented it!
'-- It's possible that there are other aspects besides those listed here, especially given
'-- the half-baked structure of data types altogether in Windows Installer. But what's here
'-- seems to be what's relevant. There seems to be no place in this hex bit mask value to
'-- store things like pseudo-sub-data-type info.
'-- The bit mask for left-side byte or "high byte": H1000 (4096) - nullable. H2000 (8192) - key.
'-- HD00 (3328) - string. H500 (1280) - short int. H900 (2304) - binary  H100 (256) - long int.  HF00 (3840) - localized[variable]
 
    For i3 = 0 to 3
       ARet(i3) = 0
    Next
      '-- mask high byte to check for proerties:
     HiB = iTyp \ 256
    If (HiB And 16) = 16 Then ARet(0) = 1
    If (HiB And 32) = 32 Then ARet(1) = 1
         '-- mask off upper range of high byte to get data type:
       If HiB > 15 Then HiB = HiB Mod 16
   Select Case HiB
       Case 1  '--long int.
           ARet(2) = 2
       Case 5  '-- short int.
           ARet(2) = 4
       Case 9  '-- binary
           ARet(2) = 1
       Case 13 '-- string
            ARet(2) = 0
            ARet(3) = iTyp And 255  '--max. length.
       Case 15  '-- localized.
            ARet(2) = 3
            ARet(3) = iTyp And 255
    End Select  
       AType = ARet  
       GetColumnDataType = 0
End Function
  
' ------------------------------------------------- Get table names -----------------------
Public Function GetTableNames(ATabList)
  Dim ANames, ATabs(), iTabs, i6
        sErr = ""  '-- clear ErrorString.
     If DicTables.count = 0 Then
        GetTableNames = -1
        sErr = "GetTableNames: No tables loaded."
        Exit Function
     End If
     ANames = DicTables.keys
        '-- this is a bit clunky, but it helps to standardize. Since arrays in the class are being returned as 1-base, this follows suit.
        '-- array(0) is not used.
     iTabs = ubound(ANames)
     ReDim ATabs(iTabs + 1)
       For i6 = 0 to iTabs
          ATabs(i6 + 1) = ANames(i6)
       Next
     ATabList = ATabs  
     GetTableNames = iTabs + 1      
End Function


'-- Summary info. --------------------------------------------------------
Public Sub ClearSummary()
  Dim SIO, iCnt, i4
     On Error Resume Next
Set SIO = DB.SummaryInformation(20)
 iCnt = SIO.PropertyCount
    If (iCnt > 2) Then
        For i4 = 2 to iCnt - 1
          Select Case i4
             Case 2, 3, 4, 5, 6, 7, 8, 9
                 SIO.Property(i4) = ""
             Case 14, 15, 16, 19
                 SIO.Property(i4) = 0
          End Select       
        Next  
    End If     
       SIO.Persist   
Set SIO = Nothing      
End Sub

Public Function GetSummaryValue(sValName)
 Dim i4, SIO
    On Error Resume Next
  Select Case UCase(sValName)
      Case "TITLE"
         i4 = 2
      Case "SUBJECT"
         i4 = 3
      Case "AUTHOR"
         i4 = 4
      Case "NOTES"
         i4 = 6
      Case Else
          GetSummaryValue = ""
          Exit Function
   End Select
     Set SIO = DB.SummaryInformation(0)
       GetSummaryValue = SIO.Property(i4)
     Set SIO = Nothing
End Function

Public Sub SetSummaryValues(Val1, Val2, Val3, Val4)
 Dim SIO, iCnt, i4
     On Error Resume Next
  Set SIO = DB.SummaryInformation(20)
  iCnt = SIO.PropertyCount
    If (iCnt > 2) Then
       For i4 = 2 to iCnt - 1
         Select Case i4
           Case 2
             SIO.Property(i4) = CStr(Val1)
           Case 3
             SIO.Property(i4) = CStr(Val2)
           Case 4
             SIO.Property(i4) = CStr(Val3)
           Case 6
             SIO.Property(i4) = CStr(Val4)
           Case 5, 7, 8, 9
             SIO.Property(i4) = ""
            Case 14, 15, 16, 19
              SIO.Property(i4) = 0
          End Select
        Next   
    End If     
       SIO.Persist
    Set SIO = Nothing   
  End Sub
  
 '--return name of first column in given table.
Public Function GetColumn1(sTab1)
    Dim ACols
       On Error Resume Next
    ACols = DicTables.Item(sTab1)
    GetColumn1 = ACols(1)
End Function

'-- returns last error if method fails. -----------------------------------------------
Public Property Get ErrorString()
  ErrorString = sErr
End Property

'-- ==================== PRIVATE =========================

  '-- find out whether a value is a string or integer. return 0-int, 1-string, -1 if either table or Column is invalid.
Private Function ValType(sTabName, sColumn, sOp)
  Dim ACols, ATypes, iCol
     If (DicTables.exists(sTabName) = False) Then
       sErr = sOp & ": Invalid table name."
       ValType = -1
       Exit Function
     End If   
     
   ACols = DicTables.Item(sTabName)
     For iCol = 1 to UBound(ACols)
        If ACols(iCol) = sColumn Then
           ATypes = DicTables2.Item(sTabName)
           ValType = ATypes(iCol)
           Exit Function
       End If
    Next
      '-- if still going then Column name was invalid.
     sErr = sOp & ": Invalid Column name."
      ValType = -1       
End Function

  '-- create or delete temp folder for tables.
Private Sub TempFolOps(iOpFol)  '-- 0-set path. 1-delete. 2-create.
  Dim FSO3, oFol
    Set FSO3 = CreateObject("Scripting.FileSystemObject")
       If (iOpFol = 0) Then  '-- set path and quit.
          TabFol = FSO3.GetSpecialFolder(2) & "\msi-ed-fol"
           Set FSO3 = Nothing
           Exit Sub
       End If
           '-- delete folder for 1 or 2. If 2 then also make a new folder.
         If (FSO3.FolderExists(TabFol) = True) Then FSO3.DeleteFolder TabFol, True
     If (iOpFol = 2) Then    
        Set oFol = FSO3.CreateFolder(TabFol)
        Set oFol = Nothing
     End If   
   Set FSO3 = Nothing   
End Sub
  '------------------------------------ class ops ------------------------------------
Private Sub Class_Initialize()
    On Error Resume Next
  cApos = Chr(39)  '-- apostrophe.
  cCar = "`"
  BooLoaded = False
  TempFolOps 0  '-- set temp folder path.
 End Sub

Private Sub Class_Terminate()
    TempFolOps 1  '-- delete temp table folder before quitting.
   UnloadDB
 End Sub

End Class