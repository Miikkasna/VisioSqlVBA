Option Explicit
Sub generate_sql()
    Dim fields, itm, fieldData, pg, shp
    Dim tables(10000) As String, keys(10000, 3)
    Dim tblCnt As Integer, keyCnt As Integer
    Dim field As String, key As String, tblName As String, dType As String, tblSql As String, spaces As String, fName As String
    
    tblCnt = 0
    keyCnt = 0
    ' go through pages
    For Each pg In ActiveDocument.Pages
        fName = pg.Name & ".sql"
        Open fName For Output As #1
        ' go through shapes (tables)
        For Each shp In pg.Shapes
            If shp.Style <> "Normal" Then GoTo nextshp ' ignore non table shapes
            tblName = shp.Shapes.Item(1).Text
            If IsInArray(tblName, tables) Then GoTo nextshp ' ignore duplicate tables
            tables(tblCnt) = tblName
            tblCnt = tblCnt + 1
            tblSql = "-- " & tblName & " -table" & vbNewLine & "CREATE TABLE " & tblName & "("
            fields = shp.Shapes.Item(2).Text
            ' go through table fields
            For Each itm In Split(fields, vbLf)
                fieldData = Split(CStr(itm), Chr(9))
                If UBound(fieldData) = -1 Then GoTo nextitm
                field = fieldData(1)
                dType = fieldData(2)
                key = Left(fieldData(0), 2)
                If key = "FK" Or key = "PK" Then
                    If key = "PK" Then dType = dType & " PRIMARY KEY"
                    keys(keyCnt, 0) = tblName
                    keys(keyCnt, 1) = key
                    keys(keyCnt, 2) = field
                    keyCnt = keyCnt + 1
                End If
        
                If Len(field) >= 32 Then
                    spaces = 1
                Else
                    spaces = addSpaces(32 - Len(field))
                End If
                tblSql = tblSql & vbNewLine & vbTab & field & spaces & dType & ","
nextitm:
            Next itm
            tblSql = Left(tblSql, Len(tblSql) - 1) ' remove last comma
            tblSql = tblSql & vbNewLine & ");"
             Print #1, tblSql & vbNewLine & vbNewLine
nextshp:
        Next shp
        Close #1
    Next pg
    createForeignKeys keys
  
  End Sub
Public Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If IsEmpty(arr(i)) Then Exit For
        If arr(i) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function
Function addSpaces(ByVal n As Integer) As String
    Dim spaces As String, i
    spaces = ""
    For i = 1 To n
        spaces = spaces & " "
    Next i
    addSpaces = spaces
End Function
Sub createForeignKeys(ByRef keys)
    Dim i, ii
    Dim tbl As String, fld As String, key As String, fName As String, sqlStr As String
    
    fName = "Foreign Keys.sql"
    Open fName For Output As #1
    For i = 0 To UBound(keys, 1)
        If IsEmpty(keys(i, 0)) Then Exit For
        tbl = keys(i, 0)
        key = keys(i, 1)
        fld = keys(i, 2)
        If keys(i, 1) = "FK" Then
            For ii = 0 To UBound(keys, 1)
                If IsEmpty(keys(ii, 0)) Then Exit For
                If keys(ii, 1) = "PK" And fld = keys(ii, 2) Then
                    sqlStr = "-- foreign key between " & tbl & " and " & keys(ii, 0) & " at " & fld & vbNewLine
                    sqlStr = sqlStr & "ALTER TABLE " & tbl & " ADD CONSTRAINT FK_" & tbl & "_" & keys(ii, 0) & "_" & fld & vbNewLine
                    sqlStr = sqlStr & "FOREIGN KEY (" & fld & ") REFERENCES " & keys(ii, 0) & "(" & fld & ");" & vbNewLine
                    Print #1, sqlStr
                End If
            Next ii
        End If
    Next i
    Close #1
End Sub
