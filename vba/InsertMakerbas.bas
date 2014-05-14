Attribute VB_Name = "Module1"
Option Explicit

Const COL_NAME_ROW = 1
Const COL_TYPE_ROW = 2
Const ROW_START = 3
Const COL_START = 1

'Const TEMPLATE = "Insert into @tableName ( @colNames ) values ( @values );"
Const TEMPLATE = "Insert into @tableName ( @colNames ) values "
Const DEL_TEMPLATE = "Delete from @tableName;"

Public Sub Insert作成()
Attribute Insert作成.VB_ProcData.VB_Invoke_Func = "i\n14"

    Application.ScreenUpdating = False
    
    Dim colTypeMacher As Object
    Set colTypeMacher = CreateObject("VBScript.RegExp")
    colTypeMacher.Pattern = "\(.*\)([ \t]*unsigned)?$"
    
    ' 区分値などに内容説明を付けて表記した場合の説明部分の除去する正規表現
    ' 例）100:回復アイテム
    Dim descTrimer As Object
    Set descTrimer = CreateObject("VBScript.RegExp")
    descTrimer.Pattern = ":.*$"
    
    ' SQL関数定義であるか調査するための正規表現
    ' 例）now()
    Dim funcChecker As Object
    Set funcChecker = CreateObject("VBScript.RegExp")
    funcChecker.Pattern = "^\w+\(.*\).*"
    
    ' 件数が多くてシートが２つに分かれても耐えられるように末尾の番号を除去する正規表現
    ' 例）items(2), items(201405補点)
    Dim tblNameTrimer As Object
    Set tblNameTrimer = CreateObject("VBScript.RegExp")
    tblNameTrimer.Pattern = "[(（].*[)）]"
    
    Dim lineFeedConvert As Object
    Set lineFeedConvert = CreateObject("VBScript.RegExp")
    lineFeedConvert.Pattern = "\n"
    lineFeedConvert.Global = True
    
    Dim row As Long
    Dim col As Integer
    Dim dataEndCol As Integer
    row = ROW_START
    col = COL_START

    Dim tableName As String
    tableName = tblNameTrimer.Replace(ActiveSheet.name, "")
    
    Dim colNames As String
    Dim typeNames As Collection
    Set typeNames = New Collection
    
    typeNames.Add colTypeMacher.Replace(Cells(COL_TYPE_ROW, col).Value, "")
    colNames = Cells(COL_NAME_ROW, col).Value
    col = col + 1
    
    While (Cells(COL_TYPE_ROW, col).Value <> "")
        typeNames.Add colTypeMacher.Replace(Cells(COL_TYPE_ROW, col).Value, "")
        colNames = colNames & "," & Cells(COL_NAME_ROW, col).Value
        col = col + 1
    Wend
    
    Dim insertFormat As String
    insertFormat = Replace(Replace(TEMPLATE, "@tableName", tableName), "@colNames", colNames)
    
    Cells(1, col + 1).Value = insertFormat
    Cells(2, col + 1).Value = Replace(DEL_TEMPLATE, "@tableName", tableName)
    
    dataEndCol = col - 1
    row = ROW_START
    col = COL_START
    
    Dim bulkInsCnt As Integer
    bulkInsCnt = 1
    Dim values As String
    While (Cells(row, COL_START).Value <> "end")
    
        col = COL_START
        
        values = getValue(row, col, typeNames, descTrimer, funcChecker, lineFeedConvert)
        col = col + 1
        While (col <= dataEndCol)
            values = values & "," & getValue(row, col, typeNames, descTrimer, funcChecker, lineFeedConvert)
            col = col + 1
        Wend
       
        Dim endChar As String
        If Cells(row + 1, COL_START).Value <> "end" And bulkInsCnt < 100 Then
            endChar = ","
        Else
            endChar = ";"
        End If
        
        If bulkInsCnt = 1 Then
            Cells(row, dataEndCol + 1).Value = insertFormat & "( " & values & " )" & endChar
        Else
            Cells(row, dataEndCol + 1).Value = "( " & values & " )" & endChar
        End If
        
        ' 一括登録行数のカウンタを繰り上げ（100行単位なので100を超えたらリセット）
        bulkInsCnt = bulkInsCnt + 1
        If bulkInsCnt > 100 Then
            bulkInsCnt = 1
        End If
        
        row = row + 1
    Wend
    
    Application.ScreenUpdating = True
End Sub



Private Function getValue(row As Long, col As Integer, typeNames As Collection, trimer As Object, funcChecker As Object, lineFeedConvert As Object) As String

    Dim val As String
    val = trimer.Replace(Cells(row, col).Value, "")

    If val = "" Then
        getValue = "Null"
        Exit Function
'    ElseIf Right(Cells(row, col).Value, 2) = "()" Then
    ElseIf funcChecker.Test(val) Then
        getValue = Cells(row, col).Value
        Exit Function
    End If
    
    Dim typeName As String
    typeName = LCase(typeNames.Item(col))
    
    If typeName = "文字" Then
        getValue = "'" & lineFeedConvert.Replace(Cells(row, col).Value, "\n") & "'"
    ElseIf typeName = "char" Then
        getValue = "'" & lineFeedConvert.Replace(Cells(row, col).Value, "\n") & "'"
    ElseIf typeName = "varchar" Then
        getValue = "'" & lineFeedConvert.Replace(Cells(row, col).Value, "\n") & "'"
    ElseIf typeName = "日付" Then
        getValue = "'" & Cells(row, col).Value & "'"
    ElseIf typeName = "datetime" Then
        getValue = "'" & Cells(row, col).Value & "'"
    ElseIf typeName = "time" Then
        getValue = "'" & Cells(row, col).Value & "'"
    ElseIf typeName = "数値" Then
        getValue = val
    ElseIf typeName = "int" Then
        getValue = val
    ElseIf typeName = "integer" Then
        getValue = val
    ElseIf typeName = "tinyint" Then
        getValue = val
    ElseIf typeName = "mediumint" Then
        getValue = val
    ElseIf typeName = "smallint" Then
        getValue = val
    ElseIf typeName = "bigint" Then
        getValue = val
    End If

End Function

