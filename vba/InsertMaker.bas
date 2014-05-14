Attribute VB_Name = "Module1"
Option Explicit

Const COL_NAME_ROW = 1
Const COL_TYPE_ROW = 2
Const ROW_START = 3
Const COL_START = 1

'Const TEMPLATE = "Insert into @tableName ( @colNames ) values ( @values );"
Const TEMPLATE = "Insert into @tableName ( @colNames ) values "
Const DEL_TEMPLATE = "Delete from @tableName;"

Public Sub Insert�쐬()
Attribute Insert�쐬.VB_ProcData.VB_Invoke_Func = "i\n14"

    Application.ScreenUpdating = False
    
    Dim colTypeMacher As Object
    Set colTypeMacher = CreateObject("VBScript.RegExp")
    colTypeMacher.Pattern = "\(.*\)([ \t]*unsigned)?$"
    
    ' �敪�l�Ȃǂɓ��e������t���ĕ\�L�����ꍇ�̐��������̏������鐳�K�\��
    ' ��j100:�񕜃A�C�e��
    Dim descTrimer As Object
    Set descTrimer = CreateObject("VBScript.RegExp")
    descTrimer.Pattern = ":.*$"
    
    ' SQL�֐���`�ł��邩�������邽�߂̐��K�\��
    ' ��jnow()
    Dim funcChecker As Object
    Set funcChecker = CreateObject("VBScript.RegExp")
    funcChecker.Pattern = "^\w+\(.*\).*"
    
    ' �����������ăV�[�g���Q�ɕ�����Ă��ς�����悤�ɖ����̔ԍ����������鐳�K�\��
    ' ��jitems(2), items(201405��_)
    Dim tblNameTrimer As Object
    Set tblNameTrimer = CreateObject("VBScript.RegExp")
    tblNameTrimer.Pattern = "[(�i].*[)�j]"
    
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
        
        ' �ꊇ�o�^�s���̃J�E���^���J��グ�i100�s�P�ʂȂ̂�100�𒴂����烊�Z�b�g�j
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
    
    If typeName = "����" Then
        getValue = "'" & lineFeedConvert.Replace(Cells(row, col).Value, "\n") & "'"
    ElseIf typeName = "char" Then
        getValue = "'" & lineFeedConvert.Replace(Cells(row, col).Value, "\n") & "'"
    ElseIf typeName = "varchar" Then
        getValue = "'" & lineFeedConvert.Replace(Cells(row, col).Value, "\n") & "'"
    ElseIf typeName = "���t" Then
        getValue = "'" & Cells(row, col).Value & "'"
    ElseIf typeName = "datetime" Then
        getValue = "'" & Cells(row, col).Value & "'"
    ElseIf typeName = "time" Then
        getValue = "'" & Cells(row, col).Value & "'"
    ElseIf typeName = "���l" Then
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

