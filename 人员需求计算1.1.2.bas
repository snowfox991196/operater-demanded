Attribute VB_Name = "人员需求计算"
Public Function findProcess(rangeM As Range) As String      '查询对应工序函数
    Dim tmp As String, columnM As Integer
    columnM = rangeM.column
    tmp = Cells(3, columnM).Text
    findProcess = tmp
End Function
Public Function findProduct(rangeM As Range) As String      '查询对应品规函数
    Dim tmp As String
    Const productTableMaxColumn As Integer = 50     '此处50为品规所在行最大列数，如添加品规超过50列请修改此处
    For i = 1 To productTableMaxColumn
        If rangeM.Interior.Color = Cells(2, i).Interior.Color Then
            tmp = Cells(2, i).Text
            Exit For
        End If
    Next i
    findProduct = tmp
End Function

Public Function findShift(rangeM As Range) As String        '查询对应班次函数
    findShift = Range("F" & rangeM.Row).Text
End Function

Public Function Semp(rangeM As Range) As Single       '查询对应需求人数函数
    Dim indexTable As String
    indexTable = Replace("人员数据库（" & ActiveSheet.Name & "）", " ", "")     '自动获取人员数据库表名称，若本表名为XXX，则数据库应命名为 人员数据库（XXX）。
    Const indexColumn As Integer = 100      '此处100为查询行数，如人工表超过100行，修改此处
    Const indexShift As Integer = 20      '此处20为总工序列数，如增加工序或班次超过20列，修改此处
    Dim tmp As Single
    If rangeM.Text = "" Then
        tmp = 0
    ElseIf Not IsNumeric(rangeM.Text) Then      '处理清场清洁换模具特殊工序代码段
        For j = 1 To indexShift
            If rangeM.Text = Sheets(indexTable).Cells(1, j).Text Then
                For i = 1 To indexColumn
                    If Sheets(indexTable).Range("A" & i).Text = findProduct(rangeM) And Sheets(indexTable).Range("B" & i).Text = findProcess(rangeM) Then
                        tmp = Sheets(indexTable).Cells(i, j).Text
                        Exit For
                    End If
                Next i
                Exit For
            End If
        Next j
    Else            '处理正常生产代码段
        For i = 1 To indexColumn
            If Sheets(indexTable).Range("A" & i).Text = findProduct(rangeM) And Sheets(indexTable).Range("B" & i).Text = findProcess(rangeM) Then
                For j = 1 To indexShift
                    If findShift(rangeM) = Sheets(indexTable).Cells(1, j).Text Then
                        tmp = Sheets(indexTable).Cells(i, j).Text
                        Exit For
                    End If
                Next j
                Exit For
            End If
        Next i
    End If
    Semp = tmp
End Function
