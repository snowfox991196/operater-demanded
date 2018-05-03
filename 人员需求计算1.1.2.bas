Attribute VB_Name = "��Ա�������"
Public Function findProcess(rangeM As Range) As String      '��ѯ��Ӧ������
    Dim tmp As String, columnM As Integer
    columnM = rangeM.column
    tmp = Cells(3, columnM).Text
    findProcess = tmp
End Function
Public Function findProduct(rangeM As Range) As String      '��ѯ��ӦƷ�溯��
    Dim tmp As String
    Const productTableMaxColumn As Integer = 50     '�˴�50ΪƷ����������������������Ʒ�泬��50�����޸Ĵ˴�
    For i = 1 To productTableMaxColumn
        If rangeM.Interior.Color = Cells(2, i).Interior.Color Then
            tmp = Cells(2, i).Text
            Exit For
        End If
    Next i
    findProduct = tmp
End Function

Public Function findShift(rangeM As Range) As String        '��ѯ��Ӧ��κ���
    findShift = Range("F" & rangeM.Row).Text
End Function

Public Function Semp(rangeM As Range) As Single       '��ѯ��Ӧ������������
    Dim indexTable As String
    indexTable = Replace("��Ա���ݿ⣨" & ActiveSheet.Name & "��", " ", "")     '�Զ���ȡ��Ա���ݿ�����ƣ���������ΪXXX�������ݿ�Ӧ����Ϊ ��Ա���ݿ⣨XXX����
    Const indexColumn As Integer = 100      '�˴�100Ϊ��ѯ���������˹�����100�У��޸Ĵ˴�
    Const indexShift As Integer = 20      '�˴�20Ϊ�ܹ��������������ӹ�����γ���20�У��޸Ĵ˴�
    Dim tmp As Single
    If rangeM.Text = "" Then
        tmp = 0
    ElseIf Not IsNumeric(rangeM.Text) Then      '�����峡��໻ģ�����⹤������
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
    Else            '�����������������
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
