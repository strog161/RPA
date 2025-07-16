Attribute VB_Name = "Calculation"
Function calk(ByVal sheetName As String) As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim maxCost As Double
    Dim minCost As Double
    Dim maxCostRow As Long
    Dim minCostRow As Long
    Dim sheetNameStr As String
	
	sheetNameStr = sheetName
	
	'���������, ���������� �� ����
	On Error Resume Next
	Set ws = ThisWorkbook.Sheets(sheetNameStr)
	On Error GoTo 0
		
    If ws Is Nothing Then
    MsgBox "���� � ������ '" & sheetNameStr & "' �� ������!", vbExclamation
    calk = "false"
    Exit Function
End If
	
    ' ������� ��������� ������ � �������
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' ��������� ��������� ��� ����� ������� "����� ���������"
    ws.Cells(1, 5).Value = "����� ���������"
    
    ' �������� ��������� ��������� ������
    ws.Range("A1:E1").Interior.Color = RGB(255, 165, 0)
    
    ' �������� �� ������ ������ � ������������ ����� ���������
    For i = 2 To lastRow
        ' ��������� �� ������ � ������������ ��������
        If IsEmpty(ws.Cells(i, 3)) Or IsEmpty(ws.Cells(i, 4)) Then
            ws.Cells(i, 5).Value = "0"
        Else
            ' ��������� �� ������������� �������� � ������ �� �� �������������
            If ws.Cells(i, 3).Value < 0 Then
                ws.Cells(i, 3).Value = Abs(ws.Cells(i, 3).Value)
            End If
            If ws.Cells(i, 4).Value < 0 Then
                ws.Cells(i, 4).Value = Abs(ws.Cells(i, 4).Value)
            End If
            
            ' ������������ ����� ���������
            ws.Cells(i, 5).Value = ws.Cells(i, 3).Value * ws.Cells(i, 4).Value
        End If
    Next i
    
    ' ��������� ������� �� �������� ������ � ���������� �������
    ws.Range("A2:E" & lastRow).Sort Key1:=ws.Range("B2"), Order1:=xlAscending, Header:=xlNo
    
    ' ������� ������ � ������������ � ����������� ����� ����������
    maxCost = Application.WorksheetFunction.Max(ws.Range("E2:E" & lastRow))
    minCost = Application.WorksheetFunction.Min(ws.Range("E2:E" & lastRow))
    
    For i = 2 To lastRow
        If ws.Cells(i, 5).Value = maxCost Then
            maxCostRow = i
        End If
        If ws.Cells(i, 5).Value = minCost Then
            minCostRow = i
        End If
    Next i
    
    ' �������� ������ � ������������ ����� ���������� �������
	ws.Range("A" & maxCostRow & ":E" & maxCostRow).Interior.Color = RGB(0, 255, 0)
    
    ' �������� ������ � ����������� ����� ���������� �������
	ws.Range("A" & minCostRow & ":E" & minCostRow).Interior.Color = RGB(255, 0, 0) 
	
	calk = "true"
	
End Function
