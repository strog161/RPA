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
	
	'Проверяем, существует ли лист
	On Error Resume Next
	Set ws = ThisWorkbook.Sheets(sheetNameStr)
	On Error GoTo 0
		
    If ws Is Nothing Then
    MsgBox "Лист с именем '" & sheetNameStr & "' не найден!", vbExclamation
    calk = "false"
    Exit Function
End If
	
    ' Находим последнюю строку с данными
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Добавляем заголовок для новой колонки "Общая стоимость"
    ws.Cells(1, 5).Value = "Общая стоимость"
    
    ' Выделяем заголовки оранжевым цветом
    ws.Range("A1:E1").Interior.Color = RGB(255, 165, 0)
    
    ' Проходим по каждой строке и рассчитываем общую стоимость
    For i = 2 To lastRow
        ' Проверяем на пустые и некорректные значения
        If IsEmpty(ws.Cells(i, 3)) Or IsEmpty(ws.Cells(i, 4)) Then
            ws.Cells(i, 5).Value = "0"
        Else
            ' Проверяем на отрицательные значения и меняем их на положительные
            If ws.Cells(i, 3).Value < 0 Then
                ws.Cells(i, 3).Value = Abs(ws.Cells(i, 3).Value)
            End If
            If ws.Cells(i, 4).Value < 0 Then
                ws.Cells(i, 4).Value = Abs(ws.Cells(i, 4).Value)
            End If
            
            ' Рассчитываем общую стоимость
            ws.Cells(i, 5).Value = ws.Cells(i, 3).Value * ws.Cells(i, 4).Value
        End If
    Next i
    
    ' Сортируем таблицу по названию товара в алфавитном порядке
    ws.Range("A2:E" & lastRow).Sort Key1:=ws.Range("B2"), Order1:=xlAscending, Header:=xlNo
    
    ' Находим строки с максимальной и минимальной общей стоимостью
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
    
    ' Выделяем строку с максимальной общей стоимостью зеленым
	ws.Range("A" & maxCostRow & ":E" & maxCostRow).Interior.Color = RGB(0, 255, 0)
    
    ' Выделяем строку с минимальной общей стоимостью красным
	ws.Range("A" & minCostRow & ":E" & minCostRow).Interior.Color = RGB(255, 0, 0) 
	
	calk = "true"
	
End Function
