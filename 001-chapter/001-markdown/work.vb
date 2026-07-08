Sub UpdateDocFromOldVersion(nameOldDoc)
    'Включение обработчика ошибок и снятие защиты документа.
    On Error Resume Next
    
    Dim newDoc As Workbook
    Dim oldDoc As Workbook
    
    Set newDoc = ActiveWorkbook
    Set oldDoc = Workbooks.Open(nameOldDoc, ReadOnly:=True)
    
    If (newDoc.Worksheets("Приложение №1").ProtectContents) Then
        newDoc.Worksheets("Приложение №1").Unprotect ("1111")
    End If
    If (newDoc.Worksheets("Приложение №2").ProtectContents) Then
        newDoc.Worksheets("Приложение №2").Unprotect ("1111")
    End If
    
    If (oldDoc.Worksheets("Приложение №1").ProtectContents) Then
        oldDoc.Worksheets("Приложение №1").Unprotect ("1111")
    End If
    If (oldDoc.Worksheets("Приложение №2").ProtectContents) Then
        oldDoc.Worksheets("Приложение №2").Unprotect ("1111")
    End If
    
    newDoc.Worksheets("Приложение №1").Activate
    
    'Приложение 1
    newDoc.Worksheets("Приложение №1").Cells(2, 7) = oldDoc.Worksheets("Приложение №1").Cells(2, 7)
    newDoc.Worksheets("Приложение №1").Cells(3, 7) = oldDoc.Worksheets("Приложение №1").Cells(3, 7)
    newDoc.Worksheets("Приложение №1").Cells(4, 7) = oldDoc.Worksheets("Приложение №1").Cells(4, 7)
    newDoc.Worksheets("Приложение №1").Cells(5, 7) = oldDoc.Worksheets("Приложение №1").Cells(5, 7)
    newDoc.Worksheets("Приложение №1").Cells(9, 1) = oldDoc.Worksheets("Приложение №1").Cells(9, 1)
    newDoc.Worksheets("Приложение №1").Cells(11, 2) = oldDoc.Worksheets("Приложение №1").Cells(11, 2)
    newDoc.Worksheets("Приложение №1").Cells(11, 5) = oldDoc.Worksheets("Приложение №1").Cells(11, 5)
    For x = 12 To 21
        newDoc.Worksheets("Приложение №1").Cells(x, 2) = oldDoc.Worksheets("Приложение №1").Cells(x, 2)
        newDoc.Worksheets("Приложение №1").Cells(x, 5) = oldDoc.Worksheets("Приложение №1").Cells(x, 5)
    Next x
    newDoc.Worksheets("Приложение №1").Cells(27, 5) = oldDoc.Worksheets("Приложение №1").Cells(27, 5)
    newDoc.Worksheets("Приложение №1").Cells(27, 6) = oldDoc.Worksheets("Приложение №1").Cells(27, 6)
    For x = 28 To 87
        For y = 2 To 7
            newDoc.Worksheets("Приложение №1").Cells(x, y) = oldDoc.Worksheets("Приложение №1").Cells(x, y)
        Next y
    Next x
    newDoc.Worksheets("Приложение №1").Cells(90, 3) = oldDoc.Worksheets("Приложение №1").Cells(90, 3)
    newDoc.Worksheets("Приложение №1").Cells(91, 3) = oldDoc.Worksheets("Приложение №1").Cells(91, 3)
    newDoc.Worksheets("Приложение №1").Cells(98, 5) = oldDoc.Worksheets("Приложение №1").Cells(98, 5)
    newDoc.Worksheets("Приложение №1").Cells(100, 1) = oldDoc.Worksheets("Приложение №1").Cells(100, 1)
    newDoc.Worksheets("Приложение №1").Cells(100, 5) = oldDoc.Worksheets("Приложение №1").Cells(100, 5)
    For I = 12 To 21
        newDoc.Worksheets("Приложение №1").Rows(I).Hidden = True
    Next I
    For I = 28 To 87
        newDoc.Worksheets("Приложение №1").Rows(I).Hidden = True
    Next I

    'Приложение 2
    newDoc.Worksheets("Приложение №2").Cells(9, 3) = oldDoc.Worksheets("Приложение №2").Cells(9, 3)
    newDoc.Worksheets("Приложение №2").Cells(9, 5).Value = oldDoc.Worksheets("Приложение №2").Cells(9, 5).Value
    newDoc.Worksheets("Приложение №2").Cells(9, 6).Value = oldDoc.Worksheets("Приложение №2").Cells(9, 6).Value
    newDoc.Worksheets("Приложение №2").Cells(9, 7).Value = oldDoc.Worksheets("Приложение №2").Cells(9, 7).Value
    For d = 10 To 69
        newDoc.Worksheets("Приложение №2").Cells(d, 2) = oldDoc.Worksheets("Приложение №2").Cells(d, 2)
        newDoc.Worksheets("Приложение №2").Cells(d, 3) = oldDoc.Worksheets("Приложение №2").Cells(d, 3)
        newDoc.Worksheets("Приложение №2").Cells(d, 4) = oldDoc.Worksheets("Приложение №2").Cells(d, 4)
        newDoc.Worksheets("Приложение №2").Cells(d, 5) = oldDoc.Worksheets("Приложение №2").Cells(d, 5)
    Next d
    newDoc.Worksheets("Приложение №2").Cells(76, 3).Value = oldDoc.Worksheets("Приложение №2").Cells(76, 3).Value

'    newDoc.Worksheets("Приложение №2").Rows(72).Hidden = True
'    newDoc.Worksheets("Приложение №2").Rows(73).Hidden = True
'    newDoc.Worksheets("Приложение №2").Rows(77).Hidden = True
'    newDoc.Worksheets("Приложение №2").Rows(74).Hidden = True
'    newDoc.Worksheets("Приложение №2").Rows(78).Hidden = True
'    newDoc.Worksheets("Приложение №2").Columns(5).Hidden = True
'    newDoc.Worksheets("Приложение №2").Columns(7).Hidden = True
'    newDoc.Worksheets("Приложение №2").Columns(8).Hidden = True
'    newDoc.Worksheets("Приложение №2").Columns(10).Hidden = True
    For I = 10 To 69
        newDoc.Worksheets("Приложение №2").Rows(I).Hidden = True
    Next I

    'Лист с текущими конфигурациями
    newDoc.Worksheets("List5").Cells(7, 1) = oldDoc.Worksheets("List5").Cells(7, 1)
    newDoc.Worksheets("List5").Cells(8, 1) = oldDoc.Worksheets("List5").Cells(8, 1)
    newDoc.Worksheets("List5").Cells(9, 1) = oldDoc.Worksheets("List5").Cells(9, 1)
    newDoc.Worksheets("List5").Cells(10, 1) = oldDoc.Worksheets("List5").Cells(10, 1)
    newDoc.Worksheets("List5").Cells(11, 1) = oldDoc.Worksheets("List5").Cells(11, 1)
    newDoc.Worksheets("List5").Cells(12, 1) = oldDoc.Worksheets("List5").Cells(12, 1)
    
    'Обработчик ошибок.
    Dim p As Integer
    If Err.Number <> 0 Then
        p = Err.Number
        Err.Clear
        If p <> 5941 Then
            MsgBox "Ошибка " & p & " очищена! Ошибка могла повлиять на заполнение документа, рекомендуется сообщить о возникновение ошибки (назвав её номер) системному администратору.", vbInformation, "Уведомление"
        End If
    End If
    
    'Установка защиты документа и закрытие формы.
    If Not (newDoc.Worksheets("Приложение №1").ProtectContents) Then
        newDoc.Worksheets("Приложение №1").Protect ("1111")
    End If
    If Not (newDoc.Worksheets("Приложение №2").ProtectContents) Then
        newDoc.Worksheets("Приложение №2").Protect ("1111")
    End If
    
    If Not (oldDoc.Worksheets("Приложение №1").ProtectContents) Then
        oldDoc.Worksheets("Приложение №1").Protect ("1111")
    End If
    If Not (oldDoc.Worksheets("Приложение №2").ProtectContents) Then
        oldDoc.Worksheets("Приложение №2").Protect ("1111")
    End If
    
    newDoc.Worksheets("Приложение №1").Activate
    
    oldDoc.Close False
End Sub




