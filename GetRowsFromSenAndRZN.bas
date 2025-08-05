Sub GetRowsFromSenAndRZN()

    If ActiveSheet.name  Разбор Then
        MsgBox Чтобы выполнить команду, откройте лист 'Разбор', vbInformation, Внимание!!!
        Exit Sub
    End If

    If sen_row = senheader_row Then
        MsgBox ОШИБКА sen_row ( & sen_row & ) не может быть = строки заголовка ( & senheader_row & ), vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Call GetColumnNumbers

    Dim ws As Worksheet Set ws = Worksheets(Разбор)

    ' Подготовка фильтров и сохранение его исходного состояния
    Dim originalFilterRow As Variant
    originalFilterRow = PrepareFilterRow(Worksheets(Разбор))

    ' Очистка старых данных
    Call DeleteExistingData(ws)

    ' Загрузка данных из всех источников
    Call LoadAndInsertRZN(Worksheets(Разбор))

    ' Восстановление строки фильтра
    Call RestoreFilterRow(Worksheets(Разбор), originalFilterRow)

    ' Включаем автофильтр и выделяем активную ячейку
    If Not Worksheets(Разбор).AutoFilterMode Then
        Worksheets(Разбор).Rows(senheader_row).AutoFilter
    End If
    Worksheets(Разбор).Cells(senheader_row + 1, column_adres).Select

    Application.ScreenUpdating = True
End Sub