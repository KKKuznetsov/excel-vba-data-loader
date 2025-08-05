Function BuildCensusAndRZNsql(ws As Worksheet, Optional maxRows As Long = 1000) As String
    Dim sql As String, val As String, sen As String, rzn As String
    Dim filtersSen As String, filtersRZN As String

    filtersSen = "WHERE 1=1 "
    filtersRZN = "WHERE 1=1 "

    val = ws.Cells(senfil_row, column_inn).Value
    If IsMeaningfulFilter(val) Then
        filtersSen = filtersSen & "AND sen.[ИНН РНС] = '" & val & "' "
        filtersRZN = filtersRZN & "AND rzn.[rzn_inn] = '" & val & "' "
    End If

    val = ws.Cells(senfil_row, column_nazvanie_apt).Value
    val = Replace(Replace(val, "ё", "е"), "Ё", "Е")
    If IsMeaningfulFilter(val) Then filtersSen = filtersSen & "AND sen.[№аптеки] LIKE '%" & val & "%' "

    val = ws.Cells(senfil_row, column_ul).Value
    val = Replace(Replace(val, "ё", "е"), "Ё", "Е")
    If IsMeaningfulFilter(val) Then
        filtersSen = filtersSen & "AND (sen.[Юрлицо РНС] LIKE '%" & val & "%' OR sen.[Полное название ЮЛ] LIKE '%" & val & "%') "
        filtersRZN = filtersRZN & "AND (rzn.[rzn_abbreviated_name_licensee] LIKE '%" & val & "%' OR rzn.[rzn_full_name_licensee] LIKE '%" & val & "%') "
    End If

    val = ws.Cells(senfil_row, column_adres).Value
    val = Replace(Replace(val, "ё", "е"), "Ё", "Е")
    If IsMeaningfulFilter(val) Then
        filtersSen = filtersSen & "AND CONCAT(sen.[Адрес РНС], ' ', sen.[Дополнение к адресу]) LIKE '%" & val & "%' "
        filtersRZN = filtersRZN & "AND rzn.[rzn_address] LIKE '%" & val & "%' "
    End If

    val = ws.Cells(senfil_row, column_apt_region).Value
    val = Replace(Replace(val, "ё", "е"), "Ё", "Е")
    If IsMeaningfulFilter(val) Then filtersSen = filtersSen & "AND sen.[Субъект] LIKE '%" & val & "%' "

    val = ws.Cells(senfil_row, column_apt_city).Value
    val = Replace(Replace(val, "ё", "е"), "Ё", "Е")
    If IsMeaningfulFilter(val) Then filtersSen = filtersSen & "AND sen.[Населенный пункт] LIKE '%" & val & "%' "

    filtersRZN = filtersRZN & "AND rzn.[rzn_address] IS NOT NULL "
    filtersRZN = filtersRZN & "AND rzn.[rzn_activity_type] IS NOT NULL "
    filtersRZN = filtersRZN & "AND rzn.[rzn_work_full] IS NOT NULL "

    sen = _
        "sen.[ID РНС], " & _
        "sen.[ID сети], " & _
        "sen.[Дата закрытия], " & _
        "sen.[Юрлицо РНС], " & _
        "sen.[ИНН РНС], " & _
        "sen.[№аптеки], " & _
        "sen.[Адрес РНС], " & _
        "sen.[Дополнение к адресу], " & _
        "sen.[Субъект], " & _
        "sen.[Населенный пункт], " & _
        "sen.[Муниципальный район], " & _
        "sen.[Административный округ Москвы], " & _
        "sen.[Тип учреждения, детализация], " & _
        "sen.[Направление точки продаж], " & _
        "sen.[Комментарий], " & _
        "[UL]                = cast(NULL as varchar), " & _
        "[rzn_inn]           = cast(NULL as varchar), " & _
        "[rzn_address]       = cast(NULL as varchar), " & _
        "[rzn_activity_type] = cast(NULL as varchar), " & _
        "[rzn_work_full]     = cast(NULL as varchar), " & _
        "[Дата с]            = cast(NULL as date), " & _
        "[Дата до]           = cast(NULL as date), "

    rzn = _
        "[ID РНС]                        = cast(NULL as int), " & _
        "[ID сети]                       = ISNULL(sen.[ID сети], 'ИНН нет в сенсусе'), " & _
        "[Дата закрытия]                 = cast(NULL as date), " & _
        "[Юрлицо РНС]                    = cast(NULL as varchar), " & _
        "[ИНН РНС]                       = rzn.[rzn_inn], " & _
        "[№аптеки]                       = cast(NULL as varchar), " & _
        "[Адрес РНС]                     = rzn.[rzn_address], " & _
        "[Дополнение к адресу]           = cast(NULL as varchar), " & _
        "[Субъект]                       = cast(NULL as varchar), " & _
        "[Населенный пункт]              = cast(NULL as varchar), " & _
        "[Муниципальный район]           = cast(NULL as varchar), " & _
        "[Административный округ Москвы] = cast(NULL as varchar), " & _
        "[Тип учреждения, детализация]   = cast(NULL as varchar), " & _
        "[Направление точки продаж]      = cast(NULL as varchar), " & _
        "[Комментарий]                   = cast(NULL as varchar), " & _
        "rzn.[UL], " & _
        "rzn.[rzn_inn], " & _
        "rzn.[rzn_address], " & _
        "rzn.[rzn_activity_type], " & _
        "rzn.[rzn_work_full], " & _
        "rzn.[Дата с], " & _
        "rzn.[Дата до], "
    
    sql = _
        "SET NOCOUNT ON" & vbCrLf & _
        "DROP TABLE IF EXISTS #sen" & vbCrLf & _
        "SELECT " & sen & "1 AS [source_order], 'Сенсус' AS [Источник], NULL AS [id_rnc_new]" & vbCrLf & _
        "INTO #sen " & vbCrLf & _
        "FROM [SSA].[dbo].[Сенсус клиентов] sen " & vbCrLf & _
        "LEFT JOIN [SSA].[dbo].[Ascensia_id_rnc_replace] idrep ON " & _
        "idrep.[id_rnc_old] = sen.[ID РНС] " & vbCrLf & _
        filtersSen & "AND sen.[ID РНС] < 1000000000 AND idrep.[id_rnc_old] IS NULL" & vbCrLf & _
        "OPTION (MAXDOP 100)" & vbCrLf & vbCrLf & _
        "DROP TABLE IF EXISTS #del" & vbCrLf & _
        "SELECT " & sen & "2 AS [source_order], 'Сенсус удаленное' AS [Источник], idrep.[id_rnc_new]" & vbCrLf & _
        "INTO #del " & vbCrLf & _
        "FROM [SSA].[dbo].[Сенсус клиентов удаленное] sen " & vbCrLf & _
        "LEFT JOIN [SSA].[dbo].[Ascensia_id_rnc_replace] idrep ON " & _
        "idrep.[id_rnc_old] = sen.[ID РНС] " & vbCrLf & _
        filtersSen & "AND sen.[ID РНС] < 1000000000 AND idrep.[id_rnc_old] IS NOT NULL" & vbCrLf & _
        "OPTION (MAXDOP 100)" & vbCrLf & vbCrLf & _
        "DROP TABLE IF EXISTS #rzn" & vbCrLf & _
        "SELECT " & rzn & "3 AS [source_order], 'Росздравнадзор' AS [Источник], NULL AS [id_rnc_new]" & vbCrLf & _
        "INTO #rzn " & vbCrLf & _
        "FROM [uvp_rzn].[dbo].[rzn_data_ret_grp] rzn " & vbCrLf & _
        "LEFT JOIN (select [ID сети] = max([ID сети]), [ИНН РНС] from [SSA].[dbo].[Сенсус клиентов] group by [ИНН РНС]) sen ON sen.[ИНН РНС] = rzn.[rzn_inn] " & vbCrLf & _
        filtersRZN & vbCrLf & _
        "OPTION (MAXDOP 100)" & vbCrLf & vbCrLf
    
    sql = sql & _
        "SELECT TOP " & maxRows + 1 & " * FROM (" & vbCrLf & _
        "SELECT * FROM #sen" & vbCrLf & _
        "UNION ALL" & vbCrLf & _
        "SELECT * FROM #del" & vbCrLf & _
        "UNION ALL" & vbCrLf & _
        "SELECT * FROM #rzn" & vbCrLf & _
        ") full_data ORDER BY source_order, [Адрес РНС], [Юрлицо РНС]"

    BuildCensusAndRZNsql = sql
End Function