Sub LoadAndInsertRZN(ws As Worksheet)
    Dim cn As Object: Set cn = CreateObject("ADODB.Connection")
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    ' Подключение к базе данных для РЗН
    cn.ConnectionString = "Provider=SQLOLEDB;Data Source=***;Initial Catalog=ssa;User ID=***;Password=***;"
    cn.ConnectionTimeout = 60
    cn.CommandTimeout = 360
    cn.Open

    Dim maxRows As Long
    maxRows = Worksheets("Настройки").Range("B1").Value
    If maxRows > 1000 Then maxRows = 1000

    ' Формирование SQL-запроса для РЗН
    Dim sql As String
    sql = BuildCensusAndRZNsql(ws, maxRows)

    rs.Open sql, cn

    ' ?? ДОБАВИТЬ ВОТ ЭТУ СТРОКУ
    Call GetColumnNumbers

    ' Вставка данных из РЗН
    Call InsertRZNData(ws, rs, "Росздравнадзор")

    rs.Close
    cn.Close
End Sub