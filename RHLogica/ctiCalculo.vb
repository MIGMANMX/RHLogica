Imports System.Data.SqlClient
Public Class ctiCalculo
    'Calculo de Horas
    '''Gridview Chequeo
    Public Function gvChequeo(ByVal idempleado As Integer, ByVal Fech1 As String, ByVal Fech2 As String) As DataTable
        Dim dt As New DataTable
        dt.Columns.Add(New DataColumn("chec", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("tipo", System.Type.GetType("System.String")))

        Dim r As DataRow
        Dim dbC As New SqlConnection(StarTconnStrRH)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT chec, tipo FROM Chequeo Where chec between '" & Fech1 & "' and '" & Fech2 & "' AND idempleado=@idempleado   ORDER BY chec", dbC)

        cmd.Parameters.AddWithValue("idempleado", idempleado)
        'cmd.Parameters.AddWithValue("idempleado", Fech1)
        'cmd.Parameters.AddWithValue("idempleado", Fech2)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            r = dt.NewRow
            r(0) = rdr("chec").ToString : r(1) = rdr("tipo").ToString
            dt.Rows.Add(r)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dt
    End Function
    Public Function datosHora(ByVal idchequeo As Integer) As String()
        Dim dbC As New SqlConnection(StarTconnStrRH)
        dbC.Open()
        Dim cmd As New SqlCommand("Select Convert(varchar(10),chec, 8)as chec  from Chequeo  where idchequeo=@idchequeo ", dbC)
        cmd.Parameters.AddWithValue("idchequeo", idchequeo)

        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim dsP As String()
        If rdr.Read Then
            ReDim dsP(1)
            dsP(0) = rdr("chec").ToString
        Else
            ReDim dsP(0) : dsP(0) = "Error: no se encuentra."
        End If
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dsP
    End Function
    Public Function agregarHora(ByVal fech As Date, ByVal idempleado As Integer, ByVal horain As Date, ByVal horafn As Date, ByVal horas As TimeSpan) As String()
        Dim ans As String()

        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("INSERT INTO Horas SELECT @idempleado,@horain,@horafn,@horas,@fecha", dbC)
        cmd.Parameters.AddWithValue("fecha", fech)
        cmd.Parameters.AddWithValue("idempleado", idempleado)
        cmd.Parameters.AddWithValue("horain", horain)
        cmd.Parameters.AddWithValue("horafn", horafn)
        cmd.Parameters.AddWithValue("horas", horas)
        cmd.ExecuteNonQuery()
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        cmd.CommandText = "SELECT idHoras FROM Horas WHERE fecha = @fecha"
        rdr = cmd.ExecuteReader
        rdr.Read()
        ReDim ans(1)
        ans(0) = "Agregado."
        ans(1) = rdr("idpuesto").ToString
        rdr.Close()

        'If rdr.HasRows Then
        '    ReDim ans(0)
        '    ans(0) = "Error: no se puede agregar, ya existe."
        '    rdr.Close()
        'Else
        '    rdr.Close()
        '    cmd.CommandText = "INSERT INTO Horas SELECT @idempleado,@horain,@horafn,@horas,@fecha"
        '    cmd.Parameters.AddWithValue("idempleado", idempleado)
        '    cmd.Parameters.AddWithValue("horain", horain)
        '    cmd.Parameters.AddWithValue("horafn", horafn)
        '    cmd.Parameters.AddWithValue("horas", horas)

        '    cmd.ExecuteNonQuery()
        '    cmd.CommandText = "SELECT idHoras FROM Horas WHERE fecha = @fecha"
        '    rdr = cmd.ExecuteReader
        '    rdr.Read()
        '    ReDim ans(1)
        '    ans(0) = "Agregado."
        '    ans(1) = rdr("idpuesto").ToString
        '    rdr.Close()
        'End If
        rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()

        Return ans
    End Function
    'Public Function datosHTotales(ByVal idempleado As Integer, ByVal chec As Date) As String()
    '    Dim dbC As New SqlConnection(StarTconnStrRH)
    '    dbC.Open()
    '    Dim cmd As New SqlCommand("Select * From Chequeo where chec>=@chec AND chec <= '" & DateAdd(DateInterval.Day, 1, chec).ToString("yyyy-dd-MM") & "' AND idempleado=@idempleado Order BY chec ", dbC)
    '    cmd.Parameters.AddWithValue("idempleado", idempleado)
    '    cmd.Parameters.AddWithValue("chec", chec)
    '    Dim rdr As SqlDataReader = cmd.ExecuteReader
    '    Dim dsP As String()
    '    If rdr.Read Then

    '        ReDim dsP(3)
    '        dsP(0) = rdr("idchequeo").ToString
    '        dsP(1) = rdr("chec").ToString
    '        dsP(2) = rdr("tipo").ToString

    '    Else
    '        ReDim dsP(0) : dsP(0) = "Error: no se encuentra."
    '    End If
    '    rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
    '    Return dsP
    'End Function
End Class
