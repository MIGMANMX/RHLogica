Imports System.Data.SqlClient
Public Class ctiCalculo
    'Calculo de Horas
    '''Gridview Chequeo
    Public Function gvChequeo(ByVal idempleado As Integer, ByVal Fech1 As String, ByVal Fech2 As String) As DataTable
        Dim dt As New DataTable
        dt.Columns.Add(New DataColumn("chec", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("tipo", System.Type.GetType("System.Int32")))

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
End Class
