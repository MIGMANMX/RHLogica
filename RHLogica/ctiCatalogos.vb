Imports System.Data.SqlClient

Public Class ctiCatalogos
    '''''Puestos
    Public Function datosPuesto(ByVal idpuesto As Integer) As String()
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idpuesto, puesto FROM Puestos WHERE idpuesto = @idP", dbC)
        cmd.Parameters.AddWithValue("idP", idpuesto)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim dsP As String()
        If rdr.Read Then
            ReDim dsP(1)
            dsP(0) = rdr("puesto").ToString
        Else
            ReDim dsP(0) : dsP(0) = "Error: no se encuentra este puesto de empleado."
        End If
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dsP
    End Function
    Public Function agregarPuesto(ByVal puesto As String) As String()
        Dim ans() As String
        If puesto <> "" Then
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT idpuesto FROM Puestos WHERE puesto = @puesto", dbC)
            cmd.Parameters.AddWithValue("puesto", puesto)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                ReDim ans(0)
                ans(0) = "Error: no se puede agregar, ya existe un puesto de empleado con este nombre."
                rdr.Close()
            Else
                rdr.Close()
                cmd.CommandText = "INSERT INTO Puestos SELECT @puesto"
                cmd.ExecuteNonQuery()
                cmd.CommandText = "SELECT idpuesto FROM Puestos WHERE puesto = @puesto"
                rdr = cmd.ExecuteReader
                rdr.Read()
                ReDim ans(1)
                ans(0) = "Puesto de empleados agregado."
                ans(1) = rdr("idpuesto").ToString
                rdr.Close()
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Else
            ReDim ans(0)
            ans(0) = "Error: no se puede agregar, es necesario capturar el puesto de empleado."
        End If
        Return ans
    End Function
    Public Function actualizarPuesto(ByVal idpuesto As Integer, _
                                     ByVal puesto As String) As String
        Dim err As String
        If puesto = "" Then
            err = "Error: no se actualizó, es necesario capturar el puesto de empleado."
        Else
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT idpuesto FROM Puestos WHERE puesto = @puesto AND idpuesto <> @idP", dbC)
            cmd.Parameters.AddWithValue("puesto", puesto)
            cmd.Parameters.AddWithValue("idP", idpuesto)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                rdr.Close()
                err = "Error: no se actualizó, ya existe."
            Else
                rdr.Close()
                cmd.CommandText = "UPDATE Puestos SET puesto = @puesto WHERE idpuesto = @idP"
                cmd.ExecuteNonQuery()
                err = "Datos del puesto de empleados actualizados."
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        End If
        Return err
    End Function
    Public Function eliminarPuesto(ByVal idpuesto As Integer) As String
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idpuesto FROM Empleados WHERE idpuesto = @idP", dbC)
        cmd.Parameters.AddWithValue("idP", idpuesto)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim err As String
        If rdr.HasRows Then
            err = "Error: este puesto de empleados no se puede eliminar, tiene empleados asociadas."
            rdr.Close()
        Else
            rdr.Close()
            cmd.CommandText = "DELETE FROM Puestos WHERE idpuesto = @idP"
            cmd.ExecuteNonQuery()
            err = "Puesto de empleados eliminado."
        End If
        rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return err
    End Function
    Public Function gvPuesto() As DataTable
        Dim dt As New DataTable
        dt.Columns.Add(New DataColumn("idpuesto", System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn("puesto", System.Type.GetType("System.String")))
        Dim r As DataRow
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idpuesto, puesto FROM Puestos ORDER BY puesto", dbC)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            r = dt.NewRow
            r(0) = rdr("idpuesto").ToString : r(1) = rdr("puesto").ToString
            dt.Rows.Add(r)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dt
    End Function
    ''''''''''Usuarios
    Public Function datosUsuario(ByVal idUsuario As Integer) As String()
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT nombre, usuario, clave, nivel, idsucursal FROM Usuarios WHERE idusuario = @idU", dbC)
        cmd.Parameters.AddWithValue("idU", idUsuario)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim dsP As String()
        If rdr.Read Then
            ReDim dsP(4)
            dsP(0) = rdr("nombre").ToString
            dsP(1) = rdr("usuario").ToString
            dsP(2) = rdr("clave").ToString
            dsP(3) = rdr("nivel").ToString
            dsP(4) = rdr("idsucursal").ToString
        Else
            ReDim dsP(0) : dsP(0) = "Error: no se encuentra este usuario."
        End If
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dsP

    End Function
    Public Function datosUsuarioV(ByVal idUsuario As Integer) As String()
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT  nivel, idsucursal FROM Usuarios WHERE idusuario = @idU", dbC)
        cmd.Parameters.AddWithValue("idU", idUsuario)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim dsP As String()
        If rdr.Read Then
            ReDim dsP(1)
            dsP(0) = rdr("nivel").ToString
            dsP(1) = rdr("idsucursal").ToString
        Else
            ReDim dsP(0) : dsP(0) = "Error: no se encuentra este usuario."
        End If
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dsP

    End Function
    Public Function agregarUsuario(ByVal nombre As String, _
                                   ByVal usuario As String, _
                                   ByVal clave As String, _
                                   ByVal nivel As Integer, _
                                   ByVal idSucursal As Integer) As String()
        Dim au() As String
        If nombre <> "" And usuario <> "" And clave <> "" Then
            If nivel > 0 And nivel < 8 Then
                Dim dbC As New SqlConnection(StarTconnStr)
                dbC.Open()
                Dim cmd As New SqlCommand("SELECT sucursal FROM Sucursales WHERE idsucursal = @ids", dbC)
                cmd.Parameters.AddWithValue("ids", idSucursal)
                Dim rdr As SqlDataReader = cmd.ExecuteReader
                If rdr.HasRows Then
                    rdr.Close()
                    cmd.CommandText = "SELECT idusuario FROM Usuarios WHERE usuario = @usuario"
                    cmd.Parameters.AddWithValue("usuario", usuario)
                    rdr = cmd.ExecuteReader
                    If rdr.HasRows Then
                        ReDim au(0)
                        au(0) = "Error: no se puede agregar, ya existe este usuario."
                        rdr.Close()
                    Else
                        rdr.Close()
                        cmd.CommandText = "INSERT INTO Usuarios SELECT @nombre, @usuario, @clave, @nivel, @ids"
                        cmd.Parameters.AddWithValue("nombre", nombre)
                        cmd.Parameters.AddWithValue("clave", clave)
                        cmd.Parameters.AddWithValue("nivel", nivel)
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "SELECT idusuario FROM Usuarios WHERE usuario = @usuario"
                        rdr = cmd.ExecuteReader
                        rdr.Read()
                        ReDim au(1)
                        au(0) = "Usuario agregado."
                        au(1) = rdr("idusuario").ToString
                        rdr.Close()
                    End If
                Else
                    ReDim au(0)
                    au(0) = "Error: no se puede agregar, es necesario seleccionar la sucursal."
                    rdr.Close()
                End If
                rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
            Else
                ReDim au(0)
                au(0) = "Error: no se puede agregar, el nivel debe ser entre 1 y 7."
            End If
        Else
            ReDim au(0)
            au(0) = "Error: no se puede agregar, es necesario capturar el nombre, usuario y clave."
        End If
        Return au
    End Function
    Public Function actualizarUsuario(ByVal idUsuario As Integer, _
                                      ByVal nombre As String, _
                                      ByVal usuario As String, _
                                      ByVal clave As String, _
                                      ByVal nivel As Integer, _
                                      ByVal idSucursal As Integer) As String
        Dim aci As String
        If nombre <> "" And usuario <> "" And clave <> "" Then
            If nivel > 0 And nivel < 8 Then
                Dim dbC As New SqlConnection(StarTconnStr)
                dbC.Open()
                Dim cmd As New SqlCommand("SELECT sucursal FROM Sucursales WHERE idsucursal = @ids", dbC)
                cmd.Parameters.AddWithValue("ids", idSucursal)
                Dim rdr As SqlDataReader = cmd.ExecuteReader
                If rdr.HasRows Then
                    rdr.Close()
                    cmd.CommandText = "SELECT idusuario FROM Usuarios WHERE usuario = @usuario AND idusuario <> @idU"
                    cmd.Parameters.AddWithValue("usuario", usuario)
                    cmd.Parameters.AddWithValue("idU", idUsuario)
                    rdr = cmd.ExecuteReader
                    If rdr.HasRows Then
                        aci = "Error: no se actualizó, ya existe este usuario."
                        rdr.Close()
                    Else
                        rdr.Close()
                        cmd.CommandText = "UPDATE Usuarios SET nombre = @nombre, usuario = @usuario, clave = @clave, nivel = @nivel, idsucursal = @ids WHERE idusuario = @idU"
                        cmd.Parameters.AddWithValue("nombre", nombre)
                        cmd.Parameters.AddWithValue("clave", clave)
                        cmd.Parameters.AddWithValue("nivel", nivel)
                        cmd.ExecuteNonQuery()
                        aci = "Datos del usuario actualizados."
                    End If
                Else
                    aci = "Error: no se actualizó, es necesario seleccionar la sucursal."
                    rdr.Close()
                End If
                rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
            Else
                aci = "Error: no se actualizó, el nivel debe ser entre 1 y 7."
            End If
        Else
            aci = "Error: no se actualizó, es necesario capturar el nombre, usuario y clave."
        End If
        Return aci
    End Function
    Public Function eliminarUsuario(ByVal idUsuario As Integer) As String
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("DELETE FROM Usuarios WHERE idusuario = @idU", dbC)
        cmd.Parameters.AddWithValue("idU", idUsuario)
        cmd.ExecuteNonQuery()
        cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return "Usuario eliminado."
    End Function
    Public Function gvUsuarios(ByVal idSucursal As Integer) As DataTable
        Dim dt As New DataTable
        dt.Columns.Add(New DataColumn("idusuario", System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn("nombre", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("usuario", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("nivel", System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn("sucursal", System.Type.GetType("System.String")))
        Dim r As DataRow
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idusuario,nombre,usuario,nivel,sucursal FROM Vista_Suc_ WHERE idsucursal=@idS ORDER BY sucursal", dbC)
        cmd.Parameters.AddWithValue("idS", idSucursal)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            r = dt.NewRow
            r(0) = rdr("idusuario").ToString
            r(1) = rdr("nombre").ToString
            r(2) = rdr("usuario").ToString
            r(3) = rdr("nivel").ToString
            r(4) = rdr("sucursal").ToString
            dt.Rows.Add(r)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dt
    End Function

    'Empleados
    Public Function datosEmpleado(ByVal idEmpleado As Integer) As String()
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT empleado, idsucursal, idpuesto, activo, nss, fecha_ingreso, rfc, fecha_nacimiento, calle, numero, colonia, cp, telefono, correo, fecha_baja, idempleado FROM Empleados WHERE idempleado = @idE", dbC)
        cmd.Parameters.AddWithValue("idE", idEmpleado)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim dsP As String()
        If rdr.Read Then
            ReDim dsP(15)
            dsP(0) = rdr("empleado").ToString
            dsP(1) = rdr("idsucursal").ToString
            dsP(2) = rdr("idpuesto").ToString
            dsP(3) = rdr("activo").ToString
            dsP(4) = rdr("nss").ToString
            dsP(5) = rdr("fecha_ingreso").ToString
            dsP(6) = rdr("rfc").ToString
            dsP(7) = rdr("fecha_nacimiento").ToString
            dsP(8) = rdr("calle").ToString
            dsP(9) = rdr("numero").ToString

            dsP(10) = rdr("colonia").ToString
            dsP(11) = rdr("cp").ToString
            dsP(12) = rdr("telefono").ToString
            dsP(13) = rdr("correo").ToString
            dsP(14) = rdr("fecha_baja").ToString
            dsP(15) = rdr("idempleado").ToString
        Else
            ReDim dsP(0) : dsP(0) = "Error: no se encuentra este empleado."
        End If
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dsP
    End Function
    Public Function agregarEmpleado(ByVal nombre As String, _
                                    ByVal idSucursal As Integer, _
                                    ByVal idpuesto As String, _
                                    ByVal activo As Boolean, _
                                    ByVal nss As String, _
                                    ByVal fecha_ingreso As String, _
                                    ByVal rfc As String, _
                                    ByVal fecha_nacimiento As String, _
                                    ByVal calle As String, _
                                    ByVal numero As String, _
                                    ByVal colonia As String, _
                                    ByVal cp As Integer, _
                                    ByVal telefono As String, _
                                    ByVal correo As String, _
                                    ByVal fecha_baja As String) As String()
        Dim ae() As String
        If nombre <> "" Then
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT sucursal FROM Sucursales WHERE idsucursal = @idS", dbC)
            cmd.Parameters.AddWithValue("idS", idSucursal)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                rdr.Close()
                cmd.CommandText = "SELECT idempleado FROM Empleados WHERE empleado = @nombre"
                cmd.Parameters.AddWithValue("nombre", nombre)
                rdr = cmd.ExecuteReader
                If rdr.HasRows Then
                    ReDim ae(0)
                    ae(0) = "Error: no se puede agregar, ya existe este empleado."
                    rdr.Close()
                Else
                    rdr.Close()
                    cmd.CommandText = "INSERT INTO Empleados SELECT @nombre, @idS, @idpuesto, @activo , @nss, @fecha_ingreso, @rfc, @fecha_nacimiento, @calle, @numero, @colonia, @cp, @telefono, @correo, @fecha_baja"
                    cmd.Parameters.AddWithValue("idpuesto", idpuesto)
                    cmd.Parameters.AddWithValue("activo", activo)
                    cmd.Parameters.AddWithValue("nss", nss)
                    cmd.Parameters.AddWithValue("fecha_ingreso", fecha_ingreso)
                    cmd.Parameters.AddWithValue("rfc", rfc)
                    cmd.Parameters.AddWithValue("fecha_nacimiento", fecha_nacimiento)
                    cmd.Parameters.AddWithValue("calle", calle)
                    cmd.Parameters.AddWithValue("numero", numero)
                    cmd.Parameters.AddWithValue("colonia", colonia)
                    cmd.Parameters.AddWithValue("cp", cp)
                    cmd.Parameters.AddWithValue("telefono", telefono)
                    cmd.Parameters.AddWithValue("correo", correo)
                    cmd.Parameters.AddWithValue("fecha_baja", fecha_baja)
                    cmd.ExecuteNonQuery()
                    cmd.CommandText = "SELECT idempleado FROM Empleados WHERE empleado = @nombre"
                    rdr = cmd.ExecuteReader
                    rdr.Read()
                    ReDim ae(1)
                    ae(0) = "Empleado agregado."
                    ae(1) = rdr("idempleado").ToString
                    rdr.Close()
                End If
            Else
                ReDim ae(0)
                ae(0) = "Error: no se puede agregar, es necesario seleccionar la sucursal."
                rdr.Close()
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Else
            ReDim ae(0)
            ae(0) = "Error: no se puede agregar, es necesario capturar el nombre del empleado."
        End If
        Return ae
    End Function
    Public Function actualizarEmpleado(ByVal idEmpleado As Integer, _
                                       ByVal nombre As String, _
                                       ByVal idSucursal As Integer, _
                                       ByVal idpuesto As String, _
                                       ByVal activo As Boolean, _
                                       ByVal nss As String, _
                                        ByVal fecha_ingreso As String, _
                                        ByVal rfc As String, _
                                        ByVal fecha_nacimiento As String, _
                                        ByVal calle As String, _
                                        ByVal numero As String, _
                                        ByVal colonia As String, _
                                        ByVal cp As Integer, _
                                        ByVal telefono As String, _
                                        ByVal correo As String, _
                                        ByVal fecha_baja As String) As String
        Dim aci As String
        If nombre <> "" Then
            Dim dbC As New SqlConnection(StarTconnStr)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT sucursal FROM Sucursales WHERE idsucursal = @idS", dbC)
            cmd.Parameters.AddWithValue("idS", idSucursal)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            If rdr.HasRows Then
                rdr.Close()
                cmd.CommandText = "SELECT idempleado FROM Empleados WHERE empleado = @nombre AND idempleado <> @idE"
                cmd.Parameters.AddWithValue("nombre", nombre)
                cmd.Parameters.AddWithValue("idE", idEmpleado)
                rdr = cmd.ExecuteReader
                If rdr.HasRows Then
                    aci = "Error: no se actualizó, ya existe este empleado."
                    rdr.Close()
                Else
                    rdr.Close()
                    cmd.CommandText = "UPDATE Empleados SET empleado = @nombre, idsucursal = @idS, idpuesto = @idpuesto, activo = @activo ,nss = @nss, fecha_ingreso = @fecha_ingreso, rfc = @rfc, fecha_nacimiento =  @fecha_nacimiento, calle = @calle, numero = @numero, colonia = @colonia, cp = @cp, telefono = @telefono, correo = @correo, fecha_baja = @fecha_baja WHERE idempleado = @idE"
                    cmd.Parameters.AddWithValue("idpuesto", idpuesto)
                    cmd.Parameters.AddWithValue("activo", activo)

                    'Convert.ToDateTime()

                    cmd.Parameters.AddWithValue("nss", nss)
                    cmd.Parameters.AddWithValue("fecha_ingreso", Convert.ToDateTime(fecha_ingreso))
                    cmd.Parameters.AddWithValue("rfc", rfc)
                    cmd.Parameters.AddWithValue("fecha_nacimiento", Convert.ToDateTime(fecha_nacimiento))
                    cmd.Parameters.AddWithValue("calle", calle)
                    cmd.Parameters.AddWithValue("numero", numero)
                    cmd.Parameters.AddWithValue("colonia", colonia)
                    cmd.Parameters.AddWithValue("cp", cp)
                    cmd.Parameters.AddWithValue("telefono", telefono)
                    cmd.Parameters.AddWithValue("correo", correo)
                    cmd.Parameters.AddWithValue("fecha_baja", Convert.ToDateTime(fecha_baja))
                    cmd.ExecuteNonQuery()
                    aci = "Datos del empleado actualizados."
                End If
            Else
                aci = "Error: no se actualizó, es necesario seleccionar la sucursal."
                rdr.Close()
            End If
            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Else
            aci = "Error: no se actualizó, es necesario capturar el nombre del empleado."
        End If
        Return aci
    End Function
    Public Function eliminarEmpleado(ByVal idEmpleado As Integer) As String
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idempleado FROM Vales WHERE idempleado = @idE", dbC)
        cmd.Parameters.AddWithValue("idE", idEmpleado)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim ee As String
        If rdr.HasRows Then
            rdr.Close()
            ee = "Error: este empleado no se puede eliminar, tiene vales registrados."
        Else
            rdr.Close()
            cmd.CommandText = "DELETE FROM Empleados WHERE idempleado = @idE"
            cmd.ExecuteNonQuery()
            ee = "Empleado eliminado."
        End If
        rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return ee
    End Function
    Public Function gvEmpleados(ByVal idsucursal As Integer, ByVal activo As Boolean) As DataTable
        Dim dt As New DataTable
        dt.Columns.Add(New DataColumn("idempleado", System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn("empleado", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("puesto", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("activo", System.Type.GetType("System.Boolean")))

        dt.Columns.Add(New DataColumn("nss", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("fecha_ingreso", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("rfc", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("fecha_nacimiento", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("calle", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("numero", System.Type.GetType("System.String")))

        dt.Columns.Add(New DataColumn("colonia", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("cp", System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn("telefono", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("correo", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("fecha_baja", System.Type.GetType("System.String")))
        Dim r As DataRow
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idempleado, empleado, puesto, activo , nss, fecha_ingreso, rfc, fecha_nacimiento, calle, numero, colonia, cp, telefono, correo, fecha_baja FROM Vista_Empleados WHERE idsucursal = @idsucursal and activo = @activo ORDER BY empleado", dbC)
        cmd.Parameters.AddWithValue("idsucursal", idsucursal)
        cmd.Parameters.AddWithValue("activo", activo)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            r = dt.NewRow
            r(0) = rdr("idempleado").ToString
            r(1) = rdr("empleado").ToString
            r(2) = rdr("puesto").ToString
            r(3) = rdr("activo").ToString

            r(4) = rdr("nss").ToString
            r(5) = rdr("fecha_ingreso").ToString
            r(6) = rdr("rfc").ToString
            r(7) = rdr("fecha_nacimiento").ToString
            r(8) = rdr("calle").ToString
            r(9) = rdr("numero").ToString
            r(10) = rdr("colonia").ToString
            r(11) = rdr("cp").ToString
            r(12) = rdr("telefono").ToString
            r(12) = rdr("correo").ToString
            r(13) = rdr("fecha_baja").ToString
            dt.Rows.Add(r)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dt
    End Function

    'Empleados/Sucursales
    Public Function datosEmpleSuc(ByVal sucursal As String) As String()
        Dim dbC As New SqlConnection(StarTconnStr)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idempleado, empleado FROM Vista_Empleados WHERE sucursal=@sucursal ORDER BY sucursal", dbC)
        cmd.Parameters.AddWithValue("sucursal", sucursal)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim dsP As String()
        While rdr.Read
            ReDim dsP(4)
            dsP(0) = rdr("idempleado").ToString
            dsP(1) = rdr("empleado").ToString

        End While
        ReDim dsP(0) : dsP(0) = "Error: no se encuentra este empleado."
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dsP
    End Function

    'Partida/Jornada
    Public Function datosPartidaJornada(ByVal idpartidas_jornada As Integer) As String()
        Dim dbC As New SqlConnection(StarTconnStrRH)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idpartidas_jornada,idempleado, idjornada, fecha FROM Partidas_Jornada WHERE idpartidas_jornada = @idpartidas_jornada", dbC)
        cmd.Parameters.AddWithValue("idpartidas_jornada", idpartidas_jornada)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        Dim dsP As String()
        If rdr.Read Then
            ReDim dsP(3)
            dsP(0) = rdr("idpartidas_jornada").ToString
            dsP(1) = rdr("idempleado").ToString
            dsP(2) = rdr("idjornada").ToString
            dsP(3) = rdr("fecha").ToString

        Else
            ReDim dsP(0) : dsP(0) = "Error: no se encuentra."
        End If
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dsP
    End Function
    Public Function agregarPartidaJornada(ByVal idempleado As Integer,
                                   ByVal idjornada As Integer,
                                   ByVal fecha As String) As String()
        Dim ans() As String
        If fecha <> "" Then
            Dim dbC As New SqlConnection(StarTconnStrRH)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT idpartidas_jornada FROM Partidas_Jornada WHERE idjornada = @idjornada", dbC)
            cmd.Parameters.AddWithValue("idjornada", idjornada)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            rdr.Close()
            cmd.CommandText = "INSERT INTO Partidas_Jornada SELECT @idempleado,@idjornada,@fecha"
            cmd.Parameters.AddWithValue("fecha", Convert.ToDateTime(fecha))
            cmd.Parameters.AddWithValue("idempleado", idempleado)
            cmd.ExecuteNonQuery()
            cmd.CommandText = "SELECT idpartidas_jornada FROM Partidas_Jornada WHERE fecha = @fecha"
            rdr = cmd.ExecuteReader
            rdr.Read()
            ReDim ans(1)
            ans(0) = "Agregado."
            ans(1) = rdr("idpartidas_jornada").ToString
            rdr.Close()

            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Else
            ReDim ans(0)
            ans(0) = "Error: no se puede agregar, es necesario capturar."
        End If
        Return ans
    End Function
    Public Function eliminarPartidas_Jornada(ByVal idpartidas_jornada As Integer) As String
        Dim dbC As New SqlConnection(StarTconnStrRH)
        dbC.Open()
        Dim cmd As New SqlCommand("DELETE FROM Partidas_Jornada WHERE idpartidas_jornada = @idpartidas_jornada", dbC)
        cmd.Parameters.AddWithValue("idpartidas_jornada", idpartidas_jornada)
        cmd.ExecuteNonQuery()
        cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return "Usuario eliminado."
    End Function
    Public Function gvPartida_Jornada(ByVal idempleado As Integer) As DataTable
        Dim dt As New DataTable
        dt.Columns.Add(New DataColumn("idpartidas_jornada", System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn("idempleado", System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn("jornada", System.Type.GetType("System.String")))

        dt.Columns.Add(New DataColumn("inicio", System.Type.GetType("System.String")))
        dt.Columns.Add(New DataColumn("fin", System.Type.GetType("System.String")))

        dt.Columns.Add(New DataColumn("fecha", System.Type.GetType("System.String")))
        Dim r As DataRow
        Dim dbC As New SqlConnection(StarTconnStrRH)
        dbC.Open()
        Dim cmd As New SqlCommand("SELECT idpartidas_jornada,jornada,inicio,fin,fecha FROM vw_AHorario WHERE idempleado=@idempleado ORDER BY fecha desc", dbC)
        cmd.Parameters.AddWithValue("idempleado", idempleado)
        Dim rdr As SqlDataReader = cmd.ExecuteReader
        While rdr.Read
            r = dt.NewRow
            r(0) = rdr("idpartidas_jornada").ToString
            r(2) = rdr("jornada").ToString
            r(3) = rdr("inicio").ToString
            r(4) = rdr("fin").ToString
            r(5) = rdr("fecha").ToString
            dt.Rows.Add(r)
        End While
        rdr.Close() : rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Return dt
    End Function
    Public Function actualizarPartidaJornada(ByVal idpartidas_jornada As Integer,
                                      ByVal idempleado As Integer,
                                      ByVal idjornada As Integer,
                                      ByVal fecha As String) As String
        Dim aci As String
        aci = ""
        If Convert.ToInt32(idempleado) > 0 And Convert.ToInt32(idjornada) > 0 Then

            Dim dbC As New SqlConnection(StarTconnStrRH)
            dbC.Open()
            Dim cmd As New SqlCommand("SELECT idpartidas_jornada FROM Partidas_Jornada WHERE idpartidas_jornada = @idpartidas_jornada", dbC)
            cmd.Parameters.AddWithValue("idpartidas_jornada", idpartidas_jornada)
            Dim rdr As SqlDataReader = cmd.ExecuteReader

            rdr.Close()
            cmd.CommandText = "UPDATE Partidas_Jornada SET idempleado = @idempleado, idjornada = @idjornada, fecha = @fecha WHERE idpartidas_jornada = @idpartidas_jornada"
            cmd.Parameters.AddWithValue("idempleado", idempleado)
            cmd.Parameters.AddWithValue("idjornada", idjornada)
            cmd.Parameters.AddWithValue("fecha", Convert.ToDateTime(fecha))
            cmd.ExecuteNonQuery()
            aci = "Datos actualizados."

            rdr = Nothing : cmd.Dispose() : dbC.Close() : dbC.Dispose()
        Else
            aci = "Error: no se actualizó, es necesario capturar"
        End If
        Return aci
    End Function
End Class

