Imports System.Data.SqlClient
Imports System.Security.Cryptography
Imports System.Text

Public Class Form1
    Private cmd As SqlCommand
    Private IDuser As Integer = 0, tipo_User As Integer = 0
    Private con As String = ""
    Private modCategoria As Integer = 0
    Private modEditorial As String = ""
    Private modCliente As Integer = 0
    Private modPago As Integer = 0
    Private modDepartamento As Integer = 0
    Private modPuesto As Integer = 0
    Private modFormato As Integer = 0
    Private modLibro As String = ""
    Private modAutor As String = ""
    Private modUsuario As Integer = 0
    Private modcon As String = ""
    Private modVenta As Integer = 0
    Private Function Encriptado(ByVal co As String) As String
        Using sha256 = New SHA256Managed()
            Return BitConverter.ToString(sha256.ComputeHash(Encoding.UTF8.GetBytes(co))).Replace("-", "")
        End Using
    End Function
    Private Sub ventasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ventasToolStripMenuItem.Click
        cerrarPaneles()
        PanelVentas.Visible = True
        PanelVentas.Enabled = True
        recargaCatalogoVentas()
    End Sub

    Private Sub cajaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles cajaToolStripMenuItem.Click
        cerrarPaneles()
        panelVentaRegistro.Visible = True
        panelVentaRegistro.Enabled = True
        cargaCombosVentaLibros()
    End Sub

    Private Sub HistorialToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HistorialToolStripMenuItem.Click
        cerrarPaneles()
        PanelHistorial.Visible = True
        PanelHistorial.Enabled = True
        recargaHistorial()
    End Sub

    Private Sub UsuariosToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles UsuariosToolStripMenuItem1.Click
        cerrarPaneles()
        PanelUsuarios.Visible = True
        PanelUsuarios.Enabled = True
        recargaCatalogoUsuario()
    End Sub

    Private Sub MiCuentaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MiCuentaToolStripMenuItem.Click
        Dim id As Integer = 0
        cerrarPaneles()

        Try

            Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                sql.Open()
                Dim consulta As String = "select Nombre_Usuario,Nombre,Contraseña,Apellido_Paterno,Apellido_Materno,Email,Telefono,Puesto,Departamento from VRegistroUsuarios where Id_Usuario=" & IDuser.ToString()
                cmd = New SqlCommand(consulta, sql)
                Dim lector As SqlDataReader
                lector = cmd.ExecuteReader()

                If lector.HasRows Then

                    While lector.Read()
                        txtUsuarioMiCuenta.Text = CStr(lector(0))
                        txtNombreMiCuenta.Text = CStr(lector(1))
                        txtContraMiCuenta.Text = CStr(lector(2))
                        txtApellidoPMiCuenta.Text = CStr(lector(3))
                        txtApellidoMMiCuenta.Text = CStr(lector(4))
                        txtCorreoMiCuenta.Text = CStr(lector(5))
                        txtTelefMiCuenta.Text = CStr(lector(6))
                        txtPuestoMiCuenta.Text = CStr(lector(7))
                        txtDepartMiCuenta.Text = CStr(lector(8))
                        con = CStr(lector(2))
                    End While

                    PanelRegistra.Visible = True
                    PanelRegistra.Enabled = True
                Else
                    PanelRegistra.Visible = False
                    PanelRegistra.Enabled = False
                    MessageBox.Show("El usuario no existe o datos incorrectos")
                End If

                sql.Close()
            End Using

        Catch ex As Exception
            MessageBox.Show("Error en conexion")
        End Try
    End Sub

    Private Sub cerrarSesiónToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles cerrarSesiónToolStripMenuItem.Click
        cerrarPaneles()
        Menu.Visible = False
        Menu.Enabled = False
        MessageBox.Show("Sesión cerrada")
        PanelLogin.Enabled = True
        PanelLogin.Visible = True
        txtContraLogin.Clear()
        txtCorreoLogin.Clear()
        IDuser = 0
        tipo_User = 0
    End Sub

    Private Sub AutoresToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AutoresToolStripMenuItem.Click
        cerrarPaneles()
        PanelAutor.Visible = True
        PanelAutor.Enabled = True
        recargaCatalogoAutor()
    End Sub

    Private Sub LibrosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LibrosToolStripMenuItem.Click
        cerrarPaneles()
        panelLibrosFormatos.Visible = True
        panelLibrosFormatos.Enabled = True
        recargaCatalogoTipoFormato()
        recargaCatalogoLibro()
    End Sub

    Private Sub CategoriasYEditorialesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CategoriasYEditorialesToolStripMenuItem.Click
        cerrarPaneles()
        PanelCateEditor.Visible = True
        PanelCateEditor.Enabled = True
        recargaCatalogoCategoria()
        recargaCatalogoEditorial()
    End Sub

    Private Sub DepartamentosYPuestosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DepartamentosYPuestosToolStripMenuItem.Click
        cerrarPaneles()
        PanelDepartamentosPuestos.Visible = True
        PanelDepartamentosPuestos.Enabled = True
        recargaCatalogoPuestos()
        recargaCatalogoDepartamento()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cerrarPaneles()
        Menu.Visible = False
        Menu.Enabled = False
        PanelLogin.Enabled = True
        PanelLogin.Visible = True
    End Sub
    Private Sub cerrarPaneles()
        PanelRegistra.Enabled = False
        PanelRegistra.Visible = False
        PanelCateEditor.Visible = False
        PanelCateEditor.Enabled = False
        gbCategoria.Enabled = False
        gbCategoria.Visible = False
        gbEditorial.Visible = False
        gbEditorial.Enabled = False
        pTiClienteTiPago.Enabled = False
        pTiClienteTiPago.Visible = False
        gbClientes.Visible = False
        gbClientes.Enabled = False
        gpPago.Visible = False
        gpPago.Enabled = False
        PanelDepartamentosPuestos.Visible = False
        PanelDepartamentosPuestos.Enabled = False
        gbPuestos.Enabled = False
        gbPuestos.Visible = False
        gbDepartamento.Enabled = False
        gbDepartamento.Visible = False
        panelLibrosFormatos.Enabled = False
        panelLibrosFormatos.Visible = False
        gbFormato.Enabled = False
        gbFormato.Visible = False
        panelLibros.Enabled = False
        panelLibros.Visible = False
        PanelAutor.Enabled = False
        PanelAutor.Visible = False
        panelUsuarioAM.Enabled = False
        panelUsuarioAM.Visible = False
        panelVentaRegistro.Enabled = False
        panelVentaRegistro.Visible = False
        PanelDetalleVenta.Enabled = False
        PanelDetalleVenta.Visible = False
        PanelVentas.Enabled = False
        PanelVentas.Visible = False
        PanelUsuarios.Enabled = False
        PanelUsuarios.Visible = False
        PanelHistorial.Enabled = False
        PanelHistorial.Visible = False
        PanelAutorAM.Enabled = False
        PanelAutorAM.Visible = False
    End Sub
    Private Sub LogsAceptable(ByVal accion As String, ByVal tabla As String)
        Try

            Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                sql.Open()
                cmd = New SqlCommand("CreateLogs", sql)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add(New SqlParameter("@Id_Usuario", IDuser))
                cmd.Parameters.Add(New SqlParameter("@Fecha_Hora", DateTime.Now))
                cmd.Parameters.Add(New SqlParameter("@Accion", accion))
                cmd.Parameters.Add(New SqlParameter("@Tabla", tabla))
                cmd.Parameters.Add(New SqlParameter("@Estatus", 1))
                cmd.Parameters.Add(New SqlParameter("@Id_Nivel", 1))
                cmd.ExecuteReader()
            End Using

        Catch ex As Exception
        End Try
    End Sub
    Private Sub btnGuardarReg_Click(sender As Object, e As EventArgs) Handles btnGuardarReg.Click
        If Not (String.IsNullOrEmpty(txtContraMiCuenta.Text) OrElse String.IsNullOrWhiteSpace(txtContraMiCuenta.Text)) AndAlso Not (txtContraMiCuenta.Text.Equals(con)) Then
            con = Encriptado(txtContraMiCuenta.Text)
        End If

        Try

            Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                sql.Open()
                Dim consulta As String = "update Usuario set Nombre_Usuario='" & txtUsuarioMiCuenta.Text & "',Nombre='" + txtNombreMiCuenta.Text & "'," & "Contraseña='" & con & "',Apellido_Paterno='" + txtApellidoPMiCuenta.Text & "',Apellido_Materno='" + txtApellidoMMiCuenta.Text & "',Email='" + txtCorreoMiCuenta.Text & "',Telefono='" + txtTelefMiCuenta.Text & "',Id_Usuario_Modificacion=" & IDuser.ToString() & ",Fecha_Hora_Modificacion= '" & DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff") & "' where Id_Usuario=" & IDuser.ToString()
                cmd = New SqlCommand(consulta, sql)
                cmd.ExecuteReader()
                PanelRegistra.Visible = True
                PanelRegistra.Enabled = True
                LogsAceptable("Modificar mi perfil", "Usuario")
                sql.Close()
                MessageBox.Show("Guardado")
            End Using

        Catch ex As Exception
            MessageBox.Show("Error en conexion")
        End Try
    End Sub

    Private Sub txtContraMiCuenta_Enter(sender As Object, e As EventArgs) Handles txtContraMiCuenta.Enter
        txtContraMiCuenta.Text = ""
    End Sub
    Private Sub recargaCatalogoCategoria()
        Dim dt As DataTable = New DataTable()

        Try

            Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                sql.Open()

                Using cmd As SqlCommand = New SqlCommand()
                    cmd.Connection = sql
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = "select Id_Categoria,Nombre,Estatus from Categoria"

                    Using da As SqlDataAdapter = New SqlDataAdapter(cmd)
                        da.Fill(dt)
                    End Using
                End Using

                sql.Close()
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        dgvCategoria.DataSource = dt
        dgvCategoria.Columns("Id_Categoria").Visible = False
        dgvCategoria.Columns("Estatus").Visible = False
    End Sub

    Private Sub recargaCatalogoEditorial()
        Dim dt As DataTable = New DataTable()

        Try

            Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                sql.Open()

                Using cmd As SqlCommand = New SqlCommand()
                    cmd.Connection = sql
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = "SELECT e.Nombre AS Editorial, e.Año_Inauguracion, e.Id_Ciudad, c.Nombre AS Ciudad,  es.Nombre AS Estado, p.Nombre AS Pais, c.Id_Estado, es.Id_Pais, e.Id_Editorial, e.Estatus FROM Editorial e INNER JOIN Ciudad c ON e.Id_Ciudad = c.Id_Ciudad INNER JOIN Estado es ON c.Id_Estado = es.Id_Estado INNER JOIN Pais p ON es.Id_Pais = p.Id_Pais"

                    Using da As SqlDataAdapter = New SqlDataAdapter(cmd)
                        da.Fill(dt)
                    End Using
                End Using

                sql.Close()
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        dgvEditorial.DataSource = dt
        dgvEditorial.Columns("Id_Ciudad").Visible = False
        dgvEditorial.Columns("Id_Estado").Visible = False
        dgvEditorial.Columns("Id_Pais").Visible = False
        dgvEditorial.Columns("Id_Editorial").Visible = False
        dgvEditorial.Columns("Estatus").Visible = False
    End Sub

    Private Sub recargaCatalogoTipoCliente()
        Dim dt As DataTable = New DataTable()

        Try

            Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                sql.Open()

                Using cmd As SqlCommand = New SqlCommand()
                    cmd.Connection = sql
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = "  select Id_Tipo_Cliente,Nombre,Estatus from Tipo_Cliente"

                    Using da As SqlDataAdapter = New SqlDataAdapter(cmd)
                        da.Fill(dt)
                    End Using
                End Using

                sql.Close()
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        dgCliente.DataSource = dt
        dgCliente.Columns("Nombre").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        dgCliente.Columns("Id_Tipo_Cliente").Visible = False
        dgCliente.Columns("Estatus").Visible = False
    End Sub

    Private Sub recargaHistorial()
        Dim dt As DataTable = New DataTable()

        Try

            Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                sql.Open()

                Using cmd As SqlCommand = New SqlCommand()
                    cmd.Connection = sql
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = "select l.Fecha_Hora,u.Nombre_Usuario,concat(u.Nombre, ' ', u.Apellido_Paterno, ' ', u.Apellido_Materno) Nombre_Completo,l.Accion, l.Tabla,n.Valor from Logs l left join Usuario u on l.Id_Usuario = u.Id_Usuario left join Nivel n on l.Id_Nivel = n.Id_Nivel order by l.Fecha_Hora"

                    Using da As SqlDataAdapter = New SqlDataAdapter(cmd)
                        da.Fill(dt)
                    End Using
                End Using

                sql.Close()
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        dgvHistorial.DataSource = dt
    End Sub

    Private Sub recargaCatalogoTipoPago()
        Dim dt As DataTable = New DataTable()

        Try

            Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                sql.Open()

                Using cmd As SqlCommand = New SqlCommand()
                    cmd.Connection = sql
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = "select Id_Tipo_Pago,Nombre,Estatus from Tipo_Pago"

                    Using da As SqlDataAdapter = New SqlDataAdapter(cmd)
                        da.Fill(dt)
                    End Using
                End Using

                sql.Close()
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        dgvPago.DataSource = dt
        dgvPago.Columns("Nombre").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        dgvPago.Columns("Id_Tipo_Pago").Visible = False
        dgvPago.Columns("Estatus").Visible = False
    End Sub

    Private Sub recargaCatalogoTipoFormato()
        Dim dt As DataTable = New DataTable()

        Try

            Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                sql.Open()

                Using cmd As SqlCommand = New SqlCommand()
                    cmd.Connection = sql
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = "select Id_Tipo_Formato,Nombre,Estatus from Tipo_Formato"

                    Using da As SqlDataAdapter = New SqlDataAdapter(cmd)
                        da.Fill(dt)
                    End Using
                End Using

                sql.Close()
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        dgvFormatos.DataSource = dt
        dgvFormatos.Columns("Nombre").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        dgvFormatos.Columns("Id_Tipo_Formato").Visible = False
        dgvFormatos.Columns("Estatus").Visible = False
    End Sub

    Private Sub recargaCatalogoAutor()
        Dim dt As DataTable = New DataTable()

        Try

            Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                sql.Open()

                Using cmd As SqlCommand = New SqlCommand()
                    cmd.Connection = sql
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = "select a.Id_Autor,a.Nombre,a.Apellido,CONCAT(a.Nombre,' ',a.Apellido) Nombre_Completo,a.Fecha_Nacimiento,a.Id_Pais,p.Nacionalidad,a.Estatus from Autor a left join Pais p on a.Id_Pais=p.Id_Pais"

                    Using da As SqlDataAdapter = New SqlDataAdapter(cmd)
                        da.Fill(dt)
                    End Using
                End Using

                sql.Close()
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        dgvAutor.DataSource = dt
        dgvAutor.Columns("Id_Autor").Visible = False
        dgvAutor.Columns("Nombre").Visible = False
        dgvAutor.Columns("Apellido").Visible = False
        dgvAutor.Columns("Id_Pais").Visible = False
        dgvAutor.Columns("Estatus").Visible = False
    End Sub

    Private Sub recargaCatalogoLibro()
        Dim dt As DataTable = New DataTable()

        Try

            Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                sql.Open()

                Using cmd As SqlCommand = New SqlCommand()
                    cmd.Connection = sql
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = "select l.Id_Libro_ISBN,l.Titulo,l.Id_Categoria,c.Nombre Categoria,isnull(l.Volumen,'') Volumen,l.Id_Tipo_Formato,f.Nombre Formato,l.Id_Editorial, e.Nombre Editorial,isnull(l.Fecha_Publicacion,'') [Fecha de publicacion],isnull(l.Paginas,'') Paginas,l.Precio,l.Id_Rating,r.Valor Rating,l.Estatus from Libro l left join Categoria c on l.Id_Categoria=c.Id_Categoria left join Tipo_Formato f on l.Id_Tipo_Formato=f.Id_Tipo_Formato left join Editorial e on l.Id_Editorial= e.Id_Editorial left join Rating r on l.Id_Rating=r.Id_Rating" & vbCrLf

                    Using da As SqlDataAdapter = New SqlDataAdapter(cmd)
                        da.Fill(dt)
                    End Using
                End Using

                sql.Close()
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        dgvLibros.DataSource = dt
        dgvLibros.Columns("Id_Libro_ISBN").Visible = False
        dgvLibros.Columns("Id_Categoria").Visible = False
        dgvLibros.Columns("Id_Tipo_Formato").Visible = False
        dgvLibros.Columns("Id_Editorial").Visible = False
        dgvLibros.Columns("Id_Rating").Visible = False
        dgvLibros.Columns("Estatus").Visible = False
    End Sub

    Private Sub recargaCatalogoDepartamento()
        Dim dt As DataTable = New DataTable()

        Try

            Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                sql.Open()

                Using cmd As SqlCommand = New SqlCommand()
                    cmd.Connection = sql
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = "select Id_Departamento,Departamento,Estatus from Departamento"

                    Using da As SqlDataAdapter = New SqlDataAdapter(cmd)
                        da.Fill(dt)
                    End Using
                End Using

                sql.Close()
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        dgvDepartamentos.DataSource = dt
        dgvDepartamentos.Columns("Departamento").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        dgvDepartamentos.Columns("Id_Departamento").Visible = False
        dgvDepartamentos.Columns("Estatus").Visible = False
    End Sub

    Private Sub recargaCatalogoUsuario()
        Dim dt As DataTable = New DataTable()

        Try

            Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                sql.Open()

                Using cmd As SqlCommand = New SqlCommand()
                    cmd.Connection = sql
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = "select u.Id_Usuario,u.Nombre_Usuario,u.Nombre,u.Contraseña,u.Apellido_Paterno,u.Apellido_Materno,u.Email,u.Telefono,u.Id_Puesto,p.Nombre Puesto,u.Id_Departamento,d.Departamento,u.Id_Tipo_Usuario,t.Nombre [Tipo de Usuario],u.Estatus from Usuario u left join Puesto p on u.Id_Puesto=p.Id_Puesto left join Departamento d on u.Id_Departamento=d.Id_Departamento left join Tipo_Usuario t on u.Id_Tipo_Usuario=t.Id_Tipo_Usuario" & vbCrLf

                    Using da As SqlDataAdapter = New SqlDataAdapter(cmd)
                        da.Fill(dt)
                    End Using
                End Using

                sql.Close()
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        dgvUsuarios.DataSource = dt
        dgvUsuarios.Columns("Id_Usuario").Visible = False
        dgvUsuarios.Columns("Contraseña").Visible = False
        dgvUsuarios.Columns("Id_Puesto").Visible = False
        dgvUsuarios.Columns("Id_Departamento").Visible = False
        dgvUsuarios.Columns("Id_Tipo_Usuario").Visible = False
        dgvUsuarios.Columns("Estatus").Visible = False
    End Sub

    Private Sub recargaCatalogoPuestos()
        Dim dt As DataTable = New DataTable()

        Try

            Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                sql.Open()

                Using cmd As SqlCommand = New SqlCommand()
                    cmd.Connection = sql
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = "select Id_Puesto,Nombre,Estatus from Puesto"

                    Using da As SqlDataAdapter = New SqlDataAdapter(cmd)
                        da.Fill(dt)
                    End Using
                End Using

                sql.Close()
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        dgvPuestos.DataSource = dt
        dgvPuestos.Columns("Nombre").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        dgvPuestos.Columns("Id_Puesto").Visible = False
        dgvPuestos.Columns("Estatus").Visible = False
    End Sub

    Private Sub cargacombosUsuarios()
        Dim t As DataTable = New DataTable()
        Dim de As DataTable = New DataTable()
        Dim p As DataTable = New DataTable()

        Try

            Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                sql.Open()

                Using cmd As SqlCommand = New SqlCommand()
                    cmd.Connection = sql
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = "select Id_Tipo_Usuario, Nombre from Tipo_Usuario"

                    Using da As SqlDataAdapter = New SqlDataAdapter(cmd)
                        da.Fill(t)
                    End Using
                End Using

                sql.Close()
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        cbTipoUsuarioUAM.DataSource = t
        cbTipoUsuarioUAM.DisplayMember = "Nombre"
        cbTipoUsuarioUAM.ValueMember = "Id_Tipo_Usuario"

        Try

            Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                sql.Open()

                Using cmd As SqlCommand = New SqlCommand()
                    cmd.Connection = sql
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = "select Id_Departamento,Departamento from Departamento where Estatus=1"

                    Using da As SqlDataAdapter = New SqlDataAdapter(cmd)
                        da.Fill(de)
                    End Using
                End Using

                sql.Close()
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        cbDepartamentoUAM.DataSource = de
        cbDepartamentoUAM.DisplayMember = "Departamento"
        cbDepartamentoUAM.ValueMember = "Id_Departamento"

        Try

            Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                sql.Open()

                Using cmd As SqlCommand = New SqlCommand()
                    cmd.Connection = sql
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = "select Id_Puesto,Nombre from Puesto where Estatus=1"

                    Using da As SqlDataAdapter = New SqlDataAdapter(cmd)
                        da.Fill(p)
                    End Using
                End Using

                sql.Close()
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        cbPuestoUAM.DataSource = p
        cbPuestoUAM.DisplayMember = "Nombre"
        cbPuestoUAM.ValueMember = "Id_Puesto"
    End Sub

    Private Sub cargaCombosVentaLibros()
        Dim l As DataTable = New DataTable()
        Dim c As DataTable = New DataTable()
        Dim p As DataTable = New DataTable()

        Try

            Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                sql.Open()

                Using cmd As SqlCommand = New SqlCommand()
                    cmd.Connection = sql
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = "select l.Id_Libro_ISBN,CONCAT(l.Titulo,', ',f.Nombre) Nombre from Libro l left join Tipo_Formato f on l.Id_Tipo_Formato=f.Id_Tipo_Formato where l.Estatus=1"

                    Using da As SqlDataAdapter = New SqlDataAdapter(cmd)
                        da.Fill(l)
                    End Using
                End Using

                sql.Close()
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        cbLibrosCaja.DataSource = l
        cbLibrosCaja.DisplayMember = "Nombre"
        cbLibrosCaja.ValueMember = "Id_Libro_ISBN"

        Try

            Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                sql.Open()

                Using cmd As SqlCommand = New SqlCommand()
                    cmd.Connection = sql
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = "select Id_Tipo_Cliente,Nombre from Tipo_Cliente where Estatus=1" & vbCrLf

                    Using da As SqlDataAdapter = New SqlDataAdapter(cmd)
                        da.Fill(c)
                    End Using
                End Using

                sql.Close()
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        cbClienteCaja.DataSource = c
        cbClienteCaja.DisplayMember = "Nombre"
        cbClienteCaja.ValueMember = "Id_Tipo_Cliente"

        Try

            Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                sql.Open()

                Using cmd As SqlCommand = New SqlCommand()
                    cmd.Connection = sql
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = "select Id_Tipo_Pago,Nombre from Tipo_Pago where Estatus=1"

                    Using da As SqlDataAdapter = New SqlDataAdapter(cmd)
                        da.Fill(p)
                    End Using
                End Using

                sql.Close()
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        cbPagoCaja.DataSource = p
        cbPagoCaja.DisplayMember = "Nombre"
        cbPagoCaja.ValueMember = "Id_Tipo_Pago"
    End Sub
    Private Sub cargaDetalleVenta(ByVal venta As Integer)
        Dim dt As DataTable = New DataTable()

        Try

            Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                sql.Open()

                Using cmd As SqlCommand = New SqlCommand()
                    cmd.Connection = sql
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = "select l.Titulo,(de.Porcentaje*100) Descuento,l.Precio from Detalle_Venta d left join Libro l on d.Id_Libro_ISBN=l.Id_Libro_ISBN left join Descuento de on d.Id_Descuento= de.id_Descuento where d.Id_Venta=" & venta.ToString()

                    Using da As SqlDataAdapter = New SqlDataAdapter(cmd)
                        da.Fill(dt)
                    End Using
                End Using

                sql.Close()
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        dgvDetalleVenta.DataSource = dt
    End Sub

    Private Sub recargaCatalogoVentas()
        Dim dt As DataTable = New DataTable()

        Try

            Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                sql.Open()

                Using cmd As SqlCommand = New SqlCommand()
                    cmd.Connection = sql
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = "select v.Id_Venta Venta,v.Fecha_Hora_Creacion Fecha, u.Nombre_Usuario,c.Nombre Cliente,p.Nombre [Forma de Pago] from venta v left join Tipo_Cliente c on v.Id_Tipo_Cliente=c.Id_Tipo_Cliente left join Usuario u on v.Id_Usuario_Creacion=u.Id_Usuario left join Tipo_Pago p on v.Id_Tipo_Pago=p.Id_Tipo_Pago where v.Estatus=1"

                    Using da As SqlDataAdapter = New SqlDataAdapter(cmd)
                        da.Fill(dt)
                    End Using
                End Using

                sql.Close()
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        dgvVerVentas.DataSource = dt
    End Sub
    Private Sub cargacombosLibros()
        Dim cat As DataTable = New DataTable()
        Dim Ed As DataTable = New DataTable()
        Dim ra As DataTable = New DataTable()
        Dim fo As DataTable = New DataTable()
        Dim au As DataTable = New DataTable()

        Try

            Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                sql.Open()

                Using cmd As SqlCommand = New SqlCommand()
                    cmd.Connection = sql
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = "select Id_Categoria,Nombre from Categoria where Estatus=1"

                    Using da As SqlDataAdapter = New SqlDataAdapter(cmd)
                        da.Fill(cat)
                    End Using
                End Using

                sql.Close()
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        cbCategoriaLibro.DataSource = cat
        cbCategoriaLibro.DisplayMember = "Nombre"
        cbCategoriaLibro.ValueMember = "Id_Categoria"

        Try

            Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                sql.Open()

                Using cmd As SqlCommand = New SqlCommand()
                    cmd.Connection = sql
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = "select Id_Editorial,Nombre from Editorial where Estatus=1"

                    Using da As SqlDataAdapter = New SqlDataAdapter(cmd)
                        da.Fill(Ed)
                    End Using
                End Using

                sql.Close()
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        cbEditorialLibro.DataSource = Ed
        cbEditorialLibro.DisplayMember = "Nombre"
        cbEditorialLibro.ValueMember = "Id_Editorial"

        Try

            Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                sql.Open()

                Using cmd As SqlCommand = New SqlCommand()
                    cmd.Connection = sql
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = "select Id_Rating,Valor from Rating where Estatus=1"

                    Using da As SqlDataAdapter = New SqlDataAdapter(cmd)
                        da.Fill(ra)
                    End Using
                End Using

                sql.Close()
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        cbRatingLibro.DataSource = ra
        cbRatingLibro.DisplayMember = "Valor"
        cbRatingLibro.ValueMember = "Id_Rating"

        Try

            Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                sql.Open()

                Using cmd As SqlCommand = New SqlCommand()
                    cmd.Connection = sql
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = "select Id_Tipo_Formato,Nombre from Tipo_Formato where Estatus=1"

                    Using da As SqlDataAdapter = New SqlDataAdapter(cmd)
                        da.Fill(fo)
                    End Using
                End Using

                sql.Close()
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        cbFormatoLibro.DataSource = fo
        cbFormatoLibro.DisplayMember = "Nombre"
        cbFormatoLibro.ValueMember = "Id_Tipo_Formato"

        Try

            Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                sql.Open()

                Using cmd As SqlCommand = New SqlCommand()
                    cmd.Connection = sql
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = "select a.Id_Autor,CONCAT(a.Nombre,' ',a.Apellido) Nombre_Completo from Autor a where a.Estatus=1" & vbCrLf

                    Using da As SqlDataAdapter = New SqlDataAdapter(cmd)
                        da.Fill(au)
                    End Using
                End Using

                sql.Close()
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        cbAutoresLibro.DataSource = au
        cbAutoresLibro.DisplayMember = "Nombre_Completo"
        cbAutoresLibro.ValueMember = "Id_Autor"
    End Sub
    Private Sub TiposDeClientesYTiposDePagosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TiposDeClientesYTiposDePagosToolStripMenuItem.Click
        cerrarPaneles()
        pTiClienteTiPago.Visible = True
        pTiClienteTiPago.Enabled = True
        recargaCatalogoTipoPago()
        recargaCatalogoTipoCliente()
    End Sub

    Private Sub btnAcceder_Click(sender As Object, e As EventArgs) Handles btnAcceder.Click
        If String.IsNullOrEmpty(txtCorreoLogin.Text) Then
            txtCorreoLogin.Select()
            MessageBox.Show("Favor de llenar todos los campos solicitados")
        ElseIf String.IsNullOrEmpty(txtContraLogin.Text) Then
            txtContraLogin.Select()
            MessageBox.Show("Favor de llenar todos los campos solicitados")
        Else

            Try

                Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                    sql.Open()
                    Dim consulta As String = "Select Id_Usuario,Id_Tipo_Usuario from Usuario where Email='" & txtCorreoLogin.Text & "'and Contraseña='" & Encriptado(txtContraLogin.Text) & "'"
                    cmd = New SqlCommand(consulta, sql)
                    Dim lector As SqlDataReader
                    lector = cmd.ExecuteReader()

                    If lector.HasRows Then

                        While lector.Read()
                            IDuser = CInt(lector(0))
                            tipo_User = CInt(lector(1))
                            MessageBox.Show("Se inicio la sesion")
                        End While

                        If tipo_User = 1 Then
                            Menu.Enabled = True
                            Menu.Visible = True
                        Else
                            UsuariosToolStripMenuItem.Visible = False
                            UsuariosToolStripMenuItem.Enabled = False
                            CatalogosToolStripMenuItem.Visible = False
                            CatalogosToolStripMenuItem.Visible = False
                            Menu.Enabled = True
                            Menu.Visible = True
                        End If

                        txtContraLogin.Clear()
                        PanelLogin.Enabled = False
                        PanelLogin.Visible = False
                        LogsAceptable("Iniciar Sesion", "Logs")
                    Else
                        txtContraLogin.Clear()
                        txtCorreoLogin.Clear()
                        MessageBox.Show("El usuario no existe o datos incorrectos")
                    End If
                End Using

            Catch ex As Exception
                MessageBox.Show("Error en conexion")
            End Try
        End If
    End Sub

    Private Sub btnAgregCate_Click(sender As Object, e As EventArgs) Handles btnAgregCate.Click
        gbCategoria.Visible = True
        gbCategoria.Enabled = True
        gbCategoria.BringToFront()
        txtUpCreNombreCat.Text = ""
        ckbCat.Checked = True
        modCategoria = 0
    End Sub

    Private Sub btnUpdCat_Click(sender As Object, e As EventArgs) Handles btnUpdCat.Click
        If dgvCategoria.SelectedRows.Count = 1 Then
            modCategoria = CInt(dgvCategoria.SelectedCells(0).Value)
            txtUpCreNombreCat.Text = CStr(dgvCategoria.SelectedCells(1).Value)

            If CBool(dgvCategoria.SelectedCells(2).Value) Then
                ckbCat.Checked = True
            Else
                ckbCat.Checked = False
            End If

            gbCategoria.Visible = True
            gbCategoria.Enabled = True
            gbCategoria.BringToFront()
        Else
            MessageBox.Show("Seleccione una categoria.")
        End If
    End Sub

    Private Sub btnElimCate_Click(sender As Object, e As EventArgs) Handles btnElimCate.Click
        If dgvCategoria.SelectedRows.Count = 1 Then

            Try

                Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                    sql.Open()
                    cmd = New SqlCommand("DeleteCategoria", sql)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.Parameters.Add(New SqlParameter("@Id", dgvCategoria.SelectedCells(0).Value))
                    cmd.ExecuteReader()
                    LogsAceptable("Eliminar Categoria", "Categoria")
                    sql.Close()
                    MessageBox.Show("Eliminada")
                    recargaCatalogoCategoria()
                End Using

            Catch ex As Exception
                MessageBox.Show("Error de conexion")
            End Try
        Else
            MessageBox.Show("Seleccione una categoria.")
        End If
    End Sub

    Private Sub btnUpdEdit_Click(sender As Object, e As EventArgs) Handles btnUpdEdit.Click
        If dgvEditorial.SelectedRows.Count = 1 Then
            Dim dt As DataTable = New DataTable()

            Try

                Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                    sql.Open()

                    Using cmd As SqlCommand = New SqlCommand()
                        cmd.Connection = sql
                        cmd.CommandType = CommandType.Text
                        cmd.CommandText = "SELECT c.Id_Ciudad, concat(c.Nombre,', ',es.Nombre,', ',p.Nombre) as localizacion FROM Ciudad c  INNER JOIN Estado es ON c.Id_Estado = es.Id_Estado INNER JOIN Pais p ON es.Id_Pais = p.Id_Pais"

                        Using da As SqlDataAdapter = New SqlDataAdapter(cmd)
                            da.Fill(dt)
                        End Using
                    End Using

                    sql.Close()
                End Using

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

            Dim año As String = "2023"

            If Not String.IsNullOrEmpty(CStr(dgvEditorial.SelectedCells(1).Value)) Then
                año = CStr(dgvEditorial.SelectedCells(1).Value)
            End If

            cbCiudadEditorial.DataSource = dt
            cbCiudadEditorial.DisplayMember = "localizacion"
            cbCiudadEditorial.ValueMember = "Id_Ciudad"
            cbCiudadEditorial.SelectedValue = CInt(dgvEditorial.SelectedCells(2).Value)
            modEditorial = CStr(dgvEditorial.SelectedCells(8).Value)
            txtNombreEditorial.Text = CStr(dgvEditorial.SelectedCells(0).Value)
            dtañoinaguracionEditorial.Value = New DateTime(Convert.ToInt32(año), 1, 1)

            If CBool(dgvEditorial.SelectedCells(9).Value) Then
                chkEditorial.Checked = True
            Else
                chkEditorial.Checked = False
            End If

            gbEditorial.Visible = True
            gbEditorial.Enabled = True
            gbEditorial.BringToFront()
        Else
            MessageBox.Show("Seleccione una Editorial.")
        End If
    End Sub
    Private Function SiglaEditExis(ByVal nombre As String) As String
        Dim siglas As String = ""

        For i As Integer = 0 To nombre.Length - 1 - 1

            If i = 0 Then
                siglas = siglas & nombre.Substring(i, 1)
            End If

            If nombre.Substring(i, 1).Equals(" ") Then
                siglas = siglas & nombre.Substring(i + 1, 1)
            End If
        Next

        If siglas.Length > 3 Then
            siglas = siglas.Substring(0, 3)
        End If

        Return siglas.ToUpper()
    End Function

    Private Sub btnElimEdit_Click(sender As Object, e As EventArgs) Handles btnElimEdit.Click
        If dgvEditorial.SelectedRows.Count = 1 Then

            Try

                Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                    sql.Open()
                    cmd = New SqlCommand("DeleteEditorial", sql)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.Parameters.Add(New SqlParameter("@Id", dgvEditorial.SelectedCells(8).Value))
                    cmd.ExecuteReader()
                    LogsAceptable("Eliminar Editorial", "Editorial")
                    sql.Close()
                    MessageBox.Show("Eliminada")
                    recargaCatalogoEditorial()
                End Using

            Catch ex As Exception
                MessageBox.Show("Error de conexion")
            End Try
        Else
            MessageBox.Show("Seleccione una Editorial.")
        End If
    End Sub

    Private Sub btnAgregaEdit_Click(sender As Object, e As EventArgs) Handles btnAgregaEdit.Click
        gbEditorial.Visible = True
        gbEditorial.Enabled = True
        gbEditorial.BringToFront()
        txtNombreEditorial.Text = ""
        dtañoinaguracionEditorial.Value = DateTime.Now
        Dim dt As DataTable = New DataTable()

        Try

            Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                sql.Open()

                Using cmd As SqlCommand = New SqlCommand()
                    cmd.Connection = sql
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = "SELECT c.Id_Ciudad, concat(c.Nombre,', ',es.Nombre,', ',p.Nombre) as localizacion FROM Ciudad c  INNER JOIN Estado es ON c.Id_Estado = es.Id_Estado INNER JOIN Pais p ON es.Id_Pais = p.Id_Pais"

                    Using da As SqlDataAdapter = New SqlDataAdapter(cmd)
                        da.Fill(dt)
                    End Using
                End Using

                sql.Close()
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        cbCiudadEditorial.DataSource = dt
        cbCiudadEditorial.DisplayMember = "localizacion"
        cbCiudadEditorial.ValueMember = "Id_Ciudad"
        chkEditorial.Checked = True
        modEditorial = "0"
    End Sub

    Private Sub btnGuarModCat_Click(sender As Object, e As EventArgs) Handles btnGuarModCat.Click
        If String.IsNullOrEmpty(txtUpCreNombreCat.Text) OrElse String.IsNullOrWhiteSpace(txtUpCreNombreCat.Text) Then
            MessageBox.Show("Ingrese los datos solicitados")
        Else
            Dim a As String = "0"

            If ckbCat.Checked Then
                a = "1"
            End If

            If modCategoria = 0 Then

                Try

                    Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                        sql.Open()
                        cmd = New SqlCommand("CreateCategoria", sql)
                        cmd.CommandType = CommandType.StoredProcedure
                        cmd.Parameters.Add(New SqlParameter("@Nombre", txtUpCreNombreCat.Text))
                        cmd.Parameters.Add(New SqlParameter("@Id_Usuario_Creacion", IDuser))
                        cmd.Parameters.Add(New SqlParameter("@Fecha_Hora_Creacion", DateTime.Now))
                        cmd.Parameters.Add(New SqlParameter("@Id_Usuario_Modificacion", IDuser))
                        cmd.Parameters.Add(New SqlParameter("@Fecha_Hora_Modificacion", DateTime.Now))
                        cmd.Parameters.Add(New SqlParameter("@Estatus", a))
                        cmd.ExecuteReader()
                        LogsAceptable("Insertar Categoria", "Categoria")
                        sql.Close()
                        MessageBox.Show("Guardado")
                        gbCategoria.Visible = False
                    End Using

                Catch ex As Exception
                    MessageBox.Show("Error de conexion")
                End Try
            Else

                Try

                    Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                        sql.Open()
                        Dim consulta As String = "Update Categoria set Nombre='" & txtUpCreNombreCat.Text & "',Estatus=" & a & ",Id_Usuario_Modificacion=" & IDuser.ToString() & ",Fecha_Hora_Modificacion= '" & DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff") & "' where id_Categoria=" & modCategoria.ToString()
                        cmd = New SqlCommand(consulta, sql)
                        cmd.ExecuteReader()
                        LogsAceptable("Modificar categorias", "Categoria")
                        sql.Close()
                        MessageBox.Show("Guardado")
                        gbCategoria.Visible = False
                    End Using

                Catch ex As Exception
                    MessageBox.Show("Error en conexion")
                End Try
            End If

            recargaCatalogoCategoria()
        End If
    End Sub

    Private Sub btnModInsEditorial_Click(sender As Object, e As EventArgs) Handles btnModInsEditorial.Click
        If String.IsNullOrEmpty(txtNombreEditorial.Text) OrElse String.IsNullOrWhiteSpace(txtNombreEditorial.Text) OrElse String.IsNullOrEmpty(cbCiudadEditorial.Text) OrElse String.IsNullOrWhiteSpace(cbCiudadEditorial.Text) OrElse String.IsNullOrEmpty(dtañoinaguracionEditorial.Text) OrElse String.IsNullOrWhiteSpace(dtañoinaguracionEditorial.Text) Then
            MessageBox.Show("Ingrese los datos solicitados")
        Else
            Dim a As String = "0"

            If chkEditorial.Checked Then
                a = "1"
            End If

            If modEditorial.Equals("0") Then

                Try

                    Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                        sql.Open()
                        cmd = New SqlCommand("CreateEditorial", sql)
                        cmd.CommandType = CommandType.StoredProcedure
                        cmd.Parameters.Add(New SqlParameter("@Id_Editorial", SiglaEditExis(txtNombreEditorial.Text)))
                        cmd.Parameters.Add(New SqlParameter("@Nombre", txtNombreEditorial.Text))
                        cmd.Parameters.Add(New SqlParameter("@Id_Ciudad", cbCiudadEditorial.SelectedValue))
                        cmd.Parameters.Add(New SqlParameter("@Año_Inaguracion", dtañoinaguracionEditorial.Text))
                        cmd.Parameters.Add(New SqlParameter("@Id_Usuario_Creacion", IDuser))
                        cmd.Parameters.Add(New SqlParameter("@Fecha_Hora_Creacion", DateTime.Now))
                        cmd.Parameters.Add(New SqlParameter("@Id_Usuario_Modificacion", IDuser))
                        cmd.Parameters.Add(New SqlParameter("@Fecha_Hora_Modificacion", DateTime.Now))
                        cmd.Parameters.Add(New SqlParameter("@Estatus", a))
                        cmd.ExecuteReader()
                        LogsAceptable("Insertar Editorial", "Editorial")
                        sql.Close()
                        MessageBox.Show("Guardado")
                        gbEditorial.Visible = False
                    End Using

                Catch ex As Exception
                    MessageBox.Show("Error de conexion")
                End Try
            Else

                Try

                    Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                        sql.Open()
                        Dim consulta As String = "Update Editorial set Nombre='" & txtNombreEditorial.Text & "',Estatus=" & a & ",Id_Usuario_Modificacion=" & IDuser.ToString() & ",Año_Inauguracion='" + dtañoinaguracionEditorial.Text & "', Id_Ciudad=" + cbCiudadEditorial.SelectedValue.ToString() & ",Fecha_Hora_Modificacion= '" & DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff") & "' where Id_Editorial='" & modEditorial & "'"
                        cmd = New SqlCommand(consulta, sql)
                        cmd.ExecuteReader()
                        LogsAceptable("Modificar Editorial", "Editorial")
                        sql.Close()
                        MessageBox.Show("Guardado")
                        gbEditorial.Visible = False
                    End Using

                Catch ex As Exception
                    MessageBox.Show("Error en conexion")
                End Try
            End If

            recargaCatalogoEditorial()
        End If
    End Sub

    Private Sub btnAgregaCliente_Click(sender As Object, e As EventArgs) Handles btnAgregaCliente.Click
        gbClientes.Visible = True
        gbClientes.Enabled = True
        gbClientes.BringToFront()
        txtNombreCliente.Text = ""
        chkTiCliente.Checked = True
        modCliente = 0
    End Sub

    Private Sub btnModCliente_Click(sender As Object, e As EventArgs) Handles btnModCliente.Click
        If dgCliente.SelectedRows.Count = 1 Then
            txtNombreCliente.Text = CStr(dgCliente.SelectedCells(1).Value)

            If CBool(dgCliente.SelectedCells(2).Value) Then
                chkTiCliente.Checked = True
            Else
                chkTiCliente.Checked = False
            End If

            modCliente = CInt(dgCliente.SelectedCells(0).Value)
            gbClientes.Visible = True
            gbClientes.Enabled = True
            gbClientes.BringToFront()
        Else
            MessageBox.Show("Seleccione un Tipo de Cliente.")
        End If
    End Sub

    Private Sub btnEliminaCliente_Click(sender As Object, e As EventArgs) Handles btnEliminaCliente.Click
        If dgCliente.SelectedRows.Count = 1 Then

            Try

                Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                    sql.Open()
                    cmd = New SqlCommand("DeleteTipo_Cliente", sql)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.Parameters.Add(New SqlParameter("@Id", dgCliente.SelectedCells(0).Value))
                    cmd.ExecuteReader()
                    LogsAceptable("Eliminar Tipo Cliente", "Tipo_Cliente")
                    sql.Close()
                    MessageBox.Show("Eliminada")
                    recargaCatalogoTipoCliente()
                End Using

            Catch ex As Exception
                MessageBox.Show("Error de conexion")
            End Try
        Else
            MessageBox.Show("Seleccione un Tipo de Cliente.")
        End If
    End Sub

    Private Sub btnAgregaPago_Click(sender As Object, e As EventArgs) Handles btnAgregaPago.Click
        gpPago.Visible = True
        gpPago.Enabled = True
        gpPago.BringToFront()
        txtNombrePago.Text = ""
        chkPago.Checked = True
        modPago = 0
    End Sub

    Private Sub btnModificaPago_Click(sender As Object, e As EventArgs) Handles btnModificaPago.Click
        If dgvPago.SelectedRows.Count = 1 Then
            txtNombrePago.Text = CStr(dgvPago.SelectedCells(1).Value)

            If CBool(dgvPago.SelectedCells(2).Value) Then
                chkPago.Checked = True
            Else
                chkPago.Checked = False
            End If

            modPago = CInt(dgvPago.SelectedCells(0).Value)
            gpPago.Visible = True
            gpPago.Enabled = True
            gpPago.BringToFront()
        Else
            MessageBox.Show("Seleccione una Forma de Pago.")
        End If
    End Sub

    Private Sub btnEliminaPago_Click(sender As Object, e As EventArgs) Handles btnEliminaPago.Click
        If dgvPago.SelectedRows.Count = 1 Then

            Try

                Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                    sql.Open()
                    cmd = New SqlCommand("DeleteTipo_Pago", sql)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.Parameters.Add(New SqlParameter("@Id", dgvPago.SelectedCells(0).Value))
                    cmd.ExecuteReader()
                    LogsAceptable("Eliminar Tipo de Pago", "Tipo_Pago")
                    sql.Close()
                    MessageBox.Show("Eliminada")
                    recargaCatalogoTipoPago()
                End Using

            Catch ex As Exception
                MessageBox.Show("Error de conexion")
            End Try
        Else
            MessageBox.Show("Seleccione un Tipo de Pago.")
        End If
    End Sub

    Private Sub btnAgregaModificaPago_Click(sender As Object, e As EventArgs) Handles btnAgregaModificaPago.Click
        If String.IsNullOrEmpty(txtNombrePago.Text) OrElse String.IsNullOrWhiteSpace(txtNombrePago.Text) Then
            MessageBox.Show("Ingrese los datos solicitados")
        Else
            Dim a As String = "0"

            If chkPago.Checked Then
                a = "1"
            End If

            If modPago = 0 Then

                Try

                    Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                        sql.Open()
                        cmd = New SqlCommand("CreateTipoPago", sql)
                        cmd.CommandType = CommandType.StoredProcedure
                        cmd.Parameters.Add(New SqlParameter("@Nombre", txtNombrePago.Text))
                        cmd.Parameters.Add(New SqlParameter("@Id_Usuario_Creacion", IDuser))
                        cmd.Parameters.Add(New SqlParameter("@Fecha_Hora_Creacion", DateTime.Now))
                        cmd.Parameters.Add(New SqlParameter("@Id_Usuario_Modificacion", IDuser))
                        cmd.Parameters.Add(New SqlParameter("@Fecha_Hora_Modificacion", DateTime.Now))
                        cmd.Parameters.Add(New SqlParameter("@Estatus", a))
                        cmd.ExecuteReader()
                        LogsAceptable("Insertar Forma de Pago", "Tipo_Pago")
                        sql.Close()
                        MessageBox.Show("Guardado")
                        gbCategoria.Visible = False
                    End Using

                Catch ex As Exception
                    MessageBox.Show("Error de conexion")
                End Try
            Else

                Try

                    Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                        sql.Open()
                        Dim consulta As String = "Update Tipo_Pago set Nombre='" & txtNombrePago.Text & "',Estatus=" & a & ",Id_Usuario_Modificacion=" & IDuser.ToString() & ",Fecha_Hora_Modificacion= '" & DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff") & "' where Id_Tipo_Pago=" & modPago.ToString()
                        cmd = New SqlCommand(consulta, sql)
                        cmd.ExecuteReader()
                        LogsAceptable("Modificar Forma de Pago", "Tipo_Pago")
                        sql.Close()
                        MessageBox.Show("Guardado")
                        gbCategoria.Visible = False
                    End Using

                Catch ex As Exception
                    MessageBox.Show("Error en conexion")
                End Try
            End If

            recargaCatalogoTipoPago()
        End If
    End Sub

    Private Sub btnAgregaModificaTiCliente_Click(sender As Object, e As EventArgs) Handles btnAgregaModificaTiCliente.Click
        If String.IsNullOrEmpty(txtNombreCliente.Text) OrElse String.IsNullOrWhiteSpace(txtNombreCliente.Text) Then
            MessageBox.Show("Ingrese los datos solicitados")
        Else
            Dim a As String = "0"

            If chkTiCliente.Checked Then
                a = "1"
            End If

            If modCliente = 0 Then

                Try

                    Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                        sql.Open()
                        cmd = New SqlCommand("CreateTipoCliente", sql)
                        cmd.CommandType = CommandType.StoredProcedure
                        cmd.Parameters.Add(New SqlParameter("@Nombre", txtNombreCliente.Text))
                        cmd.Parameters.Add(New SqlParameter("@Id_Usuario_Creacion", IDuser))
                        cmd.Parameters.Add(New SqlParameter("@Fecha_Hora_Creacion", DateTime.Now))
                        cmd.Parameters.Add(New SqlParameter("@Id_Usuario_Modificacion", IDuser))
                        cmd.Parameters.Add(New SqlParameter("@Fecha_Hora_Modificacion", DateTime.Now))
                        cmd.Parameters.Add(New SqlParameter("@Estatus", a))
                        cmd.ExecuteReader()
                        LogsAceptable("Insertar Tipo de Cliente", "Tipo_Cliente")
                        sql.Close()
                        MessageBox.Show("Guardado")
                        gbCategoria.Visible = False
                    End Using

                Catch ex As Exception
                    MessageBox.Show("Error de conexion")
                End Try
            Else

                Try

                    Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                        sql.Open()
                        Dim consulta As String = "Update Tipo_Cliente set Nombre='" & txtNombreCliente.Text & "',Estatus=" & a & ",Id_Usuario_Modificacion=" & IDuser.ToString() & ",Fecha_Hora_Modificacion= '" & DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff") & "' where id_Tipo_Cliente=" & modCliente.ToString()
                        cmd = New SqlCommand(consulta, sql)
                        cmd.ExecuteReader()
                        LogsAceptable("Modificar Tipo de Cliente", "Tipo_Cliente")
                        sql.Close()
                        MessageBox.Show("Guardado")
                        gbCategoria.Visible = False
                    End Using

                Catch ex As Exception
                    MessageBox.Show("Error en conexion")
                End Try
            End If

            recargaCatalogoTipoCliente()
        End If
    End Sub

    Private Sub btnAgregaDepartamento_Click(sender As Object, e As EventArgs) Handles btnAgregaDepartamento.Click
        gbDepartamento.Visible = True
        gbDepartamento.Enabled = True
        gbDepartamento.BringToFront()
        txtNombreDepartamento.Text = ""
        chkDepartamento.Checked = True
        modDepartamento = 0
    End Sub

    Private Sub btnAgregaPuestos_Click(sender As Object, e As EventArgs) Handles btnAgregaPuestos.Click
        gbPuestos.Visible = True
        gbPuestos.Enabled = True
        gbPuestos.BringToFront()
        txtNombrePuestos.Text = ""
        chkPuesto.Checked = True
        modPuesto = 0
    End Sub

    Private Sub btnModificaDepartamento_Click(sender As Object, e As EventArgs) Handles btnModificaDepartamento.Click
        If dgvDepartamentos.SelectedRows.Count = 1 Then
            txtNombreDepartamento.Text = CStr(dgvDepartamentos.SelectedCells(1).Value)

            If CBool(dgvDepartamentos.SelectedCells(2).Value) Then
                chkDepartamento.Checked = True
            Else
                chkDepartamento.Checked = False
            End If

            modDepartamento = CInt(dgvDepartamentos.SelectedCells(0).Value)
            gbDepartamento.Visible = True
            gbDepartamento.Enabled = True
            gbDepartamento.BringToFront()
        Else
            MessageBox.Show("Seleccione un Departamento.")
        End If
    End Sub

    Private Sub btnModificaPuestos_Click(sender As Object, e As EventArgs) Handles btnModificaPuestos.Click
        If dgvPuestos.SelectedRows.Count = 1 Then
            txtNombrePuestos.Text = CStr(dgvPuestos.SelectedCells(1).Value)

            If CBool(dgvPuestos.SelectedCells(2).Value) Then
                chkPuesto.Checked = True
            Else
                chkPuesto.Checked = False
            End If

            modPuesto = CInt(dgvPuestos.SelectedCells(0).Value)
            gbPuestos.Visible = True
            gbPuestos.Enabled = True
            gbPuestos.BringToFront()
        Else
            MessageBox.Show("Seleccione un Puesto.")
        End If
    End Sub

    Private Sub btnAgregaModificaPuesto_Click(sender As Object, e As EventArgs) Handles btnAgregaModificaPuesto.Click
        If String.IsNullOrEmpty(txtNombrePuestos.Text) OrElse String.IsNullOrWhiteSpace(txtNombrePuestos.Text) Then
            MessageBox.Show("Ingrese los datos solicitados")
        Else
            Dim a As String = "0"

            If chkPuesto.Checked Then
                a = "1"
            End If

            If modPuesto = 0 Then

                Try

                    Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                        sql.Open()
                        cmd = New SqlCommand("CreatePuesto", sql)
                        cmd.CommandType = CommandType.StoredProcedure
                        cmd.Parameters.Add(New SqlParameter("@Nombre", txtNombrePuestos.Text))
                        cmd.Parameters.Add(New SqlParameter("@Id_Usuario_Creacion", IDuser))
                        cmd.Parameters.Add(New SqlParameter("@Fecha_Hora_Creacion", DateTime.Now))
                        cmd.Parameters.Add(New SqlParameter("@Id_Usuario_Modificacion", IDuser))
                        cmd.Parameters.Add(New SqlParameter("@Fecha_Hora_Modificacion", DateTime.Now))
                        cmd.Parameters.Add(New SqlParameter("@Estatus", a))
                        cmd.ExecuteReader()
                        LogsAceptable("Insertar Puesto", "Puesto")
                        sql.Close()
                        MessageBox.Show("Guardado")
                        gbPuestos.Visible = False
                    End Using

                Catch ex As Exception
                    MessageBox.Show("Error de conexion")
                End Try
            Else

                Try

                    Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                        sql.Open()
                        Dim consulta As String = "Update Puesto set Nombre='" & txtNombrePuestos.Text & "',Estatus=" & a & ",Id_Usuario_Modificacion=" & IDuser.ToString() & ",Fecha_Hora_Modificacion= '" & DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff") & "' where Id_Puesto=" & modPuesto.ToString()
                        cmd = New SqlCommand(consulta, sql)
                        cmd.ExecuteReader()
                        LogsAceptable("Modificar Puesto", "Puesto")
                        sql.Close()
                        MessageBox.Show("Guardado")
                        gbPuestos.Visible = False
                    End Using

                Catch ex As Exception
                    MessageBox.Show("Error en conexion")
                End Try
            End If

            recargaCatalogoPuestos()
        End If
    End Sub

    Private Sub btnAgregaModificaDepartamento_Click(sender As Object, e As EventArgs) Handles btnAgregaModificaDepartamento.Click
        If String.IsNullOrEmpty(txtNombreDepartamento.Text) OrElse String.IsNullOrWhiteSpace(txtNombreDepartamento.Text) Then
            MessageBox.Show("Ingrese los datos solicitados")
        Else
            Dim a As String = "0"

            If chkDepartamento.Checked Then
                a = "1"
            End If

            If modDepartamento = 0 Then

                Try

                    Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                        sql.Open()
                        cmd = New SqlCommand("CreateDepartamento", sql)
                        cmd.CommandType = CommandType.StoredProcedure
                        cmd.Parameters.Add(New SqlParameter("@Departamento", txtNombreDepartamento.Text))
                        cmd.Parameters.Add(New SqlParameter("@Id_Usuario_Creacion", IDuser))
                        cmd.Parameters.Add(New SqlParameter("@Fecha_Hora_Creacion", DateTime.Now))
                        cmd.Parameters.Add(New SqlParameter("@Id_Usuario_Modificacion", IDuser))
                        cmd.Parameters.Add(New SqlParameter("@Fecha_Hora_Modificacion", DateTime.Now))
                        cmd.Parameters.Add(New SqlParameter("@Estatus", a))
                        cmd.ExecuteReader()
                        LogsAceptable("Insertar Departamento", "Departamento")
                        sql.Close()
                        MessageBox.Show("Guardado")
                        gbDepartamento.Visible = False
                    End Using

                Catch ex As Exception
                    MessageBox.Show("Error de conexion")
                End Try
            Else

                Try

                    Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                        sql.Open()
                        Dim consulta As String = "Update Departamento set Departamento='" & txtNombreDepartamento.Text & "',Estatus=" & a & ",Id_Usuario_Modificacion=" & IDuser.ToString() & ",Fecha_Hora_Modificacion= '" & DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff") & "' where Id_Departamento=" & modDepartamento.ToString()
                        cmd = New SqlCommand(consulta, sql)
                        cmd.ExecuteReader()
                        LogsAceptable("Modificar Departamento", "Departamento")
                        sql.Close()
                        MessageBox.Show("Guardado")
                        gbDepartamento.Visible = False
                    End Using

                Catch ex As Exception
                    MessageBox.Show("Error en conexion")
                End Try
            End If

            recargaCatalogoDepartamento()
        End If
    End Sub

    Private Sub btnEliminaDepartamento_Click(sender As Object, e As EventArgs) Handles btnEliminaDepartamento.Click
        If dgvDepartamentos.SelectedRows.Count = 1 Then

            Try

                Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                    sql.Open()
                    cmd = New SqlCommand("DeleteDepartamento", sql)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.Parameters.Add(New SqlParameter("@Id", dgvDepartamentos.SelectedCells(0).Value))
                    cmd.ExecuteReader()
                    LogsAceptable("Eliminar Departamento", "Departamento")
                    sql.Close()
                    MessageBox.Show("Eliminada")
                    recargaCatalogoDepartamento()
                End Using

            Catch ex As Exception
                MessageBox.Show("Error de conexion")
            End Try
        Else
            MessageBox.Show("Seleccione un Departamento.")
        End If
    End Sub

    Private Sub btnEliminaPuestos_Click(sender As Object, e As EventArgs) Handles btnEliminaPuestos.Click
        If dgvPuestos.SelectedRows.Count = 1 Then

            Try

                Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                    sql.Open()
                    cmd = New SqlCommand("DeletePuesto", sql)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.Parameters.Add(New SqlParameter("@Id", dgvPuestos.SelectedCells(0).Value))
                    cmd.ExecuteReader()
                    LogsAceptable("Eliminar Puesto", "Puesto")
                    sql.Close()
                    MessageBox.Show("Eliminada")
                    recargaCatalogoPuestos()
                End Using

            Catch ex As Exception
                MessageBox.Show("Error de conexion")
            End Try
        Else
            MessageBox.Show("Seleccione un Puesto.")
        End If
    End Sub

    Private Sub btnAgregaFormatos_Click(sender As Object, e As EventArgs) Handles btnAgregaFormatos.Click
        gbFormato.Visible = True
        gbFormato.Enabled = True
        gbFormato.BringToFront()
        txtNombreFormato.Text = ""
        chkFormato.Checked = True
        modFormato = 0
    End Sub

    Private Sub btnAgregarLibro_Click(sender As Object, e As EventArgs) Handles btnAgregarLibro.Click
        dgvAutoresLibro.Rows.Clear()
        panelLibros.Visible = True
        panelLibros.Enabled = True
        panelLibros.BringToFront()
        txtISBN.Text = ""
        txtISBN.Enabled = True
        txtTitulo.Text = ""
        cargacombosLibros()
        chkLibro.Checked = True
        modLibro = "0"
        txtVolumenLibro.Text = ""
        dtFechaLibro.Value = DateTime.Now
        txtPaginasLibro.Text = ""
        txtPrecioLibro.Text = ""
        dgvAutoresLibro.DataSource = Nothing
    End Sub

    Private Sub btnModificaFormato_Click(sender As Object, e As EventArgs) Handles btnModificaFormato.Click
        If dgvFormatos.SelectedRows.Count = 1 Then
            txtNombreFormato.Text = CStr(dgvFormatos.SelectedCells(1).Value)

            If CBool(dgvFormatos.SelectedCells(2).Value) Then
                chkFormato.Checked = True
            Else
                chkFormato.Checked = False
            End If

            modFormato = CInt(dgvFormatos.SelectedCells(0).Value)
            gbFormato.Visible = True
            gbFormato.Enabled = True
            gbFormato.BringToFront()
        Else
            MessageBox.Show("Seleccione un Formato.")
        End If
    End Sub
    Private Sub cargaLibrosAutoresTabla(ByVal isbn As String)
        dgvAutoresLibro.Rows.Clear()
        Dim dt As DataTable = New DataTable()

        Try

            Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                sql.Open()

                Using cmd As SqlCommand = New SqlCommand()
                    cmd.Connection = sql
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = "select l.Id_Autor,CONCAT(a.Nombre,' ',a.Apellido) Nombre_Completo from Libro_Autores l left join Autor a on l.Id_Autor=a.Id_Autor where l.Id_Libro_ISBN='" & isbn & "'"

                    Using da As SqlDataAdapter = New SqlDataAdapter(cmd)
                        da.Fill(dt)
                    End Using
                End Using

                sql.Close()
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        For i As Integer = 0 To dt.Rows.Count - 1
            dgvAutoresLibro.Rows.Add(dt.Rows(i)("Id_Autor").ToString(), dt.Rows(i)("Nombre_Completo").ToString())
        Next
    End Sub
    Private Sub btnModificaLibro_Click(sender As Object, e As EventArgs) Handles btnModificaLibro.Click
        If dgvLibros.SelectedRows.Count = 1 Then
            cargacombosLibros()
            txtISBN.Text = CStr(dgvLibros.SelectedCells(0).Value)
            txtISBN.Enabled = False
            cargaLibrosAutoresTabla(txtISBN.Text)
            txtTitulo.Text = CStr(dgvLibros.SelectedCells(1).Value)
            cbCategoriaLibro.SelectedValue = CInt(dgvLibros.SelectedCells(2).Value)
            txtVolumenLibro.Text = CStr(dgvLibros.SelectedCells(4).Value)
            cbFormatoLibro.SelectedValue = CInt(dgvLibros.SelectedCells(5).Value)
            cbEditorialLibro.SelectedValue = CStr(dgvLibros.SelectedCells(7).Value)
            dtFechaLibro.Value = CDate(dgvLibros.SelectedCells(9).Value)
            txtPaginasLibro.Text = (CInt(dgvLibros.SelectedCells(10).Value)).ToString()
            txtPrecioLibro.Text = (CDec(dgvLibros.SelectedCells(11).Value)).ToString()
            cbRatingLibro.SelectedValue = CInt(dgvLibros.SelectedCells(12).Value)

            If CBool(dgvLibros.SelectedCells(14).Value) Then
                chkLibro.Checked = True
            Else
                chkLibro.Checked = False
            End If

            modLibro = CStr(dgvLibros.SelectedCells(0).Value)
            panelLibros.Visible = True
            panelLibros.Enabled = True
            panelLibros.BringToFront()
        Else
            MessageBox.Show("Seleccione un Libro.")
        End If
    End Sub

    Private Sub btnEliminaFormato_Click(sender As Object, e As EventArgs) Handles btnEliminaFormato.Click
        If dgvFormatos.SelectedRows.Count = 1 Then

            Try

                Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                    sql.Open()
                    cmd = New SqlCommand("DeleteTipo_Formato", sql)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.Parameters.Add(New SqlParameter("@Id", dgvFormatos.SelectedCells(0).Value))
                    cmd.ExecuteReader()
                    LogsAceptable("Eliminar Formato", "Tipo_Formato")
                    sql.Close()
                    MessageBox.Show("Eliminada")
                    recargaCatalogoTipoFormato()
                End Using

            Catch ex As Exception
                MessageBox.Show("Error de conexion")
            End Try
        Else
            MessageBox.Show("Seleccione un Formato.")
        End If
    End Sub

    Private Sub btnEliminaLibro_Click(sender As Object, e As EventArgs) Handles btnEliminaLibro.Click
        If dgvLibros.SelectedRows.Count = 1 Then

            Try

                Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                    sql.Open()
                    cmd = New SqlCommand("DeleteLibro", sql)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.Parameters.Add(New SqlParameter("@Id", dgvLibros.SelectedCells(0).Value))
                    cmd.ExecuteReader()
                    LogsAceptable("Eliminar Libro", "Libro")
                    sql.Close()
                    MessageBox.Show("Eliminada")
                    recargaCatalogoLibro()
                End Using

            Catch ex As Exception
                MessageBox.Show("Error de conexion")
            End Try
        Else
            MessageBox.Show("Seleccione un Libro.")
        End If
    End Sub

    Private Sub btnAgregarModificarFormato_Click(sender As Object, e As EventArgs) Handles btnAgregarModificarFormato.Click
        If String.IsNullOrEmpty(txtNombreFormato.Text) OrElse String.IsNullOrWhiteSpace(txtNombreFormato.Text) Then
            MessageBox.Show("Ingrese los datos solicitados")
        Else
            Dim a As String = "0"

            If chkFormato.Checked Then
                a = "1"
            End If

            If modFormato = 0 Then

                Try

                    Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                        sql.Open()
                        cmd = New SqlCommand("CreateTipoFormato", sql)
                        cmd.CommandType = CommandType.StoredProcedure
                        cmd.Parameters.Add(New SqlParameter("@Nombre", txtNombreFormato.Text))
                        cmd.Parameters.Add(New SqlParameter("@Id_Usuario_Creacion", IDuser))
                        cmd.Parameters.Add(New SqlParameter("@Fecha_Hora_Creacion", DateTime.Now))
                        cmd.Parameters.Add(New SqlParameter("@Id_Usuario_Modificacion", IDuser))
                        cmd.Parameters.Add(New SqlParameter("@Fecha_Hora_Modificacion", DateTime.Now))
                        cmd.Parameters.Add(New SqlParameter("@Estatus", a))
                        cmd.ExecuteReader()
                        LogsAceptable("Insertar Formato", "Tipo_Formato")
                        sql.Close()
                        MessageBox.Show("Guardado")
                        gbFormato.Visible = False
                    End Using

                Catch ex As Exception
                    MessageBox.Show("Error de conexion")
                End Try
            Else

                Try

                    Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                        sql.Open()
                        Dim consulta As String = "Update Tipo_Formato set Nombre='" & txtNombreFormato.Text & "',Estatus=" & a & ",Id_Usuario_Modificacion=" & IDuser.ToString() & ",Fecha_Hora_Modificacion= '" & DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff") & "' where Id_Tipo_Formato=" & modFormato.ToString()
                        cmd = New SqlCommand(consulta, sql)
                        cmd.ExecuteReader()
                        LogsAceptable("Modificar Formato", "Tipo_Formato")
                        sql.Close()
                        MessageBox.Show("Guardado")
                        gbFormato.Visible = False
                    End Using

                Catch ex As Exception
                    MessageBox.Show("Error en conexion")
                End Try
            End If

            recargaCatalogoTipoFormato()
        End If
    End Sub

    Private Sub btnAgregarModificarLibro_Click(sender As Object, e As EventArgs) Handles btnAgregarModificarLibro.Click
        If String.IsNullOrEmpty(txtISBN.Text) OrElse String.IsNullOrWhiteSpace(txtISBN.Text) OrElse String.IsNullOrEmpty(txtTitulo.Text) OrElse String.IsNullOrWhiteSpace(txtTitulo.Text) OrElse String.IsNullOrEmpty(txtPrecioLibro.Text) OrElse String.IsNullOrWhiteSpace(txtPrecioLibro.Text) OrElse String.IsNullOrEmpty(cbCategoriaLibro.Text) OrElse String.IsNullOrWhiteSpace(cbCategoriaLibro.Text) OrElse String.IsNullOrEmpty(cbFormatoLibro.Text) OrElse String.IsNullOrWhiteSpace(cbFormatoLibro.Text) OrElse String.IsNullOrEmpty(cbEditorialLibro.Text) OrElse String.IsNullOrWhiteSpace(cbEditorialLibro.Text) OrElse String.IsNullOrEmpty(cbRatingLibro.Text) OrElse String.IsNullOrWhiteSpace(cbRatingLibro.Text) Then
            MessageBox.Show("Ingrese los datos solicitados")
        Else
            Dim vol As String = Nothing
            Dim f As String = Nothing
            Dim pag As String = Nothing

            If Not (String.IsNullOrEmpty(txtVolumenLibro.Text) OrElse String.IsNullOrWhiteSpace(txtVolumenLibro.Text)) Then
                vol = txtVolumenLibro.Text
            End If

            If Not (dtFechaLibro.Text.Equals("")) Then
                f = dtFechaLibro.Value.ToShortDateString()
            End If

            If Not (txtPaginasLibro.Text.Equals("")) Then
                pag = txtPaginasLibro.Text
            End If

            Dim a As String = "0"

            If chkLibro.Checked Then
                a = "1"
            End If

            If modLibro.Equals("0") Then

                Try

                    Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                        sql.Open()
                        cmd = New SqlCommand("CreateLibro", sql)
                        cmd.CommandType = CommandType.StoredProcedure
                        cmd.Parameters.Add(New SqlParameter("@Id_Libro_ISBN", txtISBN.Text))
                        cmd.Parameters.Add(New SqlParameter("@Titulo", txtTitulo.Text))
                        cmd.Parameters.Add(New SqlParameter("@Id_Categoria", cbCategoriaLibro.SelectedValue))
                        cmd.Parameters.Add(New SqlParameter("@Volumen", vol))
                        cmd.Parameters.Add(New SqlParameter("@Id_Tipo_Formato", cbFormatoLibro.SelectedValue))
                        cmd.Parameters.Add(New SqlParameter("@Id_Editorial", cbEditorialLibro.SelectedValue))
                        cmd.Parameters.Add(New SqlParameter("@Fecha_Publicacion", f))
                        cmd.Parameters.Add(New SqlParameter("@Paginas", pag))
                        cmd.Parameters.Add(New SqlParameter("@Precio", txtPrecioLibro.Text))
                        cmd.Parameters.Add(New SqlParameter("@Id_Rating", cbRatingLibro.SelectedValue))
                        cmd.Parameters.Add(New SqlParameter("@Id_Usuario_Creacion", IDuser))
                        cmd.Parameters.Add(New SqlParameter("@Fecha_Hora_Creacion", DateTime.Now))
                        cmd.Parameters.Add(New SqlParameter("@Id_Usuario_Modificacion", IDuser))
                        cmd.Parameters.Add(New SqlParameter("@Fecha_Hora_Modificacion", DateTime.Now))
                        cmd.Parameters.Add(New SqlParameter("@Estatus", a))
                        cmd.ExecuteReader()
                        EliminaTodosAutoresLibro(txtISBN.Text)

                        For i As Integer = 0 To dgvAutoresLibro.Rows.Count - 1

                            If Not (dgvAutoresLibro.Rows(i).Cells("Id_Autor").Value.ToString()).Equals("") Then
                                agregarAutoresLibro(txtISBN.Text, dgvAutoresLibro.Rows(i).Cells("Id_Autor").Value.ToString())
                            End If
                        Next

                        LogsAceptable("Insertar Libro", "Libro")
                        sql.Close()
                        MessageBox.Show("Guardado")
                        panelLibros.Visible = False
                    End Using

                Catch ex As Exception
                    MessageBox.Show("Error de conexion")
                End Try
            Else

                Try

                    Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                        sql.Open()

                        If Not (String.IsNullOrEmpty(txtVolumenLibro.Text) OrElse String.IsNullOrWhiteSpace(txtVolumenLibro.Text)) Then
                            vol = "'" & txtVolumenLibro.Text & "'"
                        Else
                            vol = "NULL"
                        End If

                        If Not (dtFechaLibro.Text.Equals("")) Then
                            f = "'" & dtFechaLibro.Value.ToShortDateString() & "'"
                        Else
                            f = "NULL"
                        End If

                        If Not (txtPaginasLibro.Text.Equals("")) Then
                            pag = txtPaginasLibro.Text
                        Else
                            pag = "NULL"
                        End If

                        Dim consulta As String = "Update Libro set Titulo='" & txtTitulo.Text & "',Estatus=" & a & ",Id_Usuario_Modificacion=" & IDuser.ToString() & ",Fecha_Hora_Modificacion= '" & DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff") & "'," & "Id_Categoria=" + cbCategoriaLibro.SelectedValue.ToString() & ",Volumen=" & vol & "," & "Paginas=" & pag & ",Id_Tipo_Formato=" + cbEditorialLibro.SelectedValue.ToString() & ",Id_Editorial=" + cbEditorialLibro.SelectedValue.ToString() & ",Id_Rating=" + cbRatingLibro.SelectedValue.ToString() & ",Precio=" + txtPrecioLibro.Text & ",Fecha_Publicacion=" & f & " where Id_Libro_ISBN='" & modLibro.ToString() & "'"
                        cmd = New SqlCommand(consulta, sql)
                        cmd.ExecuteReader()
                        LogsAceptable("Modificar Libro", "Libro")
                        EliminaTodosAutoresLibro(txtISBN.Text)

                        For i As Integer = 0 To dgvAutoresLibro.Rows.Count - 1

                            If Not (dgvAutoresLibro.Rows(i).Cells("Id_Autor").Value.ToString()).Equals("") Then
                                agregarAutoresLibro(txtISBN.Text, dgvAutoresLibro.Rows(i).Cells("Id_Autor").Value.ToString())
                            End If
                        Next

                        sql.Close()
                        MessageBox.Show("Guardado")
                        panelLibros.Visible = False
                    End Using

                Catch ex As Exception
                    MessageBox.Show("Error en conexion")
                End Try
            End If

            recargaCatalogoLibro()
        End If
    End Sub
    Private Sub agregarAutoresLibro(ByVal isbn As String, ByVal autor As String)
        Try

            Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                sql.Open()
                cmd = New SqlCommand("CreateLibroAutores", sql)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add(New SqlParameter("@Id_Libro_ISBN", isbn))
                cmd.Parameters.Add(New SqlParameter("@Id_Autor", autor))
                cmd.Parameters.Add(New SqlParameter("@Id_Usuario_Creacion", IDuser))
                cmd.Parameters.Add(New SqlParameter("@Fecha_Hora_Creacion", DateTime.Now))
                cmd.Parameters.Add(New SqlParameter("@Id_Usuario_Modificacion", IDuser))
                cmd.Parameters.Add(New SqlParameter("@Fecha_Hora_Modificacion", DateTime.Now))
                cmd.Parameters.Add(New SqlParameter("@Estatus", "1"))
                cmd.ExecuteReader()
                sql.Close()
            End Using

        Catch ex As Exception
            Throw
        End Try
    End Sub
    Private Sub btnagregaAutorLibro_Click(sender As Object, e As EventArgs) Handles btnagregaAutorLibro.Click
        Dim existe As Boolean = False
        Dim id As String = cbAutoresLibro.SelectedValue.ToString()

        For i As Integer = 0 To dgvAutoresLibro.Rows.Count - 1

            If (dgvAutoresLibro.Rows(i).Cells("Id_Autor").Value.ToString()).Equals(id) Then
                existe = True
            End If
        Next

        If Not existe Then
            dgvAutoresLibro.Rows.Add(id, cbAutoresLibro.Text)
        End If
    End Sub

    Private Sub btnQuitarLibroAutores_Click(sender As Object, e As EventArgs) Handles btnQuitarLibroAutores.Click
        If dgvAutoresLibro.SelectedRows.Count = 1 Then
            dgvAutoresLibro.Rows.Remove(dgvAutoresLibro.SelectedRows(0))
        Else
            MessageBox.Show("Seleccione un Autor.")
        End If
    End Sub

    Private Sub EliminaTodosAutoresLibro(ByVal isbn As String)
        Try

            Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                sql.Open()
                Dim consulta As String = "delete from Libro_Autores where Id_Libro_ISBN='" & isbn & "'"
                cmd = New SqlCommand(consulta, sql)
                cmd.ExecuteReader()
                sql.Close()
            End Using

        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Sub btnAgregaAutor_Click(sender As Object, e As EventArgs) Handles btnAgregaAutor.Click
        PanelAutorAM.Visible = True
        PanelAutorAM.Enabled = True
        PanelAutorAM.BringToFront()
        txtNombreAutor.Text = ""
        txtApellidoAutor.Text = ""
        dtNacioAutor.Value = DateTime.Now
        Dim dt As DataTable = New DataTable()

        Try

            Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                sql.Open()

                Using cmd As SqlCommand = New SqlCommand()
                    cmd.Connection = sql
                    cmd.CommandType = CommandType.Text
                    cmd.CommandText = "select Id_Pais,Nacionalidad from Pais"

                    Using da As SqlDataAdapter = New SqlDataAdapter(cmd)
                        da.Fill(dt)
                    End Using
                End Using

                sql.Close()
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        cbNacionalidadAutor.DataSource = dt
        cbNacionalidadAutor.DisplayMember = "Nacionalidad"
        cbNacionalidadAutor.ValueMember = "Id_Pais"
        chkAutor.Checked = True
        modAutor = "0"
    End Sub

    Private Sub btnModificaAutor_Click(sender As Object, e As EventArgs) Handles btnModificaAutor.Click
        If dgvAutor.SelectedRows.Count = 1 Then
            Dim dt As DataTable = New DataTable()

            Try

                Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                    sql.Open()

                    Using cmd As SqlCommand = New SqlCommand()
                        cmd.Connection = sql
                        cmd.CommandType = CommandType.Text
                        cmd.CommandText = "select Id_Pais,Nacionalidad from Pais"

                        Using da As SqlDataAdapter = New SqlDataAdapter(cmd)
                            da.Fill(dt)
                        End Using
                    End Using

                    sql.Close()
                End Using

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

            cbNacionalidadAutor.DataSource = dt
            cbNacionalidadAutor.DisplayMember = "Nacionalidad"
            cbNacionalidadAutor.ValueMember = "Id_Pais"
            cbNacionalidadAutor.SelectedValue = CInt(dgvAutor.SelectedCells(5).Value)

            If CBool(dgvAutor.SelectedCells(7).Value) Then
                chkAutor.Checked = True
            Else
                chkAutor.Checked = False
            End If

            txtNombreAutor.Text = CStr(dgvAutor.SelectedCells(1).Value)
            txtApellidoAutor.Text = CStr(dgvAutor.SelectedCells(2).Value)
            dtNacioAutor.Value = CDate(dgvAutor.SelectedCells(4).Value)
            modAutor = CStr(dgvAutor.SelectedCells(0).Value)
            PanelAutorAM.Visible = True
            PanelAutorAM.Enabled = True
            PanelAutorAM.BringToFront()
        Else
            MessageBox.Show("Seleccione un Autor.")
        End If
    End Sub

    Private Sub btnEliminaAutor_Click(sender As Object, e As EventArgs) Handles btnEliminaAutor.Click
        If dgvAutor.SelectedRows.Count = 1 Then

            Try

                Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                    sql.Open()
                    cmd = New SqlCommand("DeleteAutor", sql)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.Parameters.Add(New SqlParameter("@Id", dgvAutor.SelectedCells(0).Value))
                    cmd.ExecuteReader()
                    LogsAceptable("Eliminar Autor", "Autor")
                    sql.Close()
                    MessageBox.Show("Eliminada")
                    recargaCatalogoAutor()
                End Using

            Catch ex As Exception
                MessageBox.Show("Error de conexion")
            End Try
        Else
            MessageBox.Show("Seleccione un Autor.")
        End If
    End Sub

    Private Sub btnAgregarModificarAutor_Click(sender As Object, e As EventArgs) Handles btnAgregarModificarAutor.Click
        If String.IsNullOrEmpty(txtNombreAutor.Text) OrElse String.IsNullOrWhiteSpace(txtNombreAutor.Text) OrElse String.IsNullOrEmpty(txtApellidoAutor.Text) OrElse String.IsNullOrWhiteSpace(txtApellidoAutor.Text) OrElse String.IsNullOrEmpty(dtNacioAutor.Text) OrElse String.IsNullOrWhiteSpace(dtNacioAutor.Text) Then
            MessageBox.Show("Ingrese los datos solicitados")
        Else
            Dim a As String = "0"

            If chkAutor.Checked Then
                a = "1"
            End If

            If modAutor.Equals("0") Then
                Dim f As String = Nothing
                Dim p As String = Nothing

                If Not (cbNacionalidadAutor.Text.Equals("")) Then
                    p = cbNacionalidadAutor.SelectedValue.ToString()
                End If

                If Not (dtNacioAutor.Text.Equals("")) Then
                    f = dtNacioAutor.Value.ToShortDateString()
                End If

                Try

                    Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                        sql.Open()
                        cmd = New SqlCommand("CreateAutor", sql)
                        cmd.CommandType = CommandType.StoredProcedure
                        cmd.Parameters.Add(New SqlParameter("@Id_Autor", SiglaAutor(txtNombreAutor.Text, txtApellidoAutor.Text, dtNacioAutor.Value.ToString("ddMMyyyy"))))
                        cmd.Parameters.Add(New SqlParameter("@Nombre", txtNombreAutor.Text))
                        cmd.Parameters.Add(New SqlParameter("@Apellido", txtApellidoAutor.Text))
                        cmd.Parameters.Add(New SqlParameter("@Fecha_Nacimiento", f))
                        cmd.Parameters.Add(New SqlParameter("@Id_Pais", p))
                        cmd.Parameters.Add(New SqlParameter("@Id_Usuario_Creacion", IDuser))
                        cmd.Parameters.Add(New SqlParameter("@Fecha_Hora_Creacion", DateTime.Now))
                        cmd.Parameters.Add(New SqlParameter("@Id_Usuario_Modificacion", IDuser))
                        cmd.Parameters.Add(New SqlParameter("@Fecha_Hora_Modificacion", DateTime.Now))
                        cmd.Parameters.Add(New SqlParameter("@Estatus", a))
                        cmd.ExecuteReader()
                        LogsAceptable("Insertar Autor", "Autor")
                        sql.Close()
                        MessageBox.Show("Guardado")
                        PanelAutorAM.Visible = False
                    End Using

                Catch ex As Exception
                    MessageBox.Show("Error de conexion")
                End Try
            Else

                Try
                    Dim f As String = "NULL"
                    Dim p As String = "NULL"

                    If Not (cbNacionalidadAutor.Text.Equals("")) Then
                        p = cbNacionalidadAutor.SelectedValue.ToString()
                    End If

                    If Not (dtNacioAutor.Text.Equals("")) Then
                        f = "'" & dtNacioAutor.Value.ToShortDateString() & "'"
                    End If

                    Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                        sql.Open()
                        Dim consulta As String = "Update Autor set Nombre='" & txtNombreAutor.Text & "',Estatus=" & a & ",Id_Usuario_Modificacion=" & IDuser.ToString() & ",Apellido='" + txtApellidoAutor.Text & "',Fecha_Nacimiento=" & f & ", Id_Pais=" & p & ",Fecha_Hora_Modificacion= '" & DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff") & "' where Id_Autor='" & modAutor & "'"
                        cmd = New SqlCommand(consulta, sql)
                        cmd.ExecuteReader()
                        LogsAceptable("Modificar Autor", "Autor")
                        sql.Close()
                        MessageBox.Show("Guardado")
                        PanelAutorAM.Visible = False
                    End Using

                Catch ex As Exception
                    MessageBox.Show("Error en conexion")
                End Try
            End If

            recargaCatalogoAutor()
        End If
    End Sub
    Private Function SiglaAutor(ByVal nombre As String, ByVal apellido As String, ByVal fecha_nacimiento As String) As String
        Dim resul As String = ""
        resul = nombre.Substring(0, 1) & apellido.Substring(0, 1) & fecha_nacimiento
        Return resul.ToUpper()
    End Function

    Private Sub btnAgregaUsuario_Click(sender As Object, e As EventArgs) Handles btnAgregaUsuario.Click
        cargacombosUsuarios()
        panelUsuarioAM.Visible = True
        panelUsuarioAM.Enabled = True
        panelUsuarioAM.BringToFront()
        txtNombreUsuarioUAM.Text = ""
        txtNombreUAM.Text = ""
        txtContraseñaUAM.Text = ""
        txtApellidoPatUAM.Text = ""
        txtApellidoMatUAM.Text = ""
        txtCorreoUAM.Text = ""
        txtTelefonoUAM.Text = ""
        chkUsuarioUAM.Checked = True
        modUsuario = 0
        modcon = ""
    End Sub

    Private Sub btnModificaUsuario_Click(sender As Object, e As EventArgs) Handles btnModificaUsuario.Click
        If dgvUsuarios.SelectedRows.Count = 1 Then
            cargacombosUsuarios()
            txtNombreUsuarioUAM.Text = CStr(dgvUsuarios.SelectedCells(1).Value)
            txtNombreUAM.Text = CStr(dgvUsuarios.SelectedCells(2).Value)
            txtContraseñaUAM.Text = CStr(dgvUsuarios.SelectedCells(3).Value)
            modcon = txtContraseñaUAM.Text
            txtApellidoPatUAM.Text = CStr(dgvUsuarios.SelectedCells(4).Value)
            txtApellidoMatUAM.Text = CStr(dgvUsuarios.SelectedCells(5).Value)
            txtCorreoUAM.Text = CStr(dgvUsuarios.SelectedCells(6).Value)
            txtTelefonoUAM.Text = CStr(dgvUsuarios.SelectedCells(7).Value)
            cbPuestoUAM.SelectedValue = CInt(dgvUsuarios.SelectedCells(8).Value)
            cbDepartamentoUAM.SelectedValue = CInt(dgvUsuarios.SelectedCells(10).Value)
            cbTipoUsuarioUAM.SelectedValue = CInt(dgvUsuarios.SelectedCells(12).Value)

            If CBool(dgvUsuarios.SelectedCells(14).Value) Then
                chkUsuarioUAM.Checked = True
            Else
                chkUsuarioUAM.Checked = False
            End If

            modUsuario = CInt(dgvUsuarios.SelectedCells(0).Value)
            panelUsuarioAM.Visible = True
            panelUsuarioAM.Enabled = True
            panelUsuarioAM.BringToFront()
        Else
            MessageBox.Show("Seleccione un Usuario.")
        End If
    End Sub

    Private Sub btnEliminaUsuario_Click(sender As Object, e As EventArgs) Handles btnEliminaUsuario.Click
        If dgvUsuarios.SelectedRows.Count = 1 Then

            Try

                Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                    sql.Open()
                    cmd = New SqlCommand("DeleteUsuario", sql)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.Parameters.Add(New SqlParameter("@Id", dgvUsuarios.SelectedCells(0).Value))
                    cmd.ExecuteReader()
                    LogsAceptable("Eliminar Usuario", "Usuario")
                    sql.Close()
                    MessageBox.Show("Eliminada")
                    recargaCatalogoUsuario()
                End Using

            Catch ex As Exception
                MessageBox.Show("Error de conexion")
            End Try
        Else
            MessageBox.Show("Seleccione un Usuario.")
        End If
    End Sub

    Private Sub btnAgregarModificarUsuario_Click(sender As Object, e As EventArgs) Handles btnAgregarModificarUsuario.Click
        If String.IsNullOrEmpty(txtNombreUAM.Text) OrElse String.IsNullOrWhiteSpace(txtNombreUAM.Text) OrElse String.IsNullOrEmpty(txtNombreUsuarioUAM.Text) OrElse String.IsNullOrWhiteSpace(txtNombreUsuarioUAM.Text) OrElse String.IsNullOrEmpty(txtApellidoPatUAM.Text) OrElse String.IsNullOrWhiteSpace(txtApellidoPatUAM.Text) OrElse String.IsNullOrEmpty(txtCorreoUAM.Text) OrElse String.IsNullOrWhiteSpace(txtCorreoUAM.Text) OrElse String.IsNullOrEmpty(cbPuestoUAM.Text) OrElse String.IsNullOrWhiteSpace(cbPuestoUAM.Text) OrElse String.IsNullOrEmpty(cbDepartamentoUAM.Text) OrElse String.IsNullOrWhiteSpace(cbDepartamentoUAM.Text) OrElse String.IsNullOrEmpty(cbTipoUsuarioUAM.Text) OrElse String.IsNullOrWhiteSpace(cbTipoUsuarioUAM.Text) OrElse ((String.IsNullOrEmpty(txtContraseñaUAM.Text) OrElse String.IsNullOrWhiteSpace(txtContraseñaUAM.Text)) AndAlso (String.IsNullOrEmpty(modcon) OrElse String.IsNullOrWhiteSpace(modcon))) Then
            MessageBox.Show("Ingrese los datos solicitados")
        Else

            If Not (String.IsNullOrEmpty(txtContraseñaUAM.Text) OrElse String.IsNullOrWhiteSpace(txtContraseñaUAM.Text)) AndAlso Not (txtContraseñaUAM.Text.Equals(modcon)) Then
                modcon = Encriptado(txtContraseñaUAM.Text)
            End If

            Dim mater As String = "NULL"
            Dim telef As String = "NULL"
            Dim a As String = "0"

            If chkUsuarioUAM.Checked Then
                a = "1"
            End If

            If modUsuario = 0 Then

                Try

                    If Not (String.IsNullOrEmpty(txtApellidoMatUAM.Text) OrElse String.IsNullOrWhiteSpace(txtApellidoMatUAM.Text)) Then
                        mater = txtApellidoMatUAM.Text
                    End If

                    If Not (String.IsNullOrEmpty(txtTelefonoUAM.Text) OrElse String.IsNullOrWhiteSpace(txtTelefonoUAM.Text)) Then
                        telef = txtTelefonoUAM.Text
                    End If

                    Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                        sql.Open()
                        cmd = New SqlCommand("CreateTipoCliente", sql)
                        cmd.CommandType = CommandType.StoredProcedure
                        cmd.Parameters.Add(New SqlParameter("@Nombre", txtNombreUAM.Text))
                        cmd.Parameters.Add(New SqlParameter("@Contraseña", modcon))
                        cmd.Parameters.Add(New SqlParameter("@Nombre_Usuario", txtNombreUsuarioUAM.Text))
                        cmd.Parameters.Add(New SqlParameter("@Apellido_Paterno", txtApellidoPatUAM.Text))
                        cmd.Parameters.Add(New SqlParameter("@Apellido_Materno", mater))
                        cmd.Parameters.Add(New SqlParameter("@Email", txtCorreoUAM.Text))
                        cmd.Parameters.Add(New SqlParameter("@Telefono", telef))
                        cmd.Parameters.Add(New SqlParameter("@Id_Puesto", cbPuestoUAM.SelectedValue))
                        cmd.Parameters.Add(New SqlParameter("@Id_Departamento", cbDepartamentoUAM.SelectedValue))
                        cmd.Parameters.Add(New SqlParameter("@Id_Tipo_usuario", cbTipoUsuarioUAM.SelectedValue))
                        cmd.Parameters.Add(New SqlParameter("@Id_Usuario_Creacion", IDuser))
                        cmd.Parameters.Add(New SqlParameter("@Fecha_Hora_Creacion", DateTime.Now))
                        cmd.Parameters.Add(New SqlParameter("@Id_Usuario_Modificacion", IDuser))
                        cmd.Parameters.Add(New SqlParameter("@Fecha_Hora_Modificacion", DateTime.Now))
                        cmd.Parameters.Add(New SqlParameter("@Estatus", a))
                        cmd.ExecuteReader()
                        LogsAceptable("Insertar Usuario", "Usuario")
                        sql.Close()
                        MessageBox.Show("Guardado")
                        panelUsuarioAM.Visible = False
                    End Using

                Catch ex As Exception
                    MessageBox.Show("Error de conexion")
                End Try
            Else

                Try

                    If Not (String.IsNullOrEmpty(txtApellidoMatUAM.Text) OrElse String.IsNullOrWhiteSpace(txtApellidoMatUAM.Text)) Then
                        mater = "'" & txtApellidoMatUAM.Text & "'"
                    End If

                    If Not (String.IsNullOrEmpty(txtTelefonoUAM.Text) OrElse String.IsNullOrWhiteSpace(txtTelefonoUAM.Text)) Then
                        telef = "'" & txtTelefonoUAM.Text & "'"
                    End If

                    Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                        sql.Open()
                        Dim consulta As String = "update Usuario set Nombre_Usuario='" & txtNombreUsuarioUAM.Text & "',Nombre='" + txtNombreUAM.Text & "'," & "Contraseña='" & modcon & "',Apellido_Paterno='" + txtApellidoPatUAM.Text & "',Apellido_Materno=" & mater & ",Email='" + txtCorreoUAM.Text & "',Telefono=" & telef & ",Id_Usuario_Modificacion=" & IDuser.ToString() & ",Fecha_Hora_Modificacion= '" & DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff") & "',Estatus=" & a & ",Id_Puesto=" + cbPuestoUAM.SelectedValue.ToString() & ",Id_Departamento=" + cbDepartamentoUAM.SelectedValue.ToString() & ",Id_Tipo_Usuario=" + cbTipoUsuarioUAM.SelectedValue.ToString() & " where Id_Usuario=" & modUsuario.ToString()
                        cmd = New SqlCommand(consulta, sql)
                        cmd.ExecuteReader()
                        LogsAceptable("Modificar Usuario", "Usuario")
                        sql.Close()
                        MessageBox.Show("Guardado")
                        panelUsuarioAM.Visible = False
                    End Using

                Catch ex As Exception
                    MessageBox.Show("Error en conexion")
                End Try
            End If

            recargaCatalogoUsuario()
        End If
    End Sub

    Private Sub btnVerVenta_Click(sender As Object, e As EventArgs) Handles btnVerVenta.Click
        If dgvVerVentas.SelectedRows.Count = 1 Then
            cargaDetalleVenta(CInt(dgvVerVentas.SelectedCells(0).Value))
            PanelDetalleVenta.Visible = True
            PanelDetalleVenta.Enabled = True
            PanelDetalleVenta.BringToFront()
        Else
            MessageBox.Show("Seleccione una Venta.")
        End If
    End Sub

    Private Sub btnEliminaVenta_Click(sender As Object, e As EventArgs) Handles btnEliminaVenta.Click
        If dgvVerVentas.SelectedRows.Count = 1 Then

            Try

                Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                    sql.Open()
                    cmd = New SqlCommand("DeleteVenta", sql)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.Parameters.Add(New SqlParameter("@Id", dgvVerVentas.SelectedCells(0).Value))
                    cmd.ExecuteReader()
                    LogsAceptable("Eliminar Venta", "Venta")
                    sql.Close()
                    MessageBox.Show("Eliminada")
                    recargaCatalogoVentas()
                End Using

            Catch ex As Exception
                MessageBox.Show("Error de conexion")
            End Try
        Else
            MessageBox.Show("Seleccione una Venta.")
        End If
    End Sub

    Private Sub button1_Click(sender As Object, e As EventArgs) Handles button1.Click
        If dgvCaja.SelectedRows.Count = 1 Then
            dgvCaja.Rows.Remove(dgvCaja.SelectedRows(0))
        Else
            MessageBox.Show("Seleccione un Libro.")
        End If
    End Sub

    Private Sub btnBuscaCaja_Click(sender As Object, e As EventArgs) Handles btnBuscaCaja.Click
        Dim existe As Boolean = False
        Dim id As String = cbLibrosCaja.SelectedValue.ToString()

        For i As Integer = 0 To dgvCaja.Rows.Count - 1

            If (dgvCaja.Rows(i).Cells("ISBN").Value.ToString()).Equals(id) Then
                existe = True
            End If
        Next

        If Not existe Then
            Dim ti As String = ""
            Dim pre As Decimal = 0

            Try

                Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                    sql.Open()
                    Dim consulta As String = "select l.Id_Libro_ISBN ISBN,CONCAT(l.Titulo,', ',f.Nombre) Titulo,l.Precio from Libro l left join Tipo_Formato f on l.Id_Tipo_Formato=f.Id_Tipo_Formato where l.Id_Libro_ISBN='" & id & "'"
                    cmd = New SqlCommand(consulta, sql)
                    Dim lector As SqlDataReader
                    lector = cmd.ExecuteReader()

                    If lector.HasRows Then

                        While lector.Read()
                            ti = CStr(lector(1))
                            pre = CDec(lector(2))
                        End While
                    End If
                End Using

            Catch ex As Exception
                MessageBox.Show("Error en conexion")
            End Try

            dgvCaja.Rows.Add(id, ti, pre, "0")
        End If
    End Sub

    Private Sub btnLimpiaCaja_Click(sender As Object, e As EventArgs) Handles btnLimpiaCaja.Click
        dgvCaja.Rows.Clear()
    End Sub

    Private Sub btnRegistraVentaCaja_Click(sender As Object, e As EventArgs) Handles btnRegistraVentaCaja.Click
        Try

            Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                sql.Open()
                cmd = New SqlCommand("CreateVenta", sql)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add(New SqlParameter("@Id_Tipo_Cliente", cbClienteCaja.SelectedValue))
                cmd.Parameters.Add(New SqlParameter("@Id_Tipo_Pago", cbPagoCaja.SelectedValue))
                cmd.Parameters.Add(New SqlParameter("@Id_Usuario_Creacion", IDuser))
                cmd.Parameters.Add(New SqlParameter("@Fecha_Hora_Creacion", DateTime.Now))
                cmd.Parameters.Add(New SqlParameter("@Id_Usuario_Modificacion", IDuser))
                cmd.Parameters.Add(New SqlParameter("@Fecha_Hora_Modificacion", DateTime.Now))
                cmd.Parameters.Add(New SqlParameter("@Estatus", "1"))
                cmd.ExecuteReader()
                sql.Close()
            End Using

            LogsAceptable("Insertar Venta", "Venta")

            Try

                Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                    sql.Open()
                    Dim consulta As String = "select top(1) id_Venta from Venta order by Id_Venta desc"
                    cmd = New SqlCommand(consulta, sql)
                    Dim lector As SqlDataReader
                    lector = cmd.ExecuteReader()

                    If lector.HasRows Then

                        While lector.Read()
                            modVenta = CInt(lector(0))
                        End While
                    End If
                End Using

            Catch ex As Exception
                MessageBox.Show("Error en conexion")
            End Try

            For i As Integer = 0 To dgvCaja.Rows.Count - 1

                If Not (dgvCaja.Rows(i).Cells("ISBN").Value.ToString()).Equals("") Then
                    Dim d As String = "0"

                    If String.IsNullOrEmpty(dgvCaja.Rows(i).Cells("Descuento").Value.ToString()) OrElse String.IsNullOrWhiteSpace(dgvCaja.Rows(i).Cells("Descuento").Value.ToString()) Then
                        d = "0"
                    Else
                        d = dgvCaja.Rows(i).Cells("Descuento").Value.ToString()
                    End If

                    agregarDetalleVentaLibros(dgvCaja.Rows(i).Cells("ISBN").Value.ToString(), d)
                    LogsAceptable("Insertar Detalle", "Detalle_Venta")
                End If
            Next

            MessageBox.Show("Guardado")
        Catch ex As Exception
            MessageBox.Show("Error en conexion")
        End Try
    End Sub

    Private Sub btnDetalleVentaRegresar_Click(sender As Object, e As EventArgs) Handles btnDetalleVentaRegresar.Click
        PanelDetalleVenta.Visible = False
    End Sub

    Private Sub agregarDetalleVentaLibros(ByVal isbn As String, ByVal descuento As String)
        Try

            Using sql As SqlConnection = New SqlConnection("Data Source=DESKTOP-CUOAPA9\SQLEXPRESS;Initial Catalog=libreria;Integrated Security=True")
                sql.Open()
                cmd = New SqlCommand("CreateDetalleVenta", sql)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add(New SqlParameter("@Id_Venta", modVenta))
                cmd.Parameters.Add(New SqlParameter("@Id_Descuento", descuento))
                cmd.Parameters.Add(New SqlParameter("@Id_Libro_ISBN", isbn))
                cmd.Parameters.Add(New SqlParameter("@Id_Usuario_Creacion", IDuser))
                cmd.Parameters.Add(New SqlParameter("@Fecha_Hora_Creacion", DateTime.Now))
                cmd.Parameters.Add(New SqlParameter("@Id_Usuario_Modificacion", IDuser))
                cmd.Parameters.Add(New SqlParameter("@Fecha_Hora_Modificacion", DateTime.Now))
                cmd.Parameters.Add(New SqlParameter("@Estatus", "1"))
                cmd.ExecuteReader()
                sql.Close()
            End Using

        Catch ex As Exception
            Throw
        End Try
    End Sub
End Class
