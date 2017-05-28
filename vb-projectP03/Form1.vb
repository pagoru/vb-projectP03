Imports System.Data.SqlClient

Public Class Form1

    Private conn As SqlConnection
    Private dts As DataSet
    Private adapter As SqlDataAdapter
    Private bmb As BindingManagerBase

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        EmptyMember(True)

        conn = New SqlConnection()
        conn.ConnectionString = "Data Source=.\SQLEXPRESS;Initial Catalog=COMPASSTRAVEL;Trusted_Connection=true;"
        conn.Open()

        adapter = New SqlDataAdapter("SELECT * FROM members ORDER BY memberid ASC", conn)
        Dim builder As SqlCommandBuilder = New SqlCommandBuilder(adapter)

        dts = New DataSet()
        adapter.Fill(dts, "Members")

        FillDataBindings()

    End Sub

    Private Sub FillDataBindings()
        Dim textboxes As TextBox() = {memberid, firstname, lastname, address1, address2, state, postalcode, country, username, password}
        Dim rows As String() = {"MEMBERID", "FIRSTNAME", "LASTNAME", "ADDRESS1", "ADDRESS2", "STATE", "POSTALCODE", "COUNTRY", "USERNAME", "PASSWORD"}

        For idx As Integer = 0 To textboxes.Length - 1 Step 1
            textboxes(idx).DataBindings.Add(New Binding("Text", dts, $"Members.{rows(idx)}"))
        Next
        bmb = BindingContext(dts, "Members")

        If bmb.Count > 0 Then
            EmptyMember(False)
        End If
        updateControllers()
    End Sub

    Private Sub Leftcontrols(opt As Boolean)
        primer.Enabled = opt
        anterior.Enabled = opt
    End Sub

    Private Sub RightControls(opt As Boolean)
        ultim.Enabled = opt
        seguent.Enabled = opt
    End Sub

    Private Sub ToggleNewMembre()
        cancelar.Visible = Not cancelar.Visible
        acceptar.Visible = Not acceptar.Visible

        eliminar.Visible = Not eliminar.Visible
        editar.Visible = Not editar.Visible
        afegir.Visible = Not afegir.Visible
        primer.Visible = Not primer.Visible
        anterior.Visible = Not anterior.Visible
        seguent.Visible = Not seguent.Visible
        ultim.Visible = Not ultim.Visible
        buscar.Visible = Not buscar.Visible
        buscar_box.Visible = Not buscar_box.Visible
    End Sub

    Private Sub afegir_Click(sender As Object, e As EventArgs) Handles afegir.Click
        ToggleNewMembre()
        GenerateNewID()

        EmptyMember(False)
    End Sub

    Private Sub cancelar_Click(sender As Object, e As EventArgs) Handles cancelar.Click
        ToggleNewMembre()

        bmb.CancelCurrentEdit()

        If bmb.Count = 0 Then
            EmptyMember(True)
        End If
    End Sub

    Private Sub acceptar_Click(sender As Object, e As EventArgs) Handles acceptar.Click

        If Not VerifyMember() Then
            Return
        End If

        ToggleNewMembre()

        bmb.EndCurrentEdit()
        SaveChanges()
        updateControllers()

    End Sub

    Private Function VerifyMember()

        If password.Text.Length < 4 Then
            MsgBox("La contrasenya ha de tenir mes de 4 caracters")
            Return False
        End If

        If password.Text.Length > 4 And password.Text.Length < 8 Then
            MsgBox("La contrasenya es valida pero hauria de ser de mes de 9 caracters")
        End If

        If username.Text = 0 Then
            MsgBox("S'ha d'escriure un nom d'usuari")
            Return False
        End If

        If UsernameExists(username.Text) Then
            MsgBox("Aquest nom d'usuari ja existeix")
            Return False
        End If

        If postalcode.Text.Length > 0 Then
            Try
                Integer.Parse(postalcode.Text)
            Catch ex As Exception
                MsgBox("El codi postal ha de ser un número valid")
                Return False
            End Try
        End If

        Return True
    End Function

    Private Function UsernameExists(username As String)
        For idx As Integer = 0 To dts.Tables("Members").Rows.Count - 1
            If Not idx = bmb.Position Then
                If username = dts.Tables("Members").Rows(idx)("USERNAME").ToString() Then
                    Return True
                End If
            End If
        Next
        Return False
    End Function

    Private Sub EmptyMember(bool As Boolean)
        firstname.ReadOnly = bool
        lastname.ReadOnly = bool
        address1.ReadOnly = bool
        address2.ReadOnly = bool
        state.ReadOnly = bool
        postalcode.ReadOnly = bool
        country.ReadOnly = bool
        username.ReadOnly = bool
        password.ReadOnly = bool
        buscar_box.ReadOnly = bool

        eliminar.Enabled = Not bool
        editar.Enabled = Not bool
        buscar.Enabled = Not bool
    End Sub

    Private Sub GenerateNewID()
        bmb.AddNew()
        Try
            memberid.Text = dts.Tables("Members").Rows(dts.Tables("Members").Rows.Count - 1)("MEMBERID") + 1
        Catch ex As Exception
            memberid.Text = "1"
        End Try

    End Sub

    Private Sub SaveChanges()
        Try
            adapter.Update(dts.Tables("Members").GetChanges())
            dts.Tables("Members").AcceptChanges()
            MsgBox("S'ha realitzat la persistencia amb exit")
        Catch ex As Exception
            MsgBox("S'ha realitzat la persistencia amb exit")
        End Try
    End Sub

    Private Sub primer_Click(sender As Object, e As EventArgs) Handles primer.Click
        bmb.Position = 0
        updateControllers()
    End Sub

    Private Sub ultim_Click(sender As Object, e As EventArgs) Handles ultim.Click
        bmb.Position = bmb.Count - 1
        updateControllers()
    End Sub

    Private Sub anterior_Click(sender As Object, e As EventArgs) Handles anterior.Click
        bmb.Position = bmb.Position - 1
        updateControllers()
    End Sub

    Private Sub seguent_Click(sender As Object, e As EventArgs) Handles seguent.Click
        bmb.Position = bmb.Position + 1
        updateControllers()
    End Sub

    Private Sub memberid_KeyDown(sender As Object, e As KeyEventArgs) Handles memberid.KeyDown

        If e.Alt Then
            If e.KeyCode = Keys.Up Then
                bmb.Position = 0
            ElseIf e.KeyCode = Keys.Down Then
                If bmb.Count > 0 Then
                    bmb.Position = bmb.Count - 1
                End If
            ElseIf e.KeyCode = Keys.Left Then
                If bmb.Position > 0 Then
                    bmb.Position = bmb.Position - 1
                End If
            ElseIf e.KeyCode = Keys.Right Then
                If bmb.Position < bmb.Count - 1 Then
                    bmb.Position = bmb.Position + 1
                End If
            End If
            updateControllers()
        End If

    End Sub

    Private Sub updateControllers()
        Leftcontrols(False)
        RightControls(False)

        If bmb.Count = 0 Then
            Return
        End If

        If Not bmb.Position = bmb.Count - 1 Then
            RightControls(True)
        End If

        If bmb.Position > 0 Then
            Leftcontrols(True)
        End If

    End Sub

    Private Sub genusername_Click(sender As Object, e As EventArgs) Handles genusername.Click
        If firstname.Text.Length = 0 Or lastname.Text.Length = 0 Then
            MsgBox("Per poder crear un username es necesari emplenar les dades de firstname i lastname")
            Return
        End If

        If firstname.Text.Length < 3 Or lastname.Text.Length < 3 Then
            MsgBox("Per poder crear un username es necesari que firstname i lastname tinguin mes de 3 caracters")
            Return
        End If

        username.Text = firstname.Text.Substring(0, 3) + lastname.Text.Substring(lastname.Text.Length - 3, 3)
    End Sub

    Private Sub eliminar_Click(sender As Object, e As EventArgs) Handles eliminar.Click
        If memberid.Text.Length > 0 Then
            If MessageBox.Show("Estas segur que vols eliminar el membre?", "Atenció", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                bmb.RemoveAt(bmb.Position)
                SaveChanges()
                updateControllers()

                If bmb.Count = 0 Then
                    EmptyMember(True)
                End If
            End If
        End If
    End Sub

    Private Sub buscar_Click(sender As Object, e As EventArgs) Handles buscar.Click
        For idx As Integer = 0 To dts.Tables("Members").Rows.Count - 1
            If buscar_box.Text = dts.Tables("Members").Rows(idx)("MEMBERID").ToString() Then
                bmb.Position = idx
                updateControllers()
                Return
            End If
        Next

    End Sub

    Private Sub editar_Click(sender As Object, e As EventArgs) Handles editar.Click
        If Not VerifyMember() Then
            Return
        End If

        bmb.EndCurrentEdit()
        SaveChanges()
        updateControllers()
    End Sub
End Class
