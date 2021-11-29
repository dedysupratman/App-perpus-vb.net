Imports System.Data.SqlClient


Public Class FormMasterBuku
    Sub KondisiAwal()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        ComboBox1.Text = ""

        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        ComboBox1.Enabled = False
        TextBox5.Enabled = False
        TextBox6.Enabled = False
        Button1.Text = "INPUT"
        Button2.Text = "EDIT"
        Button3.Text = "HAPUS"
        Button4.Text = "TUTUP"
        Button1.Enabled = True
        Button2.Enabled = True
        Button3.Enabled = True
        Button4.Enabled = True


        Dim Tahun As String
        Tahun = DateTime.Now().Year
        ComboBox1.Items.Clear()
        With ComboBox1
            While Tahun >= 2014
                .Items.Add(Tahun)
                Tahun = Tahun - 1
            End While
        End With
        Call Koneksi()
        Da = New SqlDataAdapter("Select * From TBL_BUKU", Conn)
        Ds = New DataSet
        Da.Fill(Ds, "TBL_BUKU")
        DataGridView1.DataSource = (Ds.Tables("TBL_BUKU"))
    End Sub

    Sub SiapIsi()
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        ComboBox1.Enabled = True
        TextBox5.Enabled = True
        TextBox6.Enabled = True
    End Sub

    Private Sub FormMasterBuku_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call KondisiAwal()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Button1.Text = "INPUT" Then
            Button1.Text = "SIMPAN"
            Button2.Enabled = False
            Button3.Enabled = False
            Button4.Text = "BATAL"

            Call SiapIsi()
        Else
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or ComboBox1.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Then
                MsgBox("Field tidak boleh kosong")
            Else
                Call Koneksi()
                Dim SimpanData As String = "insert into TBL_BUKU values('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox6.Text & "','" & ComboBox1.Text & "','" & TextBox5.Text & "')"
                Cmd = New SqlCommand(SimpanData, Conn)
                Cmd.ExecuteNonQuery()
                MsgBox("Data Berhasil disimpan")
                Call KondisiAwal()
            End If
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Button2.Text = "EDIT" Then
            Button2.Text = "SIMPAN"
            Button1.Enabled = False
            Button3.Enabled = False
            Button4.Text = "BATAL"
            Call SiapIsi()
        Else
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or ComboBox1.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Then
                MsgBox("Field tidak boleh kosong")
            Else
                Call Koneksi()
                Dim EditData As String = "update TBL_BUKU set JudulBuku='" & TextBox2.Text & "', PengarangBuku='" & TextBox3.Text & "', PenerbitBuku='" & TextBox6.Text & "', TahunBuku='" & ComboBox1.Text & "',StockBuku='" & TextBox5.Text & "' where KodeBuku ='" & TextBox1.Text & "'"
                Cmd = New SqlCommand(EditData, Conn)
                Cmd.ExecuteNonQuery()
                MsgBox("Data Berhasil diedit")
                Call KondisiAwal()
            End If
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Button3.Text = "HAPUS" Then
            Button3.Text = "DELETE"
            Button1.Enabled = False
            Button2.Enabled = False
            Button4.Text = "BATAL"
            Call SiapIsi()
        Else
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or ComboBox1.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Then
                MsgBox("Field tidak boleh kosong")
            Else
                Call Koneksi()
                Dim HapusData As String = "delete from TBL_BUKU where KodeBuku='" & TextBox1.Text & "'"
                Cmd = New SqlCommand(HapusData, Conn)
                Cmd.ExecuteNonQuery()
                MsgBox("Data Berhasil di Hapus")
                Call KondisiAwal()
            End If
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If Button4.Text = "TUTUP" Then
            Me.Close()
        Else
            Call KondisiAwal()
        End If
    End Sub
    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then
            Call Koneksi()
            Cmd = New SqlCommand("select * from TBL_BUKU where KodeBuku='" & TextBox1.Text & "'", Conn)
            Rd = Cmd.ExecuteReader
            Rd.Read()
            If Rd.HasRows Then
                TextBox2.Text = Rd.Item("JudulBuku")
                TextBox3.Text = Rd.Item("PengarangBuku")
                TextBox6.Text = Rd.Item("PenerbitBuku")
                ComboBox1.Text = Rd.Item("TahunBuku")
                TextBox5.Text = Rd.Item("StockBuku")
            Else
                MsgBox("Data tidak ada")
            End If
        End If
    End Sub
End Class