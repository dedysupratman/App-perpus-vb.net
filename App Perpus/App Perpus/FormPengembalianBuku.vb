Imports System.Data.SqlClient

Public Class FormPengembalianBuku
    Private Sub KondisiAwal()
        TXTKODEAGT.Text = ""
        TextBox1.Text = ""
        TextBox2.Text = ""
        LBLJudul.Text = ""
        TextBox2.Text = ""
        LBLTglPinjam.Text = ""
        LBLLamaPjm.Text = ""
        LBLDenda.Text = ""
        LBLNoPjm.Text = ""
        LBLJam.Text = ""
        LBLPetugas.Text = ""
    End Sub
    Private Sub KondisiAwal2()
        TextBox1.Text = ""
        TextBox2.Text = ""
        LBLJudul.Text = ""
        TextBox2.Text = ""
        LBLTglPinjam.Text = ""
        LBLLamaPjm.Text = ""
        LBLDenda.Text = ""
        LBLNoPjm.Text = ""
    End Sub
    Sub Tampilgrid1()
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("Kode", "Kode Buku")
        DataGridView1.Columns.Add("Nomor", "Nomor Pinjam")
        DataGridView1.Columns.Add("Judul", "Judul Buku")
        DataGridView1.Columns.Add("Jumlah", "Jumlah Pinjam")
        DataGridView1.Columns.Add("Tanggal", "Tgl Pinjam")
        DataGridView1.Columns.Add("Lama", "Lama Pinjam")
        DataGridView1.Columns.Add("Denda", "Denda")
        DataGridView1.Columns(0).Width = 70
        DataGridView1.Columns(1).Width = 75
        DataGridView1.Columns(2).Width = 150
        DataGridView1.Columns(3).Width = 50
        DataGridView1.Columns(4).Width = 75
        DataGridView1.Columns(5).Width = 50
        DataGridView1.Columns(6).Width = 75
    End Sub
    Sub Tampilgrid2()
        DataGridView2.Columns.Add("Kode", "Kode")
        DataGridView2.Columns.Add("Judul", "Judul Buku")
        DataGridView2.Columns.Add("Jumlah", "Jumlah Buku")
        DataGridView2.Columns(0).Width = 75
        DataGridView2.Columns(1).Width = 300
        DataGridView2.Columns(2).Width = 100
        DataGridView2.ReadOnly = True
    End Sub
    Sub Otomatis()
        Call Koneksi()
        Cmd = New SqlCommand("Select * From TBL_KEMBALI where NoKembali in (select max(NoKembali) From TBL_KEMBALI)", Conn)
        Dim urutanKode As String
        Dim Hitung As Long
        Rd = Cmd.ExecuteReader
        Rd.Read()
        If Not Rd.HasRows Then
            urutanKode = "KBL" + Format(Now, "yyMMdd") + "001"
        Else
            Hitung = Microsoft.VisualBasic.Right(Rd.GetString(0), 9) + 1
            urutanKode = "KBL" + Format(Now, "yyMMdd") + Microsoft.VisualBasic.Right("000" & Hitung, 3)
        End If
        LBLNoKbl.Text = urutanKode
    End Sub

    Private Sub FormPengembalianBuku_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        Koneksi()
        Call Otomatis()
        LBLTglKbl.Text = Today
    End Sub

    Private Sub FormPengembalianBuku_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call KondisiAwal()
        Call Koneksi()
        Call Tampilgrid1()
        Call Otomatis()
        LBLPetugas.Text = FormMenuUtama.STLabel4.Text
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        LBLJam.Text = TimeOfDay
    End Sub
End Class