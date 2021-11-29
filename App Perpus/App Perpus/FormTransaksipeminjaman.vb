Imports System.Data.SqlClient


Public Class FormTransaksipeminjaman
    Sub KondisiAwal()
        Call NoOtomatis()
        Call munculKodeAnggota()
        LBLPetugas.Text = FormMenuUtama.STLabel4.Text
        LBLNama.Text = ""
        LBLAlamat.Text = ""
        LBLTelp.Text = ""
        LBLJudul.Text = ""
        LBLPengarang.Text = ""
        LBLTahun.Text = ""
        LBLTotalBuku.Text = "0"
        ComboBox1.Text = ""
        TextBox2.Enabled = False
        Call buatKolom()
    End Sub
    Private Sub FormTransaksipeminjaman_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call KondisiAwal()
        LBLTgl.Text = Today
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        LBLJam.Text = TimeOfDay
    End Sub
    Sub NoOtomatis()
        Call Koneksi()
        Cmd = New SqlCommand("Select * From TBL_PINJAM where NoPinjam in (select max(NoPinjam) From TBL_PINJAM)", Conn)
        Dim urutanKode As String
        Dim Hitung As Long
        Rd = Cmd.ExecuteReader
        Rd.Read()
        If Not Rd.HasRows Then
            urutanKode = "P" + Format(Now, "yyMMdd") + "001"
        Else
            Hitung = Microsoft.VisualBasic.Right(Rd.GetString(0), 9) + 1
            urutanKode = "P" + Format(Now, "yyMMdd") + Microsoft.VisualBasic.Right("000" & Hitung, 3)
        End If
        LBLNoPinjam.Text = urutanKode
    End Sub

    Sub munculKodeAnggota()
        Call Koneksi()
        ComboBox1.Items.Clear()
        Cmd = New SqlCommand("Select * From TBL_ANGGOTA", Conn)
        Rd = Cmd.ExecuteReader
        Do While Rd.Read
            ComboBox1.Items.Add(Rd.Item(0))
        Loop
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Call Koneksi()
        Cmd = New SqlCommand("Select * From TBL_ANGGOTA where KodeAnggota='" & ComboBox1.Text & "'", Conn)
        Rd = Cmd.ExecuteReader
        Rd.Read()
        If Rd.HasRows Then
            LBLNama.Text = Rd!NamaAnggota
            LBLAlamat.Text = Rd!AlamatAnggota
            LBLTelp.Text = Rd!TelpAnggota

            Call cekPeminjaman()
        End If
    End Sub

    Sub buatKolom()
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("Kode", "Kode Buku")
        DataGridView1.Columns.Add("Judul", "Judul Buku")
        DataGridView1.Columns.Add("Pengarang", "pengarang")
        DataGridView1.Columns.Add("Penerbit", "Penerbit")
        DataGridView1.Columns.Add("Tahun", "Tahun")
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then
            Call Koneksi()
            Cmd = New SqlCommand("Select * From TBL_BUKU where KodeBuku='" & TextBox1.Text & "'", Conn)
            Rd = Cmd.ExecuteReader
            Rd.Read()
            If Not Rd.HasRows Then
                MsgBox("Kode Buku Tidak Ada")
            Else
                TextBox1.Text = Rd.Item("KodeBuku")
                LBLJudul.Text = Rd.Item("JudulBuku")
                LBLPengarang.Text = Rd.Item("PengarangBuku")
                LBLTahun.Text = Rd.Item("TahunBuku")
                TextBox2.Enabled = True
                TextBox2.Text = "1"
            End If
        End If
    End Sub

    Sub rumusTotalBuku()
        Dim Hitung As Integer = 0
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            Hitung = Hitung + DataGridView1.Rows(i).Cells(4).Value
            LBLTotalBuku.Text = Hitung
        Next
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Val(LBLTelahPinjam.Text) + Val(LBLTotalBuku.Text) >= 5 Or Val(LBLTotalBuku.Text) + Val(TextBox2.Text) > 5 Then
            MsgBox("Buku yang di Pinjam Maksimal 5")
        Else
            DataGridView1.Rows.Add(New String() {TextBox1.Text, LBLJudul.Text, LBLPengarang.Text, LBLTahun.Text, TextBox2.Text})
            TextBox1.Text = ""
            TextBox2.Text = ""
            LBLJudul.Text = ""
            LBLPengarang.Text = ""
            LBLTahun.Text = ""
            Call rumusTotalBuku()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If LBLNama.Text = "" Or Label9.Text = "" Then
            MsgBox("Transaksi Tidak Ada, Silahkan Transaksi Lebih Dahulu")
        Else
            Call Koneksi()
            Dim tglsql As String
            tglsql = Format(Today, "yyyy-MM-dd")
            Dim pinjamBuku As String = "Insert into tbl_pinjam values('" & LBLNoPinjam.Text & "', '" & tglsql & "', '" & LBLJam.Text & "', '" & ComboBox1.Text & "', '" & LBLTotalBuku.Text & "', '" & LBLTotalBuku.Text & "', '" & FormMenuUtama.STLabel2.Text & "')"
            Cmd = New SqlCommand(pinjamBuku, Conn)
            Cmd.ExecuteNonQuery()
            For baris As Integer = 0 To DataGridView1.Rows.Count - 2
                Call Koneksi()
                Dim simpanDetail As String = "Insert into tbl_detailpinjam values ('" & LBLNoPinjam.Text & "', '" & DataGridView1.Rows(baris).Cells(0).Value & "', '" & DataGridView1.Rows(baris).Cells(4).Value & "')"
                Cmd = New SqlCommand(simpanDetail, Conn)
                Cmd.ExecuteNonQuery()
                Call Koneksi()
                Cmd = New SqlCommand("Select * From tbl_buku where KodeBuku='" & DataGridView1.Rows(baris).Cells(0).Value & "'", Conn)
                Rd = Cmd.ExecuteReader
                Rd.Read()
                Call Koneksi()
                Dim kurangiStock As String = "Update tbl_buku set StockBuku= '" & (Rd.Item("StockBuku") - DataGridView1.Rows(baris).Cells(4).Value) & "' where KodeBuku= '" & DataGridView1.Rows(baris).Cells(0).Value & "'"
                Cmd = New SqlCommand(kurangiStock, Conn)
                Cmd.ExecuteNonQuery()
            Next
            Call KondisiAwal()
            MsgBox("Transaksi Telah Berhasil disimpan")
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Call KondisiAwal()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Sub cekPeminjaman()
        Call Koneksi()
        Da = New SqlDataAdapter("Select tbl_buku.KodeBuku,JudulBuku,JumlahBuku From TBL_Anggota,TBL_Pinjam,TBL_Buku,TBL_Detailpinjam where TBL_Buku.KodeBuku=TBL_Detailpinjam.KodeBuku And TBL_Pinjam.NoPinjam=TBL_Detailpinjam.NoPinjam And TBL_Anggota.KodeAnggota=TBL_Pinjam.KodeAnggota And TBL_Anggota.KodeAnggota='" & ComboBox1.Text & "' AND TBL_Detailpinjam.JumlahBuku>0", Conn)
        Ds = New DataSet
        Ds.Clear()
        Da.Fill(Ds, "Detail")
        DataGridView2.DataSource = Ds.Tables("Detail")
        LBLTelahPinjam.Text = DataGridView2.Rows.Count - 1
        DataGridView2.ReadOnly = True
    End Sub
End Class