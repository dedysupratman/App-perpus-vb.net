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
        LBLNoPjm.Text = ""
        LBLJam.Text = ""
        LBLPetugas.Text = ""
        LBLTelahPinjam.Text = ""
        LBLTotal.Text = ""
        LBLDenda.Text = ""
        LBLNamaAgt.Text = ""
    End Sub
    Private Sub KondisiAwal2()
        TextBox1.Text = ""
        TextBox2.Text = ""
        LBLJudul.Text = ""
        TextBox2.Text = ""
        LBLTglPinjam.Text = ""
        LBLLamaPjm.Text = ""
        LBL_Denda.Text = ""
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

    Sub pinjaman()
        Call Koneksi()
        Da = New SqlDataAdapter("Select TBL_BUKU.KodeBuku, JudulBuku, JumlahBuku From TBL_Anggota, TBL_Pinjam,TBL_Buku,TBL_DetailPinjam where TBL_Buku.KodeBuku=TBL_DetailPinjam.KodeBuku and TBL_Pinjam.NoPinjam=TBL_DetailPinjam.NoPinjam and TBL_Anggota.KodeAnggota=TBL_Pinjam.KodeAnggota and TBL_Anggota.KodeAnggota= '" & TXTKODEAGT.Text & "' and TBL_DetailPinjam.JumlahBuku >0", Conn)
        Ds = New DataSet
        Ds.Clear()
        Da.Fill(Ds, "Detail")
        DataGridView2.DataSource = Ds.Tables("Detail")
        DataGridView2.ReadOnly = True
        DataGridView2.Columns(0).Width = 80
        DataGridView2.Columns(0).Width = 300
        DataGridView2.Columns(0).Width = 100
    End Sub

    Sub CariData()
        Call Koneksi()
        Cmd = New SqlCommand("Select * From TBL_Anggota where KodeAnggota= '" & TXTKODEAGT.Text & "'", Conn)
        Rd = Cmd.ExecuteReader
        Rd.Read()
        If Rd.HasRows Then
            LBLNamaAgt.Text = Rd.Item("NamaAnggota")
            Call Koneksi()
            Cmd = New SqlCommand("Select KodeAnggota From TBL_Pinjam where KodeAnggota ='" & TXTKODEAGT.Text & "'", Conn)
            Rd = Cmd.ExecuteReader
            Rd.Read()
            If Rd.HasRows Then
                Call Koneksi()
                Cmd = New SqlCommand("Select SUM(TotalPinjam) as Ketemu from TBL_Pinjam where KodeAnggota='" & TXTKODEAGT.Text & "'", Conn)
                Rd = Cmd.ExecuteReader
                Rd.Read()
                LBLTelahPinjam.Text = Rd.Item(0)
                Call pinjaman()
            End If
        End If
    End Sub

    Private Sub TXTKODEAGT_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TXTKODEAGT.KeyPress
        If e.KeyChar = Chr(13) Then
            Call CariData()
        End If
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then
            Call Koneksi()
            Cmd = New SqlCommand("Select distinct TBL_Buku.KodeBuku, TBL_DetailPinjam.NoPinjam,JudulBuku,JumlahBuku,TglPinjam From TBL_Anggota,TBL_Pinjam,TBL_Buku,TBL_DetailPinjam where TBL_Buku.KOdeBuku=TBL_DetailPinjam.KodeBuku and TBL_Pinjam.NoPinjam= TBL_DetailPinjam.NoPinjam and TBL_Anggota.KodeAnggota = TBL_Pinjam.KodeAnggota and TBL_Anggota.KodeAnggota= '" & TXTKODEAGT.Text & "' and TBL_DetailPinjam.KodeBuku = '" & TextBox1.Text & "' and TBL_DetailPinjam.JumlahBuku>0", Conn)
            Rd = Cmd.ExecuteReader
            Rd.Read()
            If Not Rd.HasRows Then
                MsgBox("Kode Buku Yang di Masukan Salah")
            Else
                LBLNoPjm.Text = Rd.Item(1)
                LBLTglPinjam.Text = Rd.Item(4)
                LBLJudul.Text = Rd.Item(2)
                TextBox2.Text = Rd.Item(3)
                LBLLamaPjm.Text = DateDiff(DateInterval.Day, Rd.Item(4), Today())
                If LBLLamaPjm.Text > 5 Then
                    LBLDenda.Text = Val((LBLLamaPjm.Text) - 5) * 2000
                Else
                    LBLDenda.Text = 0
                End If
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DataGridView1.Rows.Add(New String() {TextBox1.Text, LBLJudul.Text, LBLNoPjm.Text, TextBox2.Text, LBLTglPinjam.Text, LBLLamaPjm.Text, LBLDenda.Text})
        Call HapusMaster()
        Call TotalDenda()
        Call TotalKembali()
    End Sub

    Sub HapusMaster()
        TextBox1.Text = ""
        LBLNoPjm.Text = ""
        LBLTglPinjam.Text = ""
        LBLJudul.Text = ""
        TextBox2.Text = ""
        LBLLamaPjm.Text = ""
    End Sub

    Sub TotalDenda()
        Dim subtot As Integer = 0
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            subtot += Val(DataGridView1.Rows(1).Cells(6).Value)
            LBL_Denda.Text = subtot
        Next
    End Sub
    Sub TotalKembali()
        Dim subtot As Integer = 0
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            subtot += Val(DataGridView1.Rows(1).Cells(3).Value)
            LBLTotal.Text = subtot
        Next
    End Sub

    Private Sub TXTDIBAYAR_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TXTDIBAYAR.KeyPress
        If e.KeyChar = Chr(13) Then
            If Val(TXTDIBAYAR.Text) < Val(LBL_Denda.Text) Then
                MsgBox("Pembayaran Kurang")
                TXTDIBAYAR.Focus()
            ElseIf Val(TXTDIBAYAR.Text) = Val(LBL_Denda.Text) Then
                LBLKembali.Text = 0
                Button4.Focus()
            Else
                LBLKembali.Text = Val(LBL_Denda.Text) - Val(TXTDIBAYAR.Text)
                Button4.Focus()
            End If
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If TXTKODEAGT.Text = "" Then
            MsgBox("Transaksi Tidak ada, Silahkan transaksi terlebih dahulu")
        Else
            Call Koneksi()
            Dim tglsql As String
            tglsql = Format(Today, "yyyy-MM-dd")
            Dim BukuKembali As String = "Insert into TBL_kembali values('" & LBLNoKbl.Text & "', '" & tglsql & "', '" & LBLTotal.Text & "', '" & LBLDenda.Text & "', '" & TXTDIBAYAR.Text & "', '" & LBLKembali.Text & "','" & TXTKODEAGT.Text & "', '" & FormMenuUtama.STLabel2.Text & "')"
            Cmd = New SqlCommand(BukuKembali, Conn)
            Cmd.ExecuteNonQuery()
            For baris As Integer = 0 To DataGridView1.Rows.Count - 2
                Call Koneksi()
                Dim simpanDetail As String = "Insert into tbl_detailkembali values ('" & LBLNoKbl.Text & "', '" & DataGridView1.Rows(baris).Cells(0).Value & "', '" & DataGridView1.Rows(baris).Cells(3).Value & "')"
                Cmd = New SqlCommand(simpanDetail, Conn)
                Cmd.ExecuteNonQuery()

                'update buku
                Call Koneksi()
                Cmd = New SqlCommand("Select * From tbl_buku where KodeBuku='" & DataGridView1.Rows(baris).Cells(0).Value & "'", Conn)
                Rd = Cmd.ExecuteReader
                Rd.Read()
                Call Koneksi()
                Dim kurangiStock As String = "Update tbl_buku set StockBuku= '" & (Rd.Item("StockBuku") + DataGridView1.Rows(baris).Cells(3).Value) & "' where KodeBuku= '" & DataGridView1.Rows(baris).Cells(0).Value & "'"
                Cmd = New SqlCommand(kurangiStock, Conn)
                Cmd.ExecuteNonQuery()

                'update pinjam buku
                Call Koneksi()
                Cmd = New SqlCommand("Select jumlahbuku from TBL_detailpinjam where KodeBuku='" & DataGridView1.Rows(baris).Cells(0).Value & "' and NoPinjam='" & DataGridView1.Rows(baris).Cells(1).Value & "'", Conn)
                Rd = Cmd.ExecuteReader
                Rd.Read()
                If Rd.HasRows Then
                    Call Koneksi()
                    Dim updateDetailPinjam As String = "Update TBL_detailpinjam set JumlahBuku= '" & (Rd.Item(0) - DataGridView1.Rows(baris).Cells(3).Value) & "' where KodeBuku= '" & DataGridView1.Rows(baris).Cells(0).Value & "' and NoPinjam='" & DataGridView1.Rows(baris).Cells(1).Value & "'"
                    Cmd = New SqlCommand(updateDetailPinjam, Conn)
                    Cmd.ExecuteNonQuery()
                End If

                'pdate total pinjam
                Call Koneksi()
                Cmd = New SqlCommand("Select totalpinjam from TBL_pinjam where NoPinjam='" & DataGridView1.Rows(baris).Cells(1).Value & "'", Conn)
                Rd = Cmd.ExecuteReader
                Rd.Read()
                If Rd.HasRows Then
                    Call Koneksi()
                    Dim sqleditpjm As String = "Update TBL_pinjam set totalpinjam= '" & (Rd.Item(0) - DataGridView1.Rows(baris).Cells(3).Value) & "' where NoPinjam= '" & DataGridView1.Rows(baris).Cells(1).Value & "'"
                    Cmd = New SqlCommand(sqleditpjm, Conn)
                    Cmd.ExecuteNonQuery()
                End If

            Next
            Call KondisiAwal()
                    Call Otomatis()
                    DataGridView1.Columns.Clear()
                    DataGridView2.Columns.Clear()
                    MsgBox("Transaksi Telah Berhasil disimpan")
                    Call Tampilgrid1()
                End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call KondisiAwal()
    End Sub

End Class
