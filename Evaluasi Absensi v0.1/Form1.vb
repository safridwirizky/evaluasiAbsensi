Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Access.Dao
Imports System.Data.OleDb

Public Class Form1
    Dim ConnA, ConnXl As OleDbConnection
    Dim cmd As OleDbCommand
    Dim DRA As OleDbDataReader
    Dim DA As OleDbDataAdapter
    Dim DT As DataTable
    Dim engine As New DBEngine
    Dim db As Database
    Dim tdf As TableDef
    Dim fld As Field
    Dim prop As Access.Dao.Property
    Dim xlApp As New Excel.Application
    Dim xlWorkBook As Excel.Workbook
    Dim xlWorkSheet As Excel.Worksheet
    Dim rang As Excel.Range
    Dim tbl As String
    Public abc As String

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Button2.Enabled = False
        Button3.Enabled = False

        'CREATE FOLDER Evaluasi Absensi
        If Not IO.Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\Evaluasi Absensi\") Then
            IO.Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\Evaluasi Absensi\")
        End If

        'CREATE Database Evaluasi Absensi.accdb
        If Not IO.File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\Evaluasi Absensi\Database Evaluasi Absensi.accdb") Then
            db = engine.CreateDatabase(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\Evaluasi Absensi\Database Evaluasi Absensi.accdb", LanguageConstants.dbLangGeneral, DatabaseTypeEnum.dbVersion120)
            db.Close()
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim sheetName As New List(Of String)
        Dim tgl As Date
        Dim NIP, Nama, Bagian As String

        Dialog1.ListBox1.DataSource = Nothing
        DGV.DataSource = Nothing
        Button2.Enabled = False

        OpenFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        OpenFileDialog1.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls"

        If OpenFileDialog1.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
            Try
                xlWorkBook = xlApp.Workbooks.Open(OpenFileDialog1.FileName)

                For Each sht As Excel.Worksheet In xlWorkBook.Sheets
                    sheetName.Add(sht.Name)
                Next

                Dialog1.ListBox1.DataSource = sheetName

                If Dialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                    ConnXl = New OleDbConnection("provider=microsoft.ace.oledb.12.0;data source=" & OpenFileDialog1.FileName & ";Extended Properties=Excel 12.0")
                    DA = New OleDbDataAdapter("SELECT * FROM [" & abc & "$]", ConnXl)
                    DT = New DataTable
                    DA.Fill(DT)
                    DGV.DataSource = DT
                    DGV.ReadOnly = True

                    xlWorkBook.Close(False, OpenFileDialog1.FileName)
                Else
                    xlWorkBook.Close(False, OpenFileDialog1.FileName)
                    Exit Sub
                End If
            Catch ex As Exception
                xlWorkBook.Close(False)
                MsgBox("Kill all EXCEL.EXE processes in Task Manager! or repair your Office")
                Dialog1.Close()
                Me.Close()
                Exit Sub
            End Try

            tgl = DGV.Rows(10).Cells(3).Value.ToString
            tbl = Format(tgl, "MM_yyyy")

            ' Mengecek TABLE 'tbl' Sudah Ada Atau Belum
            ConnA = New OleDbConnection("provider=microsoft.ace.oledb.12.0;data source=" & Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\Evaluasi Absensi\Database Evaluasi Absensi.accdb;Jet Oledb:System Database=" & Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\microsoft\Access\System.mdw")
            cmd = New OleDbCommand("SELECT Name FROM MSysObjects WHERE Name='" & tbl & "'", ConnA)
            ConnA.Open()
            DRA = cmd.ExecuteReader

            If Not DRA.HasRows Then
                DRA.Close()
                ConnA.Close()

                ' Membuat TABLE 'tbl'(bulan evaluasi excel) Beserta FIELD-nya
                ConnA = New OleDbConnection("provider=microsoft.ace.oledb.12.0;data source=" & Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\Evaluasi Absensi\Database Evaluasi Absensi.accdb")
                cmd = New OleDbCommand("CREATE TABLE " & tbl & " ([NO] counter, NIP text(225), NAMA_PEGAWAI text(225), BAGIAN text(225), " & _
                                       "Kehadiran_Normal byte Default 0, SPPD byte Default 0, CUTI_atau_IJIN byte Default 0, SAKIT byte Default 0, " & _
                                       "TIDAK_ABSEN_atau_KOSONG byte Default 0, TIDAK_ABSEN_DATANG byte Default 0, TIDAK_ABSEN_PULANG byte Default 0, " & _
                                       "JUMLAH_HARI_KERJA byte Default 0, Prosentase_Disiplin_Absen text(225), Jumlah_terlambat byte Default 0)", ConnA)
                ConnA.Open()
                cmd.ExecuteNonQuery()
                ConnA.Close()

                ' MENAMBAH DATA PEGAWAI
                ConnA = New OleDbConnection("provider=microsoft.ace.oledb.12.0;data source=" & Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\Evaluasi Absensi\Database Evaluasi Absensi.accdb")
                ConnA.Open()
                For a As Integer = 3 To DGV.RowCount - 2
                    Bagian = DGV.Rows(a).Cells(6).Value.ToString
                    Nama = DGV.Rows(a).Cells(2).Value.ToString
                    NIP = DGV.Rows(a).Cells(1).Value.ToString

                    cmd = New OleDbCommand("SELECT NAMA_PEGAWAI FROM " & tbl & " WHERE NAMA_PEGAWAI='" & Nama & "'", ConnA)
                    DRA = cmd.ExecuteReader
                    If Not DRA.HasRows Then
                        cmd = New OleDbCommand("INSERT INTO " & tbl & "(NIP, NAMA_PEGAWAI, BAGIAN) Values('" & NIP & "', '" & Nama & "', '" & Bagian & "')", ConnA)
                        cmd.ExecuteNonQuery()
                    End If
                    DRA.Close()
                Next
                ConnA.Close()

            Else
                DRA.Close()
                ConnA.Close()
            End If

            Button2.Enabled = True
            Button3.Enabled = True
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim jml_hri_krja, jml_sppd, jml_cuti, jml_sakit, jml_kosong, jml_tdk_datang, _
            jml_tdk_pulang, jml_khdiran_normal, tlt As Byte
        Dim prosentase_dspln_absn As Single
        Dim jam As DateTime
        Dim nam_kary, nama_karyawan, jumlah_hari_kerja, jumlah_sppd, jumlah_cuti, jumlah_sakit, _
            jumlah_kosong, jumlah_tidak_datang, jumlah_tidak_pulang, jumlah_kehadiran_normal, _
            prosentase_disiplin_absen, telat As String

        For y As Byte = 3 To 35
            If DGV.Rows(y).Cells(2).Value.ToString = DGV.Rows(3).Cells(2).Value.ToString Then
                jml_hri_krja += 1

            Else
                jumlah_hari_kerja = jml_hri_krja

                ConnA = New OleDbConnection("provider=microsoft.ace.oledb.12.0;data source=" & Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\Evaluasi Absensi\Database Evaluasi Absensi.accdb")
                cmd = New OleDbCommand("UPDATE " & tbl & " SET JUMLAH_HARI_KERJA=" & jumlah_hari_kerja, ConnA)
                ConnA.Open()
                cmd.ExecuteNonQuery()
                ConnA.Close()
                Exit For
            End If
        Next

        ConnA = New OleDbConnection("provider=microsoft.ace.oledb.12.0;data source=" & Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\Evaluasi Absensi\Database Evaluasi Absensi.accdb")
        ConnA.Open()
        nama_karyawan = DGV.Rows(3).Cells(2).Value.ToString
        For t As Integer = 3 To DGV.RowCount - 2
            nam_kary = DGV.Rows(t).Cells(2).Value.ToString
            If nam_kary <> nama_karyawan Then
                jml_khdiran_normal = jml_hri_krja - jml_sppd - jml_cuti - jml_sakit - jml_kosong _
                                    - jml_tdk_datang - jml_tdk_pulang
                prosentase_dspln_absn = (jml_khdiran_normal + jml_sppd + jml_cuti + jml_sakit) / jml_hri_krja
                
                jumlah_kehadiran_normal = jml_khdiran_normal
                jumlah_sppd = jml_sppd
                jumlah_cuti = jml_cuti
                jumlah_sakit = jml_sakit
                jumlah_kosong = jml_kosong
                jumlah_tidak_datang = jml_tdk_datang
                jumlah_tidak_pulang = jml_tdk_pulang
                telat = tlt
                prosentase_disiplin_absen = FormatPercent(prosentase_dspln_absn, 0)
                cmd = New OleDbCommand("UPDATE " & tbl & " SET Kehadiran_Normal=" & jumlah_kehadiran_normal & ", SPPD=" & jumlah_sppd & _
                                       ", CUTI_atau_IJIN=" & jumlah_cuti & _
                                       ", SAKIT=" & jumlah_sakit & ", TIDAK_ABSEN_atau_KOSONG=" _
                                       & jumlah_kosong & ", TIDAK_ABSEN_DATANG=" & _
                                       jumlah_tidak_datang & ", TIDAK_ABSEN_PULANG=" & _
                                       jumlah_tidak_pulang & ", Prosentase_Disiplin_Absen='" & prosentase_disiplin_absen & "', Jumlah_terlambat=" & telat & " Where NAMA_PEGAWAI='" & _
                                       nama_karyawan & "'", ConnA)
                cmd.ExecuteNonQuery()

                jml_sppd = 0
                jumlah_sppd = jml_sppd

                jml_cuti = 0
                jumlah_cuti = jml_cuti

                jml_sakit = 0
                jumlah_sakit = jml_sakit

                jml_kosong = 0
                jumlah_kosong = jml_kosong

                jml_tdk_datang = 0
                jumlah_tidak_datang = jml_tdk_datang

                jml_tdk_pulang = 0
                jumlah_tidak_pulang = jml_tdk_pulang

                jml_khdiran_normal = 0
                jumlah_kehadiran_normal = jml_khdiran_normal

                prosentase_dspln_absn = 0
                prosentase_disiplin_absen = FormatPercent(prosentase_dspln_absn, 0)

                tlt = 0
                telat = tlt

                If DGV.Rows(t).Cells(5).Value.ToString = "SPPD" Then
                    jml_sppd += 1
                ElseIf DGV.Rows(t).Cells(5).Value.ToString = "CUTI" Then
                    jml_cuti += 1
                ElseIf DGV.Rows(t).Cells(5).Value.ToString = "SAKIT" Then
                    jml_sakit += 1
                ElseIf DGV.Rows(t).Cells(4).Value.ToString = "" And DGV.Rows(t).Cells(5).Value.ToString = "" Or DGV.Rows(t).Cells(4).Value.ToString.ToLower.StartsWith("0,") And DGV.Rows(t).Cells(5).Value.ToString.ToLower.StartsWith("0,") Then
                    jml_kosong += 1
                ElseIf DGV.Rows(t).Cells(4).Value.ToString = "" Or DGV.Rows(t).Cells(4).Value.ToString.ToLower.StartsWith("0,") Then
                    jml_tdk_datang += 1
                ElseIf DGV.Rows(t).Cells(5).Value.ToString = "" Or DGV.Rows(t).Cells(5).Value.ToString.ToLower.StartsWith("0,") Then
                    jml_tdk_pulang += 1
                Else
                    jam = New DateTime
                    jam = DGV.Rows(t).Cells(4).Value.ToString

                    Select Case jam
                        Case TimeValue("07:31:00") To TimeValue("16:30:00")
                            tlt += 1
                    End Select
                End If

                nama_karyawan = DGV.Rows(t).Cells(2).Value.ToString
            Else
                If DGV.Rows(t).Cells(5).Value.ToString = "SPPD" Then
                    jml_sppd += 1
                ElseIf DGV.Rows(t).Cells(5).Value.ToString = "CUTI" Then
                    jml_cuti += 1
                ElseIf DGV.Rows(t).Cells(5).Value.ToString = "SAKIT" Then
                    jml_sakit += 1
                ElseIf DGV.Rows(t).Cells(4).Value.ToString = "" And DGV.Rows(t).Cells(5).Value.ToString = "" Or DGV.Rows(t).Cells(4).Value.ToString.ToLower.StartsWith("0,") And DGV.Rows(t).Cells(5).Value.ToString.ToLower.StartsWith("0,") Then
                    jml_kosong += 1
                ElseIf DGV.Rows(t).Cells(4).Value.ToString = "" Or DGV.Rows(t).Cells(4).Value.ToString.ToLower.StartsWith("0,") Then
                    jml_tdk_datang += 1
                ElseIf DGV.Rows(t).Cells(5).Value.ToString = "" Or DGV.Rows(t).Cells(5).Value.ToString.ToLower.StartsWith("0,") Then
                    jml_tdk_pulang += 1
                Else
                    jam = New DateTime
                    jam = DGV.Rows(t).Cells(4).Value.ToString

                    Select Case jam
                        Case TimeValue("07:31:00") To TimeValue("16:30:00")
                            tlt += 1
                    End Select
                End If
            End If
        Next

        jml_khdiran_normal = jml_hri_krja - jml_sppd - jml_cuti - jml_sakit - jml_kosong _
                                    - jml_tdk_datang - jml_tdk_pulang
        prosentase_dspln_absn = (jml_khdiran_normal + jml_sppd + jml_cuti + jml_sakit) / jml_hri_krja

        jumlah_kehadiran_normal = jml_khdiran_normal
        jumlah_sppd = jml_sppd
        jumlah_cuti = jml_cuti
        jumlah_sakit = jml_sakit
        jumlah_kosong = jml_kosong
        jumlah_tidak_datang = jml_tdk_datang
        jumlah_tidak_pulang = jml_tdk_pulang
        telat = tlt
        prosentase_disiplin_absen = FormatPercent(prosentase_dspln_absn, 0)
        cmd = New OleDbCommand("UPDATE " & tbl & " SET Kehadiran_Normal=" & jumlah_kehadiran_normal & ", SPPD=" & jumlah_sppd & _
                               ", CUTI_atau_IJIN=" & jumlah_cuti & _
                               ", SAKIT=" & jumlah_sakit & ", TIDAK_ABSEN_atau_KOSONG=" _
                               & jumlah_kosong & ", TIDAK_ABSEN_DATANG=" & _
                               jumlah_tidak_datang & ", TIDAK_ABSEN_PULANG=" & _
                               jumlah_tidak_pulang & ", Prosentase_Disiplin_Absen='" & prosentase_disiplin_absen & "', Jumlah_terlambat=" & telat & " Where NAMA_PEGAWAI='" & _
                               nama_karyawan & "'", ConnA)
        cmd.ExecuteNonQuery()

        jml_sppd = 0
        jumlah_sppd = jml_sppd

        jml_cuti = 0
        jumlah_cuti = jml_cuti

        jml_sakit = 0
        jumlah_sakit = jml_sakit

        jml_kosong = 0
        jumlah_kosong = jml_kosong

        jml_tdk_datang = 0
        jumlah_tidak_datang = jml_tdk_datang

        jml_tdk_pulang = 0
        jumlah_tidak_pulang = jml_tdk_pulang

        jml_khdiran_normal = 0
        jumlah_kehadiran_normal = jml_khdiran_normal

        prosentase_dspln_absn = 0
        prosentase_disiplin_absen = FormatPercent(prosentase_dspln_absn, 0)

        tlt = 0
        telat = tlt

        ConnA.Close()
        ' khusus jumlah hari kerja
        jml_hri_krja = 0
        jumlah_hari_kerja = jml_hri_krja

        MsgBox("Done")
    End Sub

    Private Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim xlFile, hh As String

        xlFile = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\" & tbl & ".xlsx"
        If Not IO.File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\" & tbl & ".xlsx") Then
            Try
                xlWorkBook = xlApp.Workbooks.Add
                xlWorkBook.SaveAs(xlFile)
                xlWorkBook.Close()
            Catch ex As Exception
                xlWorkBook.Close(False)
                MsgBox("Kill all EXCEL.EXE processes in Task Manager! or repair your Office")
                Dialog1.Close()
                Me.Close()
                Exit Sub
            End Try
        Else
            MsgBox("File Excel Sudah Ada atau Sudah Dibuat")
            Exit Sub
        End If

        ConnA = New OleDbConnection("provider=microsoft.ace.oledb.12.0;data source=" & Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\Evaluasi Absensi\Database Evaluasi Absensi.accdb")
        ConnA.Open()

        cmd = New OleDbCommand("SELECT BAGIAN FROM " & tbl, ConnA)
        DRA = cmd.ExecuteReader
        DRA.Read()
        hh = DRA.GetString(0)
        cmd = New OleDbCommand("SELECT * INTO [Excel 12.0 Xml;DATABASE=" & xlFile & ";HDR=Yes].[" & hh & "] FROM " & tbl & " Where BAGIAN='" & hh & "'", ConnA)
        cmd.ExecuteNonQuery()
        While DRA.Read
            If DRA.GetString(0) <> hh Then
                hh = DRA.GetString(0)
                cmd = New OleDbCommand("SELECT * INTO [Excel 12.0 Xml;DATABASE=" & xlFile & ";HDR=Yes].[" & hh & "] FROM " & tbl & " Where BAGIAN='" & hh & "'", ConnA)
                cmd.ExecuteNonQuery()
            End If
        End While
        DRA.Close()
        ConnA.Close()

        Button3.Enabled = False
    End Sub

    Private Sub Form1_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        xlApp.Quit()
        xlWorkBook = Nothing
        xlApp = Nothing

        db = Nothing
        engine = Nothing
    End Sub
End Class
