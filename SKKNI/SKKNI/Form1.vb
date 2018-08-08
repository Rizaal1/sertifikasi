Imports System.Data.OleDb
Public Class Form1
    Dim Conn As OleDbConnection
    Dim da As OleDbDataAdapter
    Dim ds As DataSet
    Dim LokasiDB As String
    Sub Koneksi()
        LokasiDB = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Database3.accdb"
        Conn = New OleDbConnection(LokasiDB)
        If Conn.State = ConnectionState.Closed Then Conn.Open()
    End Sub
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Koneksi()
        da = New OleDbDataAdapter("Select * from dtbs_skkni", Conn)
        ds = New DataSet
        ds.Clear()
        da.Fill(ds, "dtbs_skkni")
        DataGridView1.DataSource = (ds.Tables("dtbs_skkni"))
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Then
            MsgBox("Masukkan NIK")
        ElseIf TextBox2.Text = "" Then
            MsgBox("Masukkan Nama")
        ElseIf TextBox3.Text = "" Then
            MsgBox("Masukkan HP")
        ElseIf TextBox4.Text = "" Then
            MsgBox("Masukkan email")
        ElseIf TextBox5.Text = "" Then
            MsgBox("Masukkan Skema Sertifikasi")
        ElseIf TextBox6.Text = "" Then
            MsgBox("Masukkan Tempat Uji Kompetensi")
        ElseIf TextBox8.Text = "" Then
            MsgBox("Masukkan Tanggal Terbit")
        ElseIf TextBox9.Text = "" Then
            MsgBox("Masukkan Tanggal Lahir")
        ElseIf TextBox10.Text = "" Then
            MsgBox("Masukkan Organisasi")
        ElseIf ComboBox1.Text = "" Then
            MsgBox("Silahkan Isi Semua Form")
        Else
            Dim CMD As OleDbCommand
            Call Koneksi()
            Dim simpan As String = "insert into dtbs_skkni values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "','" & TextBox6.Text & "','" & TextBox8.Text & "','" & TextBox9.Text & "','" & TextBox10.Text & "','" & ComboBox1.Text & "')"
            CMD = New OleDbCommand(simpan, Conn)
            CMD.ExecuteNonQuery()

        End If
    End Sub
    
    Private Sub TextBox1_KeyPress1(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        TextBox1.MaxLength = 6
        If e.KeyChar = Chr(13) Then
            Call Koneksi()
            Dim CMD As OleDbCommand
            Dim RD As OleDbDataReader
            CMD = New OleDbCommand("Select * From dtbs_skkni  where NIK='" & TextBox1.Text & "'", Conn)
            RD = CMD.ExecuteReader
            RD.Read()
            If Not RD.HasRows Then
                MsgBox("Kode Barang Tidak Ada, Silahkan coba lagi!")
                TextBox1.Focus()
            Else
                TextBox2.Text = RD.Item("nama")
                TextBox3.Text = RD.Item("hp")
                TextBox4.Text = RD.Item("email")
                TextBox5.Text = RD.Item("skema_sertifikasi")
                TextBox6.Text = RD.Item("tmpuji_kompetensi")
                TextBox8.Text = RD.Item("rekomendasi")
                TextBox9.Text = RD.Item("tglterbit_sertifikasi")
                TextBox10.Text = RD.Item("tgl_lahir")
                ComboBox1.Text = RD.Item("organisasi")
                TextBox2.Focus()
            End If
        End If
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
    End Sub
End Class
