Imports System.Data.OleDb
Imports ReimeFinderDatenbank.ReimeDataSet
Imports System.IO
Imports NHunspell
'Setup neu
Public Class Form1
#Region "Datenbank"
    Dim con As New OleDbConnection
    Dim cmd As New OleDbCommand
    Dim reader As OleDbDataReader
#End Region
    Dim ReimendesWort As String
    Dim Rechtschr As New Hunspell
    Dim grammar As MyThes

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        My.Settings.text = My.Computer.Clock.LocalTime.Day & "." & My.Computer.Clock.LocalTime.Month & "." & My.Computer.Clock.LocalTime.Year & vbNewLine & RichTextBox1.Text
        My.Settings.Save()

    End Sub

    'Rechtschreibkorrekturvorschläge in der Listboy anzeigen
    'Beim Klick auf Hinzufügen das WOrt in die Datenbank schreiben
    Private Sub Form1_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        If ListBox1.Focused = True Then
            RichTextBox1.Text += ListBox1.SelectedItem
        End If

    End Sub

 
    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'Falsche URL
      

        MsgBox(Application.StartupPath)
        'Dim res As MsgBoxResult
        If My.Settings.Pfad = "" Then
            MsgBox("Geben Sie den Pfad zum Ordner in dem sich die Dateien befinden")
            FolderBrowserDialog1.ShowDialog()
            My.Settings.Pfad = FolderBrowserDialog1.SelectedPath
            My.Settings.Save()
            File.CreateText(My.Settings.Pfad & "\Wörterbuch.txt")
            File.Copy(My.Settings.Pfad & "\Hunspellx64.dll", Application.StartupPath & "\Hunspellx64.dll")
            File.Copy(My.Settings.Pfad & "\Hunspellx86.dll", Application.StartupPath & "\Hunspellx86.dll")

            'MsgBox("Und zuletzt kopiere bitte noch die beiden beiliegenden .dll Dateinen an diesen Pfad" & Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData))
        End If

        If My.Settings.Pfad <> "" Then
            Try
                grammar = New MyThes(My.Settings.Pfad & "\th_de_AT_v2.idx", My.Settings.Pfad & "\th_de_AT_v2.dat")
            Catch ex As Exception
                MsgBox("Geben Sie den Pfad zum Ordner in dem sich die Dateien befinden")
                FolderBrowserDialog1.ShowDialog()
                My.Settings.Pfad = FolderBrowserDialog1.SelectedPath
                My.Settings.Save()
            End Try


            Rechtschr.Load(My.Settings.Pfad & "\de-AT.aff", My.Settings.Pfad & "\de-AT.dic")





            If My.Computer.Network.IsAvailable = True Then
                InternetToolStripMenuItem.Checked = True
            Else
                InternetToolStripMenuItem.Checked = False
            End If
        End If



    End Sub


    Private Sub FindenToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles FindenToolStripMenuItem.Click
        ReimendesWort = RichTextBox1.SelectedText
        ListBox1.Items.Clear()
        If InternetToolStripMenuItem.Checked = True Then
            SucheOnline()

        End If
        If LokalToolStripMenuItem.Checked = True Then
            If RichTextBox1.SelectedText <> "" Then
                ReimFinden(RichTextBox1.SelectedText)
            Else
                Dim le As String = InputBox("Geben Sie das zu suchende Wort ein", "Reime finden")
                ReimFinden(le)
            End If
        End If


    End Sub
    Private Sub ReimFinden(ByVal suchwort As String)
        con.Close()

        con.ConnectionString =
     My.Settings.ReimeConnectionString
        '"Data Source=C:\Users\Normal\Documents\Produkte1.mdb"
        cmd.Connection = con
        cmd.CommandText = "select * from Reime"
        Try
            con.Open()

            reader = cmd.ExecuteReader()

            Dim index As Integer = 0
            Do While reader.Read()
                'reader("Reim1")
                'Leere Felder Herausnehmen
                If suchwort = reader("ZureimendesFeld").ToString Then
                    
                    If reader("Reim1").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim1")).ToString()
                    End If

                    If reader("Reim2").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim2")).ToString()
                    End If
                    If reader("Reim3").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim3")).ToString()
                    End If
                    If reader("Reim4").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim4")).ToString()
                    End If
                    If reader("Reim5").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim5")).ToString()
                    End If

                    If reader("Reim6").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim6")).ToString()
                    End If
                    If reader("Reim7").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim7")).ToString()
                    End If
                    If reader("Reim8").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim8")).ToString()
                    End If
                    If reader("Reim9").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim9")).ToString()
                    End If

                ElseIf suchwort = reader("Reim1").ToString Then
                    If reader("ZureimendesFeld").ToString <> "" Then
                        ListBox1.Items.Add(reader("ZureimendesFeld")).ToString()
                    End If
                  

                    If reader("Reim2").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim2")).ToString()
                    End If
                    If reader("Reim3").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim3")).ToString()
                    End If
                    If reader("Reim4").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim4")).ToString()
                    End If
                    If reader("Reim5").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim5")).ToString()
                    End If

                    If reader("Reim6").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim6")).ToString()
                    End If
                    If reader("Reim7").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim7")).ToString()
                    End If
                    If reader("Reim8").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim8")).ToString()
                    End If
                    If reader("Reim9").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim9")).ToString()
                    End If

                ElseIf suchwort = reader("Reim2").ToString Then
                    If reader("ZureimendesFeld").ToString <> "" Then
                        ListBox1.Items.Add(reader("ZureimendesFeld")).ToString()
                    End If
                    If reader("Reim1").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim1")).ToString()
                    End If

                 
                    If reader("Reim3").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim3")).ToString()
                    End If
                    If reader("Reim4").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim4")).ToString()
                    End If
                    If reader("Reim5").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim5")).ToString()
                    End If

                    If reader("Reim6").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim6")).ToString()
                    End If
                    If reader("Reim7").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim7")).ToString()
                    End If
                    If reader("Reim8").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim8")).ToString()
                    End If
                    If reader("Reim9").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim9")).ToString()
                    End If


                ElseIf suchwort = reader("Reim3").ToString Then
                    If reader("ZureimendesFeld").ToString <> "" Then
                        ListBox1.Items.Add(reader("ZureimendesFeld")).ToString()
                    End If
                    If reader("Reim1").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim1")).ToString()
                    End If

                    If reader("Reim2").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim2")).ToString()
                    End If
                  
                    If reader("Reim4").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim4")).ToString()
                    End If
                    If reader("Reim5").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim5")).ToString()
                    End If

                    If reader("Reim6").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim6")).ToString()
                    End If
                    If reader("Reim7").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim7")).ToString()
                    End If
                    If reader("Reim8").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim8")).ToString()
                    End If
                    If reader("Reim9").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim9")).ToString()
                    End If

                ElseIf suchwort = reader("Reim4").ToString Then
                    If reader("ZureimendesFeld").ToString <> "" Then
                        ListBox1.Items.Add(reader("ZureimendesFeld")).ToString()
                    End If
                    If reader("Reim1").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim1")).ToString()
                    End If

                    If reader("Reim2").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim2")).ToString()
                    End If
                    If reader("Reim3").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim3")).ToString()
                    End If
                  
                    If reader("Reim5").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim5")).ToString()
                    End If

                    If reader("Reim6").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim6")).ToString()
                    End If
                    If reader("Reim7").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim7")).ToString()
                    End If
                    If reader("Reim8").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim8")).ToString()
                    End If
                    If reader("Reim9").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim9")).ToString()
                    End If


                ElseIf suchwort = reader("Reim5").ToString Then
                    If reader("ZureimendesFeld").ToString <> "" Then
                        ListBox1.Items.Add(reader("ZureimendesFeld")).ToString()
                    End If
                    If reader("Reim1").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim1")).ToString()
                    End If

                    If reader("Reim2").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim2")).ToString()
                    End If
                    If reader("Reim3").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim3")).ToString()
                    End If
                    If reader("Reim4").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim4")).ToString()
                    End If
                  

                    If reader("Reim6").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim6")).ToString()
                    End If
                    If reader("Reim7").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim7")).ToString()
                    End If
                    If reader("Reim8").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim8")).ToString()
                    End If
                    If reader("Reim9").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim9")).ToString()
                    End If


                ElseIf suchwort = reader("Reim6").ToString Then
                    If reader("ZureimendesFeld").ToString <> "" Then
                        ListBox1.Items.Add(reader("ZureimendesFeld")).ToString()
                    End If
                    If reader("Reim1").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim1")).ToString()
                    End If

                    If reader("Reim2").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim2")).ToString()
                    End If
                    If reader("Reim3").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim3")).ToString()
                    End If
                    If reader("Reim4").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim4")).ToString()
                    End If
                    If reader("Reim5").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim5")).ToString()
                    End If

                  
                    If reader("Reim7").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim7")).ToString()
                    End If
                    If reader("Reim8").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim8")).ToString()
                    End If
                    If reader("Reim9").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim9")).ToString()
                    End If


                ElseIf suchwort = reader("Reim7").ToString Then
                    If reader("ZureimendesFeld").ToString <> "" Then
                        ListBox1.Items.Add(reader("ZureimendesFeld")).ToString()
                    End If
                    If reader("Reim1").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim1")).ToString()
                    End If

                    If reader("Reim2").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim2")).ToString()
                    End If
                    If reader("Reim3").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim3")).ToString()
                    End If
                    If reader("Reim4").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim4")).ToString()
                    End If
                    If reader("Reim5").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim5")).ToString()
                    End If

                    If reader("Reim6").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim6")).ToString()
                    End If
                   
                    If reader("Reim8").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim8")).ToString()
                    End If
                    If reader("Reim9").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim9")).ToString()
                    End If

                ElseIf suchwort = reader("Reim8").ToString Then
                    If reader("ZureimendesFeld").ToString <> "" Then
                        ListBox1.Items.Add(reader("ZureimendesFeld")).ToString()
                    End If
                    If reader("Reim1").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim1")).ToString()
                    End If

                    If reader("Reim2").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim2")).ToString()
                    End If
                    If reader("Reim3").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim3")).ToString()
                    End If
                    If reader("Reim4").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim4")).ToString()
                    End If
                    If reader("Reim5").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim5")).ToString()
                    End If

                    If reader("Reim6").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim6")).ToString()
                    End If
                    If reader("Reim7").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim7")).ToString()
                    End If
                   
                    If reader("Reim9").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim9")).ToString()
                    End If


                ElseIf suchwort = reader("Reim9").ToString Then
                    If reader("ZureimendesFeld").ToString <> "" Then
                        ListBox1.Items.Add(reader("ZureimendesFeld")).ToString()
                    End If
                    If reader("Reim1").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim1")).ToString()
                    End If

                    If reader("Reim2").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim2")).ToString()
                    End If
                    If reader("Reim3").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim3")).ToString()
                    End If
                    If reader("Reim4").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim4")).ToString()
                    End If
                    If reader("Reim5").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim5")).ToString()
                    End If

                    If reader("Reim6").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim6")).ToString()
                    End If
                    If reader("Reim7").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim7")).ToString()
                    End If
                    If reader("Reim8").ToString <> "" Then
                        ListBox1.Items.Add(reader("Reim8")).ToString()
                    End If
                   


                Else

                    index += 1
                End If



            Loop
            reader.Close()
            con.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            SucheOnline()
        End Try

        ListBox1.Focus()
    End Sub
    Private Sub SucheOnline()
        If My.Computer.Network.IsAvailable = True Then
            WebBrowser1.Navigate("http://www.reime.woxikon.de/ger/" & RichTextBox1.SelectedText & ".php")
        Else
            MsgBox("Sie haben keine Internetverbindung")
            InternetToolStripMenuItem.Checked = False
        End If

    End Sub
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        If ListBox1.SelectedItem <> "" Then


            con.Close()
            RichTextBox1.Text += ListBox1.SelectedItem
            con.ConnectionString =
            My.Settings.ReimeConnectionString
            Dim Wort() As String = ListBox1.SelectedItem.ToString.Split(" ")
            Dim de As String = Wort(Wort.Length - 1)
            Try


                cmd.CommandText = "insert into Reime (ZureimendesFeld, Reim1, Reim2, Reim3, Reim4, Reim5, Reim6, Reim7, Reim8, Reim9) values('" & de & "', '" & ReimendesWort & "' , '" & "" & "' , '" & "" & "' , '" & "" & "' , '" & "" & "' , '" & "" & "' , '" & "" & "' , '" & "" & "' , '" & "" & "')"
                cmd.Connection = con

                con.Open()
                cmd.ExecuteNonQuery()

                con.Close()
            Catch ex As Exception
                Try
                    cmd.CommandText = "insert into Reime (ZureimendesFeld, Reim1, Reim2, Reim3, Reim4, Reim5, Reim6, Reim7, Reim8, Reim9) values('" & de & "', '" & ListBox1.SelectedItem & "' , '" & "" & "' , '" & "" & "' , '" & "" & "' , '" & "" & "' , '" & "" & "' , '" & "" & "' , '" & "" & "' , '" & "" & "')"
                    cmd.Connection = con

                    con.Open()
                    cmd.ExecuteNonQuery()

                    con.Close()
                Catch exe As Exception

                End Try
            End Try
            Dim Alts As String = File.ReadAllText(My.Settings.Pfad & "\Wörterbuch.txt")
            Dim Zeichen() As Char = ListBox1.SelectedItem.ToString.ToCharArray
            Dim Wortsd As String
            For Each el As Char In Zeichen
                If el <> "" Then
                    Wortsd += el
                End If
            Next
            File.WriteAllText(My.Settings.Pfad & "\Wörterbuch.txt", Alts & ReimendesWort & "," & de & ";")
        End If
    End Sub

    Private Sub HinzufügenToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles HinzufügenToolStripMenuItem.Click

        con.ConnectionString =
         My.Settings.ReimeConnectionString

        cmd.CommandText = "insert into Reime (ZureimendesFeld, Reim1, Reim2, Reim3, Reim4, Reim5, Reim6, Reim7, Reim8, Reim9) values('" & InputBox("Geben Sie ein Reimwort ein", "Reimwort") & "', '" & InputBox("Geben Sie ein sich reimendes Wort ein", "Reimwort") & "' , '" & InputBox("Geben Sie ein sich reimendes Wort ein", "Reimwort") & "' , '" & InputBox("Geben Sie ein sich reimendes Wort ein", "Reimwort") & "' , '" & InputBox("Geben Sie ein sich reimendes Wort ein", "Reimwort") & "' , '" & InputBox("Geben Sie ein sich reimendes Wort ein", "Reimwort") & "' , '" & InputBox("Geben Sie ein sich reimendes Wort ein", "Reimwort") & "' , '" & InputBox("Geben Sie ein sich reimendes Wort ein", "Reimwort") & "' , '" & InputBox("Geben Sie ein sich reimendes Wort ein", "Reimwort") & "' , '" & InputBox("Geben Sie ein sich reimendes Wort ein", "Reimwort") & "')"
        cmd.Connection = con

        con.Open()
        cmd.ExecuteNonQuery()

        con.Close()
       
    End Sub

    Private Sub ReimFindenToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ReimFindenToolStripMenuItem.Click
        ListBox1.Items.Clear()
        If InternetToolStripMenuItem.Checked = True Then
            SucheOnline()
        End If
        If LokalToolStripMenuItem.Checked = True Then
            ReimFinden(RichTextBox1.SelectedText)
        End If
        If InternetToolStripMenuItem.Checked = False And LokalToolStripMenuItem.Checked = False Then
            MsgBox("Sie haben keinen Suchort ausgewählt")
        End If
    End Sub

    Private Sub AlleAnzeigenToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AlleAnzeigenToolStripMenuItem.Click
        con.Close()
        con.ConnectionString =
    My.Settings.ReimeConnectionString
        '"Data Source=C:\Users\Normal\Documents\Produkte1.mdb"
        cmd.Connection = con
        cmd.CommandText = "select * from Reime"
        Try
            con.Open()

            reader = cmd.ExecuteReader()
            Dim Reime As String
            Do While reader.Read()
                Reime += reader("ZureimendesFeld").ToString() & " "
                Reime += reader("Reim1").ToString() & " "
                Reime += reader("Reim2").ToString() & " "
                Reime += reader("Reim4").ToString() & " "
                Reime += reader("Reim5").ToString() & " "

                Reime += reader("Reim6").ToString() & " "
                Reime += reader("Reim7").ToString() & " "
                Reime += (reader("Reim8")).ToString() & " "
                Reime += (reader("Reim3")).ToString() & " "
                Reime += (reader("Reim9")).ToString() & vbNewLine
            Loop
            MsgBox(Reime)
        Catch ex As Exception
        End Try
        reader.Close()
        con.Close()
    End Sub

    Private Sub WebBrowser1_DocumentCompleted(ByVal sender As Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted

        Try
            Dim text() As String
            text = WebBrowser1.Document.GetElementById("content").InnerText.Split("Deutsch")


            For i = 1 To text.Length - 1
                If text(i) <> "" Or text(i) <> " " Then
                    ListBox1.Items.Add(text(i).Substring(6))
                End If

            Next
        Catch ex As Exception

        End Try

    End Sub

    Private Sub InternetToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InternetToolStripMenuItem.Click
        If InternetToolStripMenuItem.Checked = False And LokalToolStripMenuItem.Checked = False Then
            MsgBox("Sie haben keinen Suchort ausgewählt")
        End If

    End Sub

    Private Sub LokalToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LokalToolStripMenuItem.Click
        If InternetToolStripMenuItem.Checked = False And LokalToolStripMenuItem.Checked = False Then
            MsgBox("Sie haben keinen Suchort ausgewählt")
        End If
    End Sub

    Private Sub ÖffnenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ÖffnenToolStripMenuItem.Click
        OpenFileDialog1.ShowDialog()
        Try
            RichTextBox1.LoadFile(OpenFileDialog1.FileName)
        Catch ex As Exception

            Dim sr As New StreamReader(OpenFileDialog1.FileName)
            RichTextBox1.Text = sr.ReadToEnd
            sr.Close()
        End Try
    End Sub

    Private Sub NeuToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NeuToolStripMenuItem.Click
        Dim result As MsgBoxResult
        result = MsgBox("Sind Sie sich sicher das Sie den Text löschen wollen?", MsgBoxStyle.OkCancel, "Neu")
        If result = MsgBoxResult.Ok Then
            RichTextBox1.Clear()
        End If
    End Sub

    Private Sub SpeichernUnterToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SpeichernUnterToolStripMenuItem.Click
        SaveFileDialog1.ShowDialog()
        RichTextBox1.SaveFile(SaveFileDialog1.FileName)
    End Sub

    Private Sub SpeichernToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SpeichernToolStripMenuItem.Click
        If SaveFileDialog1.FileName = "" Then
            SaveFileDialog1.ShowDialog()
            RichTextBox1.SaveFile(SaveFileDialog1.FileName)
        Else
            RichTextBox1.SaveFile(SaveFileDialog1.FileName)
        End If
    End Sub

    Private Sub DruckenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DruckenToolStripMenuItem.Click
        PrintDialog1.ShowDialog()

        PrintDocument1.PrinterSettings = PrintDialog1.PrinterSettings
        Dim result As MsgBoxResult
        result = MsgBox("Wollen Sie den Text drucken?", MsgBoxStyle.OkCancel, "Drucken")
        If result = MsgBoxResult.Ok Then
            PrintDocument1.Print()
        End If

    End Sub

    Private Sub RestToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RestToolStripMenuItem.Click
        FontDialog1.ShowDialog()
        RichTextBox1.SelectionFont = FontDialog1.Font
    End Sub

    Private Sub FarbeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FarbeToolStripMenuItem.Click
        ColorDialog1.ShowDialog()
        RichTextBox1.SelectionColor = ColorDialog1.Color
    End Sub

    Private Sub TastenkürzelToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TastenkürzelToolStripMenuItem.Click

    End Sub

    Private Sub FormatierenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FormatierenToolStripMenuItem.Click
        MsgBox("Mittig: Strg + E" & vbNewLine & "Links: Strg + L" & vbNewLine & "Rechts: Strg + R")
    End Sub

    Private Sub ReimeToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReimeToolStripMenuItem1.Click

    End Sub

    Private Sub ReimeSuchenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReimeSuchenToolStripMenuItem.Click
        MsgBox("Reimwort auswählen und dann auf Finden klicken")
    End Sub

    Private Sub ReimeHinzufügenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReimeHinzufügenToolStripMenuItem.Click
        MsgBox("Im Reime Menü auf Hinzufügen klicken und dann die Felder ausfüllen und auf OK klicken. Es müssen nicht alle Felder ausgefüllt werden. Nur das Erste muss ausgefüllt werden. Die hinzugefügten Reime werden Lokal gespeichert")
    End Sub

    Private Sub AlleAnzeigenToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AlleAnzeigenToolStripMenuItem1.Click
        MsgBox("Diese Funktion zeigt ihnen alle gespeicherten Wörter")
    End Sub

    Private Sub SuchorteToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SuchorteToolStripMenuItem1.Click
        MsgBox("Hier können Sie einstellen ob Sie nur lokale Wörter oder auch Wörter aus dem Internet suchen wollen")
    End Sub

    Private Sub SupportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SupportToolStripMenuItem.Click
        MsgBox("Bitte besuchen Sie die Seite http://reimefinder.jimdo.com/")
    End Sub

    Private Sub TopicsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub AutoReimToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AutoReimToolStripMenuItem.Click

    End Sub

    Private Sub RichTextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles RichTextBox1.KeyDown

        If AutoReimToolStripMenuItem.Checked = True Then
            If e.KeyCode = Keys.Space Or e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Decimal Then
                Dim Wörter() As String = RichTextBox1.Text.Split(" ")
                ListBox1.Items.Clear()
                ReimFinden(Wörter(Wörter.Length - 1))

                Application.DoEvents()
                RichTextBox1.Focus()
            End If
        End If
    End Sub

    Private Sub AutoReimToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AutoReimToolStripMenuItem1.Click
        MsgBox("Der Computer sucht wenn bei 'AutoReim' ein Häkchen ist automatisch für jedes Wort ein Reimwort. Dieses sehen Sie dann wie gewohnt rechts", MsgBoxStyle.Information)
    End Sub

    Private Sub EinfügenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EinfügenToolStripMenuItem.Click
        RichTextBox1.Paste()
    End Sub

    Private Sub KopierenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KopierenToolStripMenuItem.Click
        RichTextBox1.Copy()
    End Sub

    Private Sub AllgemeinToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AllgemeinToolStripMenuItem.Click
        MsgBox("Die Tastenkürzel die nicht für das formatieren des Textes benutzt werden finden Sie neben der Funktion welche Sie benutzen wollen. Zum Beispiel Strg + Alt + A für AutoReime. Sie finden eine Auflistung der Tastenkürzel für das Formatieren des Textes in der Hilfe unter 'Tastenkürzel' 'Formatieren'", MsgBoxStyle.Information)
    End Sub

    Private Sub VersionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VersionToolStripMenuItem.Click
        MsgBox("2.4")
    End Sub

    Private Sub InstallationsortToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InstallationsortToolStripMenuItem.Click
        MsgBox(My.Application.Deployment.DataDirectory)
    End Sub

    Private Sub EinfügenToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EinfügenToolStripMenuItem1.Click
        If ListBox1.SelectedItem <> "" Then


            con.Close()
            RichTextBox1.Text += ListBox1.SelectedItem
            con.ConnectionString =
            My.Settings.ReimeConnectionString
            Dim Wort() As String = ListBox1.SelectedItem.ToString.Split(" ")
            Dim de As String = Wort(Wort.Length - 1)
            Try


                cmd.CommandText = "insert into Reime (ZureimendesFeld, Reim1, Reim2, Reim3, Reim4, Reim5, Reim6, Reim7, Reim8, Reim9) values('" & de & "', '" & ReimendesWort & "' , '" & "" & "' , '" & "" & "' , '" & "" & "' , '" & "" & "' , '" & "" & "' , '" & "" & "' , '" & "" & "' , '" & "" & "')"
                cmd.Connection = con

                con.Open()
                cmd.ExecuteNonQuery()

                con.Close()
            Catch ex As Exception
                Try
                    cmd.CommandText = "insert into Reime (ZureimendesFeld, Reim1, Reim2, Reim3, Reim4, Reim5, Reim6, Reim7, Reim8, Reim9) values('" & de & "', '" & ListBox1.SelectedItem & "' , '" & "" & "' , '" & "" & "' , '" & "" & "' , '" & "" & "' , '" & "" & "' , '" & "" & "' , '" & "" & "' , '" & "" & "')"
                    cmd.Connection = con

                    con.Open()
                    cmd.ExecuteNonQuery()

                    con.Close()
                Catch exe As Exception

                End Try
            End Try
            Dim Alts As String = File.ReadAllText(My.Settings.Pfad & "\Wörterbuch.txt")
          
            File.WriteAllText(My.Settings.Pfad & "\Wörterbuch.txt", Alts & ReimendesWort & "," & de & ";")
        End If

    End Sub

    Private Sub FarbeToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FarbeToolStripMenuItem1.Click
        ColorDialog1.ShowDialog()
        RichTextBox1.SelectionColor = ColorDialog1.Color
    End Sub

    Private Sub EigenschaftenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EigenschaftenToolStripMenuItem.Click
        FontDialog1.ShowDialog()
        RichTextBox1.SelectionFont = FontDialog1.Font
    End Sub

    Private Sub LetzteSitzungHerstellenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LetzteSitzungHerstellenToolStripMenuItem.Click
        RichTextBox1.Text = My.Settings.text
    End Sub

    Private Sub LetzteSitzungLöschenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LetzteSitzungLöschenToolStripMenuItem.Click
        My.Settings.text = ""
        My.Settings.Save()
    End Sub

    Private Sub ZurückToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ZurückToolStripMenuItem.Click
        MsgBox("Strg + Z")
    End Sub

    Private Sub ZeichenZählernToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ZeichenZählernToolStripMenuItem.Click

    End Sub

    Private Sub WörterZählenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WörterZählenToolStripMenuItem.Click
        Dim w() As String = RichTextBox1.Text.Split(" ")
        MsgBox(w.Length)
    End Sub

    Private Sub MitLeerzeichenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MitLeerzeichenToolStripMenuItem.Click
        MsgBox(RichTextBox1.Text.Length)
    End Sub

    Private Sub OhneLeerzeichenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OhneLeerzeichenToolStripMenuItem.Click
        Dim w() As String = RichTextBox1.Text.Split(" ")
        Dim lä As Integer
        For Each ed As String In w
            lä += ed.Length
        Next
        MsgBox(lä)
    End Sub

    Private Sub RechtschreibungToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RechtschreibungToolStripMenuItem.Click

        Dim liste As New List(Of String)
        Dim Wörter As String() = RichTextBox1.Text.Split(" ")
        Dim index As Integer = 0
        For Each Wort As String In Wörter
            If Rechtschr.Spell(Wort) Then

            Else

                Dim res As MsgBoxResult
                res = MsgBox("Das Wort " & Wort & " wurde falsch geschrieben", MsgBoxStyle.OkCancel)
                If res = MsgBoxResult.Ok Then
                    liste = Rechtschr.Suggest(Wort)

                    For Each el As String In liste
                        Dim rese As MsgBoxResult
                        rese = MsgBox("Könnte " & Wort & " vielleicht " & el & " bedeuten?", MsgBoxStyle.YesNoCancel)
                        If rese = MsgBoxResult.Yes Then
                            Wörter(index) = el
                            Exit For
                        End If
                        If rese = MsgBoxResult.Cancel Then
                            Exit For
                        End If

                    Next

                End If

                RichTextBox1.Text = ""
                For Each element As String In Wörter
                    RichTextBox1.Text += element & " "
                Next
            End If
            index += 1
        Next

    End Sub

    Private Sub UnterstreichenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UnterstreichenToolStripMenuItem.Click

        RichTextBox1.SelectionFont = New Font(RichTextBox1.SelectionFont, FontStyle.Underline)

    End Sub

    Private Sub SuchenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SuchenToolStripMenuItem.Click
        Dim a As String
        Dim b As String
        a = InputBox("Zu findenden Text eingeben")
        b = InStr(RichTextBox1.Text, a)
        If b Then
            RichTextBox1.Focus()
            RichTextBox1.SelectionStart = b - 1
            RichTextBox1.SelectionLength = Len(a)
        Else
            MsgBox("Text nicht gefunden.")
        End If
    End Sub

    Private Sub KursivToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KursivToolStripMenuItem.Click
        RichTextBox1.SelectionFont = New Font(RichTextBox1.SelectionFont, FontStyle.Italic)
    End Sub

    Private Sub FettToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FettToolStripMenuItem.Click
        RichTextBox1.SelectionFont = New Font(RichTextBox1.SelectionFont, FontStyle.Bold)
    End Sub

    Private Sub DurchgestrichenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DurchgestrichenToolStripMenuItem.Click
        RichTextBox1.SelectionFont = New Font(RichTextBox1.SelectionFont, FontStyle.Strikeout)
    End Sub

    Private Sub NormalToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NormalToolStripMenuItem.Click
        RichTextBox1.SelectionFont = New Font(RichTextBox1.SelectionFont, FontStyle.Regular)
    End Sub
  
    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click

        RichTextBox1.Size = New Size(RichTextBox1.Size.Width - 2, RichTextBox1.Size.Height)
        ListBox1.Size = New Size(ListBox1.Size.Width + 2, ListBox1.Size.Height)
        ListBox1.Location = New Point(ListBox1.Location.X - 2, 27)
        Button1.Size = New Size(Button1.Size.Width + 2, Button1.Size.Height)
        Button1.Location = New Point(Button1.Location.X - 2, Button1.Location.Y)
        Panel1.Location = New Point(Panel1.Location.X - 2, 27)
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        RichTextBox1.Size = New Size(RichTextBox1.Size.Width + 2, RichTextBox1.Size.Height)
        ListBox1.Size = New Size(ListBox1.Size.Width - 2, ListBox1.Size.Height)
        ListBox1.Location = New Point(ListBox1.Location.X + 2, 27)
        Button1.Size = New Size(Button1.Size.Width - 2, Button1.Size.Height)
        Button1.Location = New Point(Button1.Location.X + 2, Button1.Location.Y)
        Panel1.Location = New Point(Panel1.Location.X + 2, 27)

    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        RichTextBox1.Size = New Size(RichTextBox1.Size.Width - 10, RichTextBox1.Size.Height)
        ListBox1.Size = New Size(ListBox1.Size.Width + 10, ListBox1.Size.Height)
        ListBox1.Location = New Point(ListBox1.Location.X - 10, 27)
        Button1.Size = New Size(Button1.Size.Width + 10, Button1.Size.Height)
        Button1.Location = New Point(Button1.Location.X - 10, Button1.Location.Y)
        Panel1.Location = New Point(Panel1.Location.X - 10, 27)
    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        RichTextBox1.Size = New Size(RichTextBox1.Size.Width + 10, RichTextBox1.Size.Height)
        ListBox1.Size = New Size(ListBox1.Size.Width - 10, ListBox1.Size.Height)
        ListBox1.Location = New Point(ListBox1.Location.X + 10, 27)
        Button1.Size = New Size(Button1.Size.Width - 10, Button1.Size.Height)
        Button1.Location = New Point(Button1.Location.X + 10, Button1.Location.Y)
        Panel1.Location = New Point(Panel1.Location.X + 10, 27)
    End Sub

    Private Sub WortstammToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles WortstammToolStripMenuItem.Click
        If RichTextBox1.SelectedText <> "" Then
            Dim morphs As List(Of String) = Rechtschr.Stem(RichTextBox1.SelectedText)
            For Each morph As String In morphs
                MsgBox(morph)
            Next
        Else
            Dim morphs As List(Of String) = Rechtschr.Stem(InputBox("Wort um den Wortstamm zu ermitteln", "Wortstamm"))
            For Each morph As String In morphs
                MsgBox(morph)
            Next
        End If

    End Sub

    Private Sub SynonymeToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SynonymeToolStripMenuItem.Click
        If RichTextBox1.SelectedText <> "" Then
            Try
                Dim tr As New ThesResult(grammar.Lookup(RichTextBox1.SelectedText, Rechtschr).Meanings, False)
                For Each meaning As ThesMeaning In tr.Meanings
                    MsgBox(meaning.Description)


                    For Each synonym As String In meaning.Synonyms
                        Dim res As MsgBoxResult
                        res = MsgBox("    Synonym: " & synonym, MsgBoxStyle.OkCancel)
                        If res = MsgBoxResult.Cancel Then
                            Exit For
                        End If
                    Next
                Next
      
            Catch ex As Exception
                MsgBox("Bitte nur den Infinitiv")
            End Try




        End If

    End Sub

End Class
