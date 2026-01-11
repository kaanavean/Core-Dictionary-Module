Imports System.IO
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Light_Term_Encryption

Public Class Form1
    ' Instanz deiner hochsicheren Engine
    Private ReadOnly engine As New LTE

    ' Passwort-Variable (Könnte auch aus einer TextBox kommen)
    Private Const MySecretKey As String = "DeinSuperGeheimesPasswort123"

    ''' <summary>
    ''' Button 1: Kryptieren
    ''' Verschlüsselt den Text aus TextBox1 und zeigt das Ergebnis als Base64 an.
    ''' </summary>
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If String.IsNullOrWhiteSpace(TextBox1.Text) Then Return

        Try
            ' Protect(Text, UseHardwareBinding, Password)
            Dim encryptedData As Byte() = engine.Protect(TextBox1.Text, True, MySecretKey)

            ' In Base64 wandeln, um es in der TextBox anzuzeigen
            TextBox1.Text = Convert.ToBase64String(encryptedData)
            MessageBox.Show("Erfolgreich verschlüsselt!", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Fehler beim Verschlüsseln: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Button 2: Dekryptieren
    ''' Nimmt den Base64-Code aus TextBox1 und wandelt ihn in Klartext zurück.
    ''' </summary>
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If String.IsNullOrWhiteSpace(TextBox1.Text) Then Return

        Try
            ' Base64 zurück in Bytes wandeln
            Dim encryptedData As Byte() = Convert.FromBase64String(TextBox1.Text)

            ' Unprotect(Bytes, UseHardwareBinding, Password)
            Dim decryptedText As String = engine.Unprotect(encryptedData, True, MySecretKey)

            TextBox1.Text = decryptedText
        Catch ex As Exception
            MessageBox.Show("Entschlüsselung fehlgeschlagen! Falscher Key oder andere Hardware.", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Button 3: Speichern
    ''' Speichert den aktuell verschlüsselten Inhalt von TextBox1 in eine Datei.
    ''' </summary>
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If String.IsNullOrWhiteSpace(TextBox1.Text) Then Return

        Dim sfd As New SaveFileDialog With {
            .Filter = "Crypt Files (*.crypt)|*.crypt",
            .Title = "Verschlüsselte Datei speichern"
        }

        If sfd.ShowDialog() = DialogResult.OK Then
            Try
                ' Wir speichern die rohen Bytes der aktuellen Verschlüsselung
                ' Falls TextBox1 Base64 enthält, wandeln wir es zurück in Bytes für die Datei
                Dim dataToSave As Byte() = Convert.FromBase64String(TextBox1.Text)
                File.WriteAllBytes(sfd.FileName, dataToSave)
                MessageBox.Show("Datei gespeichert!")
            Catch
                MessageBox.Show("Bitte erst auf 'Kryptieren' klicken, bevor du speicherst.")
            End Try
        End If
    End Sub

    ''' <summary>
    ''' Button 4: Öffnen
    ''' Lädt eine .crypt Datei und zeigt den Inhalt als Base64 in der TextBox an.
    ''' </summary>
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim ofd As New OpenFileDialog With {
            .Filter = "Crypt Files (*.crypt)|*.crypt",
            .Title = "Verschlüsselte Datei öffnen"
        }

        If ofd.ShowDialog() = DialogResult.OK Then
            Try
                Dim fileData As Byte() = File.ReadAllBytes(ofd.FileName)
                ' Zeige den geladenen Inhalt als Base64 an (bereit zum Dekryptieren)
                TextBox1.Text = Convert.ToBase64String(fileData)
            Catch ex As Exception
                MessageBox.Show("Fehler beim Laden: " & ex.Message)
            End Try
        End If
    End Sub

End Class