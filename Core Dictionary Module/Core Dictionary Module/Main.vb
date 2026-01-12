Imports System.IO
Imports System.Windows.Forms

Public Class Main

    'Manual read function is only used for special applications, not for standard XELA applications
    'Issues arrising using manual read functions in standard applications will not be supported and must be fixed by the creator of the application
    'A Object based read function is not provided for manual read functions to avoid possible issues

    Public Function MRead_String(path As String) As String 'Manual Path selection
        If path.EndsWith(".word") Then
            Dim result As String = File.ReadAllText(path)
            Return result
        End If

        Return String.Empty
    End Function

    Public Function MRead_Integer(path As String) As Integer 'Manual Path selection
        If path.EndsWith(".val") Then
            Dim content As String = File.ReadAllText(path)
            Dim result As Integer
            If Integer.TryParse(content, result) Then
                Return result
            End If
        End If
        Return 0
    End Function

    Public Function MRead_Boolean(path As String) As Boolean 'Manual Path selection
        If path.EndsWith(".sta") Then
            Dim content As String = File.ReadAllText(path)
            Dim result As Boolean
            If Boolean.TryParse(content, result) Then
                Return result
            End If
        End If
        Return False
    End Function

    'Auto read function for standard applications
    'Deleting important data can result in an exception being returned
    'Please be aware of this when using the Auto Read function
    'Issues could arrise because of the result being an object, manual fix might be required
    'See the readme for a solution

    Public Enum ObjectLevel
        SysApps
        SysExt
        Root
    End Enum

    Public Function AutoRead(data As String, type As ObjectType) As Object 'Auto Path selection, for standard applications
        Dim systemPath As String = MRead_String("C:\KAVN\%mela.arc%\bin_path.word") 'Use MELA system path as standard
        Dim path As String = String.Empty

        ' Determine the correct path based on the ObjectType
        If ObjectLevel.SysApps Then
            path = systemPath & "SYSTEM_APPLICATION\"
        ElseIf ObjectLevel.SysExt Then
            path = systemPath & "SYSTEM_EXTENSION\"
        ElseIf ObjectLevel.Root Then
            path = systemPath
        End If

        ' Attempt to read the file based on its extension
        Try
            If File.Exists(IO.Path.Combine(path, data & ".word")) Then
                Return MRead_String(IO.Path.Combine(path, data & ".word"))
            ElseIf File.Exists(IO.Path.Combine(path, data & ".val")) Then
                Return MRead_Integer(IO.Path.Combine(path, data & ".val"))
            ElseIf File.Exists(IO.Path.Combine(path, data & ".sta")) Then
                Return MRead_Boolean(IO.Path.Combine(path, data & ".sta"))
            Else
                Return New FileNotFoundException("file '" & path & data & "' not found")
            End If
        Catch ex As Exception
            Return ex
        End Try
    End Function


    ' Manual write function is only used for special applications, not for standard XELA applications
    ' Issues arrising using manual write functions in standard applications will not be supported and must be fixed by the creator of the application
    ' A Object based write function is not provided for manual write functions to avoid possible issues

    Public Function MWrite_String(path As String, content As String) As Boolean
        If path.EndsWith(".word") Then
            Try
                File.WriteAllText(path, content)
                Return True
            Catch ex As Exception
                Return False
            End Try
        Else
            Return False
        End If
    End Function

    Public Function MWrite_Integer(path As String, content As Integer) As Boolean
        If path.EndsWith(".val") Then
            Try
                File.WriteAllText(path, content.ToString())
                Return True
            Catch ex As Exception
                Return False
            End Try
        Else
            Return False
        End If
    End Function

    Public Function MWrite_Boolean(path As String, content As Boolean) As Boolean
        If path.EndsWith(".sta") Then
            Try
                File.WriteAllText(path, content.ToString())
                Return True
            Catch ex As Exception
                Return False
            End Try
        Else
            Return False
        End If
    End Function

    ' Auto write function for standard applications
    ' Issues could arrise because of the content being an object, manual fix might be required
    ' See the readme for a solution

    Public Function AutoWrite(data As String, content As Object, type As ObjectType) As Boolean
        Dim systemPath As String = MRead_String("C:\KAVN\%mela.arc%\bin_path.word") 'Use MELA system path as standard
        Dim path As String = String.Empty

        ' Determine the correct path based on the ObjectType
        If ObjectLevel.SysApps Then
            path = systemPath & "SYSTEM_APPLICATION\"
        ElseIf ObjectLevel.SysExt Then
            path = systemPath & "SYSTEM_EXTENSION\"
        ElseIf ObjectLevel.Root Then
            path = systemPath
        End If

        ' Attempt to write the file based on the content type
        Try
            If TypeOf content Is Boolean Then
                Return MWrite_String(IO.Path.Combine(path, data & ".word"), CType(content, String))
            ElseIf TypeOf content Is Integer Then
                Return MWrite_Integer(IO.Path.Combine(path, data & ".val"), CType(content, Integer))
            ElseIf TypeOf content Is String Then
                Return MWrite_Boolean(IO.Path.Combine(path, data & ".sta"), CType(content, Boolean))
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try

    End Function


    ' Create an object (file) based on its type
    ' Endings are automatically added based on the object type
    ' Content is optional and can be left empty

    Public Enum ObjectType
        StringObject
        IntegerObject
        BooleanObject
    End Enum

    Public Function CreateObject(objectName As String, path As String, objectType As ObjectType, Optional content As Object = "") As Boolean
        If objectType = ObjectType.BooleanObject Then
            Dim fullPath As String = IO.Path.Combine(path, objectName & ".sta")
            If content = Not Nothing Then
                Return MWrite_Boolean(fullPath, CType(content, Boolean))
            End If
        ElseIf objectType = ObjectType.IntegerObject Then
            Dim fullPath As String = IO.Path.Combine(path, objectName & ".val")
            If content = Not Nothing Then
                Return MWrite_Integer(fullPath, CType(content, Integer))
            End If
        ElseIf objectType = ObjectType.StringObject Then
            Dim fullPath As String = IO.Path.Combine(path, objectName & ".word")
            If content = Not Nothing Then
                Return MWrite_String(fullPath, CType(content, String))
            End If
        Else
            Return False
        End If

        Return False
    End Function

    ' Rename an object (file) based on its type
    ' Files named incorrectly or with the wrong extension will not be renamed


    Public Function RenameObject(currentName As String, newName As String) As Boolean
        Try
            If currentName.EndsWith(".sta") Then
                If File.Exists(currentName) Then
                    Dim content As Object = File.ReadAllText(currentName)
                    If TypeOf content Is Boolean Then
                        Try
                            Rename(currentName, newName)
                            Return True
                        Catch ex As Exception
                            Return False
                        End Try
                    End If
                End If
            ElseIf currentName.EndsWith(".val") Then
                If File.Exists(currentName) Then
                    Dim content As Object = File.ReadAllText(currentName)
                    If TypeOf content Is Integer Then
                        Try
                            Rename(currentName, newName)
                            Return True
                        Catch ex As Exception
                            Return False
                        End Try
                    End If
                End If
            ElseIf currentName.EndsWith(".word") Then
                If File.Exists(currentName) Then
                    Dim content As Object = File.ReadAllText(currentName)
                    If TypeOf content Is String Then
                        Try
                            Rename(currentName, newName)
                            Return True
                        Catch ex As Exception
                            Return False
                        End Try
                    End If
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try

        Return False
    End Function


    ' Get a tree of all files and directories within a specified path


    Public Sub Tree(tv As TreeView, rootPath As String)
        ' Hole alle Pfade über deine bestehende Funktion
        Dim allPaths As List(Of String) = MGetTree(rootPath)

        tv.Nodes.Clear()

        ' Falls der Pfad nicht existiert, abbrechen
        If Not Directory.Exists(rootPath) Then Return

        ' Root-Knoten erstellen
        Dim rootNode As New TreeNode(IO.Path.GetFileName(rootPath.TrimEnd("\"c)))
        rootNode.Tag = rootPath ' Wir speichern den echten Pfad im Tag
        tv.Nodes.Add(rootNode)

        ' Liste der erlaubten Endungen
        Dim allowedExtensions As String() = {".word", ".val", ".sta"}

        For Each fullPath As String In allPaths
            ' PRÜFUNG: Ist es ein Verzeichnis ODER eine Datei mit der richtigen Endung?
            Dim isDirectory As Boolean = Directory.Exists(fullPath)
            Dim extension As String = IO.Path.GetExtension(fullPath).ToLower()

            ' Nur fortfahren, wenn es ein Ordner ist oder die Endung passt
            If Not isDirectory AndAlso Not allowedExtensions.Contains(extension) Then
                Continue For
            End If

            ' Berechne den relativen Pfad für die Baumstruktur
            Dim relativePath As String = fullPath.Substring(rootPath.Length).TrimStart("\"c)
            If String.IsNullOrEmpty(relativePath) Then Continue For

            Dim parts As String() = relativePath.Split("\"c)
            Dim currentNode As TreeNode = rootNode

            For Each part As String In parts
                ' Suche, ob der Knoten auf dieser Ebene bereits existiert
                Dim existingNode As TreeNode = Nothing
                For Each node As TreeNode In currentNode.Nodes
                    If node.Text = part Then
                        existingNode = node
                        Exit For
                    End If
                Next

                If existingNode Is Nothing Then
                    ' Falls nicht vorhanden, neu anlegen
                    Dim newNode As New TreeNode(part)
                    newNode.Tag = fullPath ' Pfad für späteren Zugriff speichern
                    currentNode.Nodes.Add(newNode)
                    currentNode = newNode
                Else
                    ' Falls vorhanden, tiefer gehen
                    currentNode = existingNode
                End If
            Next
        Next

        rootNode.Expand()
    End Sub

    ' Deine bestehende MGetTree Funktion
    Public Function MGetTree(path As String) As List(Of String)
        Dim result As New List(Of String)
        If Directory.Exists(path) Then
            ' Sortierte Liste zurückgeben (hilft beim Aufbau des Baums)
            Dim directories = Directory.GetDirectories(path, "*", SearchOption.AllDirectories).OrderBy(Function(x) x)
            Dim files = Directory.GetFiles(path, "*", SearchOption.AllDirectories).OrderBy(Function(x) x)
            result.AddRange(directories)
            result.AddRange(files)
        End If
        Return result
    End Function

    ' XELA support is not provided for AutoGetTree function
    ' A new tree for MELA and XELA is being designed

    ''' <summary>
    ''' Ermittelt automatisch den Pfad basierend auf dem ObjectLevel und befüllt das TreeView.
    ''' </summary>
    Public Sub AutoTree(tv As TreeView, level As ObjectLevel)
        ' 1. Pfad ermitteln (Logik aus AutoGetTree übernommen)
        Dim systemPath As String = MRead_String("C:\KAVN\%mela.arc%\bin_path.word")
        Dim targetPath As String = String.Empty

        Select Case level
            Case ObjectLevel.SysApps
                targetPath = systemPath & "SYSTEM_APPLICATION\"
            Case ObjectLevel.SysExt
                targetPath = systemPath & "SYSTEM_EXTENSION\"
            Case ObjectLevel.Root
                targetPath = systemPath
        End Select

        ' 2. Falls der Pfad gültig ist, die bereits erstellte FillTreeView Logik nutzen
        If Not String.IsNullOrEmpty(targetPath) AndAlso Directory.Exists(targetPath) Then
            Tree(tv, targetPath)
        Else
            tv.Nodes.Clear()
            tv.Nodes.Add("Pfad nicht gefunden: " & targetPath)
        End If
    End Sub

    ' Korrigierte AutoGetTree Funktion
    Public Function AutoGetTree(level As ObjectLevel) As List(Of String)
        Dim systemPath As String = MRead_String("C:\KAVN\%mela.arc%\bin_path.word")
        Dim path As String = String.Empty

        ' Nutze Select Case für Enums - das ist sauberer und vermeidet Fehler
        Select Case level
            Case ObjectLevel.SysApps
                path = systemPath & "SYSTEM_APPLICATION\"
            Case ObjectLevel.SysExt
                path = systemPath & "SYSTEM_EXTENSION\"
            Case ObjectLevel.Root
                path = systemPath
        End Select

        Return MGetTree(path)
    End Function

End Class