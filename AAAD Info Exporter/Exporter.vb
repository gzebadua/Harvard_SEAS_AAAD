Imports System.Xml
Imports System.IO
Imports System.Text
Imports System.Globalization

Public Class Exporter


    Private Sub exportForms()

        btnExit.Enabled = False

        Try

            Dim inputPath As String = "Z:/AAAD Forms v3/" 'Mounted a drive to the AAAD sharepoint directly... although that means that you shouldn't logoff from that machine ever...
            Dim fileName As String

            Dim fileEntries As String() = Directory.GetFiles(inputPath)

            For Each fileName In fileEntries

                rtb.AppendText("Reading " & inputPath & "\" & fileName & Chr(13))

                Dim reader As StreamReader = New StreamReader(inputPath & "\" & fileName & ".xml") 'Don't know if inputPath is with filename or not... yet.
                Dim line As String = reader.ReadToEnd()
                reader.Dispose()

                Dim exportQuery As String = ""

                If getSQLQueryAsBoolean("SELECT ID from DW_AAAD.dummyTableName WHERE ID = " & getValue(line, "AutoID")) = False Then

                    ' "New" Person (not in db)
                    exportQuery = "INSERT INTO DW_AAAD.dummyTableName VALUES('"  'Don't forget the update datetime

                Else

                    'Person already in database
                    Dim fileDate As Date = System.IO.File.GetLastWriteTime(inputPath).ToShortDateString()

                    Dim dbDate As Date
                    Dim formats As String() = {"yyyyMMdd", "MM/dd/yy"}
                    Date.TryParseExact(getSQLQueryAsString("SELECT last_date from DW_AAAD.dummyTableName WHERE ID = " & getValue(line, "AutoID")), formats, CultureInfo.CurrentCulture, DateTimeStyles.None, dbDate)

                    If fileDate.CompareTo(dbDate) > 0 Then 'Update needed

                        exportQuery = "UPDATE DW_AAAD.dummyTableName SET    WHERE ID = " & getValue(line, "AutoID")

                    End If

                End If

                executeSQLCommand(exportQuery)

                line = ""

            Next fileName

        Catch ex As Exception

        End Try

        btnExit.Enabled = True

    End Sub


    Private Function getValue(ByRef line As String, ByVal field As String, Optional ByVal defaultValue As String = "") As String

        If line.Contains("<" & field & "></") = True Then 'empty

            Return defaultValue

        ElseIf line.Contains("<" & field & ">") = True And line.Contains("<" & field & "></") = False Then 'Contains something

            Dim charsToCut As Integer = line.IndexOf("</" & field & ">") - line.IndexOf("<" & field & ">") - (field.Length + 2)

            Return line.Substring(line.IndexOf("<" & field & ">") + field.Length + 2, charsToCut)

        End If

    End Function


    Private Sub UpgradingForms_Shown(sender As Object, e As System.EventArgs) Handles Me.Shown

        'Application.DoEvents()

    End Sub


    Private Sub UpgradingForms_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        exportForms()
        System.Environment.Exit(0)

    End Sub


    Private Sub btnExit_Click(sender As System.Object, e As System.EventArgs) Handles btnExit.Click

        System.Environment.Exit(0)

    End Sub


    Private Sub btnUpdate_Click(sender As System.Object, e As System.EventArgs) Handles btnUpdate.Click

        exportForms()

    End Sub


End Class
