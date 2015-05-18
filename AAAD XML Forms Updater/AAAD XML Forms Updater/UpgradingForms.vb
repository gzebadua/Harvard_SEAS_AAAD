Imports System.Xml
Imports System.IO
Imports System.Text

Public Class UpgradingForms


    Private Sub UpgradingForms_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub


    Private Sub UpgradingForms_Shown(sender As Object, e As System.EventArgs) Handles Me.Shown

        Application.DoEvents()
        'updateFormsv5()

    End Sub


    Private Sub btnUpdate_Click(sender As System.Object, e As System.EventArgs) Handles btnUpdate.Click

        updateFormsv5()

    End Sub


    Private Sub formatOneLinedXML_Click(sender As System.Object, e As System.EventArgs) Handles formatOneLinedXML.Click

        btnExit.Enabled = False

        Dim iterator As Integer = 0
        Dim endOfElement As Integer = 0

        Try

            inputFolder.ShowNewFolderButton = False
            inputFolder.Description = "Please choose the folder where the outdated forms are"

            If inputFolder.ShowDialog() = DialogResult.OK Then

                outputFolder.ShowNewFolderButton = False
                outputFolder.Description = "Please choose the folder where the UPDATED forms will be saved"

                If outputFolder.ShowDialog() = DialogResult.OK Then

                    'Dim dirInfo As New DirectoryInfo(inputFolder.SelectedPath)

                    'For Each fi As FileInfo In dirInfo.GetFiles()

                    'Dim filename As String = fi.Name
                    Dim filename As String = "831.xml"

                    Try

                        Dim reader As StreamReader = New StreamReader(inputFolder.SelectedPath & "\" & filename)

                        Dim line As String = reader.ReadToEnd()
                        reader.Dispose()

                        line = line.Replace("xml:lang=""en-US"">", "xml:lang=""en-US"">" & vbCrLf)

                        For i As Integer = 1 To line.Length - 1

                            iterator = i

                            If line(i) = "?" And line(i + 1) = ">" Then

                                line = line.Substring(0, i + 2) & vbCrLf & line.Substring(i + 2)
                                i = i + 2

                            ElseIf line(i) = "<" And line(i + 1) = "/" Then

                                For j = i To line.Length - 1

                                    If line(j) = ">" Then
                                        endOfElement = j + 1
                                        Exit For
                                    End If

                                Next j

                                line = line.Substring(0, endOfElement) & vbCrLf & line.Substring(endOfElement)
                                i = endOfElement


                            ElseIf line(i) = "/" And line(i - 1) = " " And line(i + 1) = ">" Then

                                line = line.Substring(0, i + 2) & vbCrLf & line.Substring(i + 2)
                                i = i + 2

                            End If

                        Next i

                        'Repeating for the last part because I don't know why it quits the For before finishing...

                        For i As Integer = 7315 To line.Length - 1

                            iterator = i

                            If line(i) = "<" And line(i + 1) = "/" Then

                                For j = i To line.Length - 1

                                    If line(j) = ">" Then
                                        endOfElement = j + 1
                                        Exit For
                                    End If

                                Next j

                                line = line.Substring(0, endOfElement) & vbCrLf & line.Substring(endOfElement)
                                i = endOfElement


                            ElseIf line(i) = "/" And line(i - 1) = " " And line(i + 1) = ">" Then

                                line = line.Substring(0, i + 2) & vbCrLf & line.Substring(i + 2)
                                i = i + 2

                            End If

                        Next i


                        Dim writer As New StreamWriter(outputFolder.SelectedPath & "\" & filename)

                        writer.Write(line)

                        writer.Dispose()

                        line = ""
                        filename = ""

                    Catch ex As Exception

                        Dim test As String
                        test = "error"

                    End Try

                    'Next fi

                    'End If ' Ends the OK of DirectoryBrowser

                End If ' Ends the OK of FileBrowser

            End If ' Ends the OK of DirectoryBrowser

        Catch ex2 As Exception

        End Try

        btnExit.Enabled = True

    End Sub


    Private Sub btnExit_Click(sender As System.Object, e As System.EventArgs) Handles btnExit.Click

        System.Environment.Exit(0)

    End Sub


    Private Sub updateFormsv5()

        btnExit.Enabled = False

        Try

            inputFolder.ShowNewFolderButton = False
            inputFolder.Description = "Please choose the folder where the outdated forms are"

            If inputFolder.ShowDialog() = DialogResult.OK Then

                outputFolder.ShowNewFolderButton = False
                outputFolder.Description = "Please choose the folder where the UPDATED forms will be saved"

                If outputFolder.ShowDialog() = DialogResult.OK Then

                    Dim dirInfo As New DirectoryInfo(inputFolder.SelectedPath)

                    For Each fi As FileInfo In dirInfo.GetFiles()

                        Dim filename As String = fi.Name

                        Try

                            rtb.AppendText("Reading " & inputFolder.SelectedPath & "\" & filename & Chr(13))

                            Dim reader As StreamReader = New StreamReader(inputFolder.SelectedPath & "\" & filename)

                            Dim writer As New StreamWriter(outputFolder.SelectedPath & "\" & filename)

                            Dim line As String = reader.ReadToEnd()
                            reader.Dispose()

                            Dim outputStringBuilder As New StringBuilder

                            outputStringBuilder.Append("<?xml version=""1.0"" encoding=""utf-8""?>")
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<?mso-infoPathSolution name=""urn:schemas-microsoft-com:office:infopath:AAAD-Forms-v3:-myXSD-2009-11-02T17-20-04"" solutionVersion=""1.0.0.2626"" productVersion=""12.0.0.0"" PIVersion=""1.0.0.0"" href=""https://sp.seas.harvard.edu/sites/eforms/AAAD%20Forms%20v3/Forms/template.xsn""?>")
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<?mso-application progid=""InfoPath.Document"" versionProgid=""InfoPath.Document.2""?>")
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:myFields xmlns:my=""http://schemas.microsoft.com/office/infopath/2003/myXSD/2009-11-02T17:20:04"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:dfs=""http://schemas.microsoft.com/office/infopath/2003/dataFormSolution"" xmlns:xhtml=""http://www.w3.org/1999/xhtml"" xmlns:pc=""http://schemas.microsoft.com/office/infopath/2007/PartnerControls"" xmlns:s=""uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882"" xmlns:dt=""uuid:C2F41010-65B3-11d1-A29F-00AA00C14882"" xmlns:rs=""urn:schemas-microsoft-com:rowset"" xmlns:z=""#RowsetSchema"" xmlns:d=""http://schemas.microsoft.com/office/infopath/2003/ado/dataFields"" xmlns:xd=""http://schemas.microsoft.com/office/infopath/2003"" xml:lang=""en-US"">")
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:hr_aspire_posting_title")
                            outputStringBuilder.Append(importValues(line, "<my:hr_aspire_posting_title>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:finance_costing_approved")
                            outputStringBuilder.Append(importValues(line, "<my:finance_costing_approved>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:facilities_furniture_redistributed")
                            outputStringBuilder.Append(importValues(line, "<my:facilities_furniture_redistributed>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:pi")
                            Dim pi As String = importValues(line, "<my:pi>", True)
                            outputStringBuilder.Append(pi)
                            If pi <> "" And pi <> " />" Then
                                outputStringBuilder.Append("</my:pi>")
                            End If
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:it_account_created")
                            outputStringBuilder.Append(importValues(line, "<my:it_account_created>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:it_account_disabled")
                            outputStringBuilder.Append(importValues(line, "<my:it_account_disabled>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:it_equipment_reclaimed")
                            outputStringBuilder.Append(importValues(line, "<my:it_equipment_reclaimed>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:sponsresrch_costing_approval_for_academic_or_staff_appointment")
                            outputStringBuilder.Append(importValues(line, "<my:sponsresrch_costing_approval_for_academic_or_staff_appointment>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:sponsresrch_costing_approved")
                            outputStringBuilder.Append(importValues(line, "<my:sponsresrch_costing_approved>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:sponsresrch_gmas_access_granted")
                            outputStringBuilder.Append(importValues(line, "<my:sponsresrch_gmas_access_granted>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:sponsresrch_pi_award_transferred")
                            outputStringBuilder.Append(importValues(line, "<my:sponsresrch_pi_award_transferred>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:sponsresrch_pi_award_terminated")
                            outputStringBuilder.Append(importValues(line, "<my:sponsresrch_pi_award_terminated>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:studaffrs_notified_by_area_admins_of_faculty_arrival_date")
                            outputStringBuilder.Append(importValues(line, "<my:studaffrs_notified_by_area_admins_of_faculty_arrival_date>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:studaffrs_grad_students_assigned_to_new_faculty")
                            outputStringBuilder.Append(importValues(line, "<my:studaffrs_grad_students_assigned_to_new_faculty>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:studaffrs_keys_returned_to_facilities")
                            outputStringBuilder.Append(importValues(line, "<my:studaffrs_keys_returned_to_facilities>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:studaffrs_id_deactivated_and_returned")
                            outputStringBuilder.Append(importValues(line, "<my:studaffrs_id_deactivated_and_returned>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:studaffrs_payroll_notified_of_termination")
                            outputStringBuilder.Append(importValues(line, "<my:studaffrs_payroll_notified_of_termination>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:studaffrs_area_admins_notified_of_departure")
                            outputStringBuilder.Append(importValues(line, "<my:studaffrs_area_admins_notified_of_departure>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:procurement_iprocurement_authorization_form_completed_by_pi")
                            outputStringBuilder.Append(importValues(line, "<my:procurement_iprocurement_authorization_form_completed_by_pi>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:procurement_iprocurement_training_scheduled")
                            outputStringBuilder.Append(importValues(line, "<my:procurement_iprocurement_training_scheduled>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:procurement_pcard_training_scheduled")
                            outputStringBuilder.Append(importValues(line, "<my:procurement_pcard_training_scheduled>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:procurement_corporate_card_training_scheduled")
                            outputStringBuilder.Append(importValues(line, "<my:procurement_corporate_card_training_scheduled>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:procurement_notified_by_sponsored_research_of_pi_departure")
                            outputStringBuilder.Append(importValues(line, "<my:procurement_notified_by_sponsored_research_of_pi_departure>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:facilities_comments")
                            outputStringBuilder.Append(importValues(line, "<my:facilities_comments>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:areaadmins_studaffrs_notified_of_faculty_arrival")
                            outputStringBuilder.Append(importValues(line, "<my:areaadmins_studaffrs_notified_of_faculty_arrival>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:areaadmins_appointing_office_notified_of_departure")
                            outputStringBuilder.Append(importValues(line, "<my:areaadmins_appointing_office_notified_of_departure>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:areaadmins_faculty_support_reassigned")
                            outputStringBuilder.Append(importValues(line, "<my:areaadmins_faculty_support_reassigned>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:info_url")
                            outputStringBuilder.Append(importValues(line, "<my:info_url>", False, "").Replace("OpenIn=Browser", "Source=https%3A%2F%2Fsp%2Eseas%2Eharvard%2Eedu%2Fsites%2Feforms%2FAAAD%2520Forms%2520v3%2FForms%2FAllItems%2Easpx&amp;DefaultItemOpen=1"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:position")
                            Dim position As String = importValues(line, "<my:position>", True)
                            position = position.Replace(" (computer eligible)", "").Replace(" (not computer eligible)", "")
                            outputStringBuilder.Append(position)
                            If position <> "" And position <> " />" Then
                                outputStringBuilder.Append("</my:position>")
                            End If
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:alternate_email_address")
                            outputStringBuilder.Append(importValues(line, "<my:alternate_email_address>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:supervisor")
                            Dim supervisor As String = importValues(line, "<my:supervisor>", True)
                            outputStringBuilder.Append(supervisor)
                            If supervisor <> "" And supervisor <> " />" Then
                                outputStringBuilder.Append("</my:supervisor>")
                            End If
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:huid")
                            outputStringBuilder.Append(importValues(line, "<my:huid>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:email_address")
                            Dim email_address As String = importValues(line, "<my:email_address>", True, "")
                            outputStringBuilder.Append(email_address)
                            If email_address <> "" And email_address <> " />" Then
                                outputStringBuilder.Append("</my:email_address>")
                            End If
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:hr_req_number")
                            outputStringBuilder.Append(importValues(line, "<my:hr_req_number>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:office_phone_number")
                            outputStringBuilder.Append(importValues(line, "<my:office_phone_number>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:it_pc_delivered")
                            outputStringBuilder.Append(importValues(line, "<my:it_pc_delivered>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:it_pc_set_up")
                            outputStringBuilder.Append(importValues(line, "<my:it_pc_set_up>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:hr_unit")
                            Dim hr_unit As String = importValues(line, "<my:hr_unit>", True, "")
                            hr_unit = hr_unit.Replace("	", "")
                            outputStringBuilder.Append(hr_unit)
                            If hr_unit <> "" And hr_unit <> " />" Then
                                outputStringBuilder.Append("</my:hr_unit>")
                            End If
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:office_building")
                            outputStringBuilder.Append(importValues(line, "<my:office_building>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:office_room_number")
                            outputStringBuilder.Append(importValues(line, "<my:office_room_number>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:first_name")
                            outputStringBuilder.Append(importValues(line, "<my:first_name>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:middle_name")
                            outputStringBuilder.Append(importValues(line, "<my:middle_name>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:last_name")
                            outputStringBuilder.Append(importValues(line, "<my:last_name>"))
                            outputStringBuilder.AppendLine()
                            appendDateElement(outputStringBuilder, line, "my:start_date")
                            outputStringBuilder.AppendLine()
                            appendDateElement(outputStringBuilder, line, "my:termination_date")
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:comments")
                            outputStringBuilder.Append(importValues(line, "<my:comments>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:acdaffrs_area")
                            Dim area As String = importValues(line, "<my:acdaffrs_area>", True)

                            If pi <> "" And area = "" Then

                                Select Case pi

                                    Case "ayacoby"
                                        area = "Applied Physics"
                                    Case "aziz"
                                        area = "Materials Science and Mechanical Engineering"
                                    Case "bertoldi"
                                        area = "Materials Science and Mechanical Engineering"
                                    Case "bmw"
                                        area = "Applied Physics"
                                    Case "bossert"
                                        area = "Applied Math"
                                    Case "brenner"
                                        area = "Applied Math"
                                    Case "brockett"
                                        area = "Electrical Engineering"
                                    Case "camurray"
                                        area = "Applied Physics"
                                    Case "capasso"
                                        area = "Applied Physics"
                                    Case "cfriend"
                                        area = "Materials Science and Mechanical Engineering"
                                    Case "chong"
                                        area = "Computer Science"
                                    Case "clarke"
                                        area = "Materials Science and Mechanical Engineering"
                                    Case "clieber"
                                        area = "Materials Science and Mechanical Engineering"
                                    Case "cluzel"
                                        area = "Applied Physics"
                                    Case "cshen"
                                        area = "Computer Science"
                                    Case "davidcox"
                                        area = "Computer Science"
                                    Case "dbrooks"
                                        area = "Computer Science"
                                    Case "dcb"
                                        area = "Applied Physics"
                                    Case "dedwards"
                                        area = "Bioengineering"
                                    Case "djj"
                                        area = "Environmental Science and Engineering"
                                    Case "dkeith"
                                        area = "Applied Physics"
                                    Case "dneedle"
                                        area = "Applied Physics"
                                    Case "donhee"
                                        area = "Electrical Engineering"
                                    Case "ehu"
                                        area = "Electrical Engineering"
                                    Case "eli"
                                        area = "Environmental Science and Engineering"
                                    Case "erez"
                                        area = "Computer Science"
                                    Case "farrell"
                                        area = "Environmental Science and Engineering"
                                    Case "golovchenko"
                                        area = "Applied Physics"
                                    Case "greg"
                                        area = "Computer Science"
                                    Case "grosz"
                                        area = "Computer Science"
                                    Case "guyeon"
                                        area = "Electrical Engineering"
                                    Case "hau"
                                        area = "Applied Physics"
                                    Case "howe"
                                        area = "Materials Science and Mechanical Engineering"
                                    Case "htk"
                                        area = "Computer Science"
                                    Case "hutchins"
                                        area = "Materials Science and Mechanical Engineering"
                                    Case "ingber"
                                        area = "Bioengineering"
                                    Case "jaiz"
                                        area = "Materials Science and Mechanical Engineering"
                                    Case "jalewis"
                                        area = "Materials Science and Mechanical Engineering"
                                    Case "janders"
                                        area = "Environmental Science and Engineering"
                                    Case "jbriscoe"
                                        area = "Environmental Science and Engineering"
                                    Case "jlogan"
                                        area = "Environmental Science and Engineering"
                                    Case "jvlassak"
                                        area = "Materials Science and Mechanical Engineering"
                                    Case "jwmunger"
                                        area = "Environmental Science and Engineering"
                                    Case "kamckinney"
                                        area = "Environmental Science and Engineering"
                                    Case "kaxiras"
                                        area = "Applied Physics"
                                    Case "kcrozier"
                                        area = "Electrical Engineering"
                                    Case "kgajos"
                                        area = "Computer Science"
                                    Case "kkparker"
                                        area = "Bioengineering"
                                    Case "kohler"
                                        area = "Computer Science"
                                    Case "lewis"
                                        area = "Computer Science"
                                    Case "lm"
                                        area = "Applied Math"
                                    Case "loncar"
                                        area = "Electrical Engineering"
                                    Case "margo"
                                        area = "Computer Science"
                                    Case "mas"
                                        area = "Bioengineering"
                                    Case "mazur"
                                        area = "Applied Physics"
                                    Case "mbm"
                                        area = "Environmental Science and Engineering"
                                    Case "michaelm"
                                        area = "Computer Science"
                                    Case "mickley"
                                        area = "Environmental Science and Engineering"
                                    Case "mitchell"
                                        area = "Environmental Science and Engineering"
                                    Case "mooneyd"
                                        area = "Bioengineering"
                                    Case "navin"
                                        area = "Electrical Engineering"
                                    Case "nelson"
                                        area = "Applied Physics"
                                    Case "njoshi"
                                        area = "Bioengineering"
                                    Case "parkes"
                                        area = "Computer Science"
                                    Case "pershan"
                                        area = "Applied Physics"
                                    Case "pfister"
                                        area = "Computer Science"
                                    Case "rad"
                                        area = "Computer Science"
                                    Case "rgordon"
                                        area = "Materials Science and Mechanical Engineering"
                                    Case "rice"
                                        area = "Materials Science and Mechanical Engineering"
                                    Case "rjwood"
                                        area = "Materials Science and Mechanical Engineering"
                                    Case "rmadix"
                                        area = "Environmental Science and Engineering"
                                    Case "rogers"
                                        area = "Environmental Science and Engineering"
                                    Case "rpa"
                                        area = "Computer Science"
                                    Case "salil"
                                        area = "Computer Science"
                                    Case "schrag"
                                        area = "Environmental Science and Engineering"
                                    Case "sharad"
                                        area = "Applied Physics"
                                    Case "shieber"
                                        area = "Computer Science"
                                    Case "shriram"
                                        area = "Materials Science and Mechanical Engineering"
                                    Case "sjg"
                                        area = "Computer Science"
                                    Case "smartin"
                                        area = "Environmental Science and Engineering"
                                    Case "spaepen"
                                        area = "Materials Science and Mechanical Engineering"
                                    Case "suo"
                                        area = "Materials Science and Mechanical Engineering"
                                    Case "swofsy"
                                        area = "Environmental Science and Engineering"
                                    Case "ttwu"
                                        area = "Applied Physics"
                                    Case "vahid"
                                        area = "Electrical Engineering"
                                    Case "valiant"
                                        area = "Computer Science"
                                    Case "vecitis"
                                        area = "Environmental Science and Engineering"
                                    Case "venky"
                                        area = "Applied Physics"
                                    Case "vnm"
                                        area = "Applied Physics"
                                    Case "waldo"
                                        area = "Computer Science"
                                    Case "walsh"
                                        area = "Materials Science and Mechanical Engineering"
                                    Case "weitz"
                                        area = "Applied Physics"
                                    Case "woody"
                                        area = "Electrical Engineering"
                                    Case "yiling"
                                        area = "Computer Science"
                                    Case "yuelu"
                                        area = "Electrical Engineering"
                                    Case "zickler"
                                        area = "Electrical Engineering"
                                    Case "zittrain"
                                        area = "Computer Science"
                                    Case "zkuang"
                                        area = "Environmental Science and Engineering"

                                End Select

                            End If

                            outputStringBuilder.Append(area)
                            If area <> "" And area <> " />" Then
                                outputStringBuilder.Append("</my:acdaffrs_area>")
                            End If
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:seas_username")
                            Dim seas_username As String = importValues(line, "<my:seas_username>", True)
                            outputStringBuilder.Append(seas_username)
                            If seas_username <> "" And seas_username <> " />" Then
                                outputStringBuilder.Append("</my:seas_username>")
                            End If
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:office_phone_jack_number")
                            outputStringBuilder.Append(importValues(line, "<my:office_phone_jack_number>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:office_phone_jack_activated")
                            outputStringBuilder.Append(importValues(line, "<my:office_phone_jack_activated>", False, "false"))
                            outputStringBuilder.AppendLine()
                            appendDateElement(outputStringBuilder, line, "my:expected_pc_arrival_date")
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:pc_ordered")
                            outputStringBuilder.Append(importValues(line, "<my:pc_ordered>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:copier_code")
                            outputStringBuilder.Append(importValues(line, "<my:copier_code>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:assigned_mailbox")
                            outputStringBuilder.Append(importValues(line, "<my:assigned_mailbox>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:office_key_provided")
                            outputStringBuilder.Append(importValues(line, "<my:office_key_provided>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:swipe_access_activated")
                            outputStringBuilder.Append(importValues(line, "<my:swipe_access_activated>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:assigned_faculty_support")
                            outputStringBuilder.Append(importValues(line, "<my:assigned_faculty_support>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:safety_general_training_complete")
                            outputStringBuilder.Append(importValues(line, "<my:safety_general_training_complete>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:safety_general_training_na")
                            outputStringBuilder.Append(importValues(line, "<my:safety_general_training_na>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:safety_hazardous_waste_training_complete")
                            outputStringBuilder.Append(importValues(line, "<my:safety_hazardous_waste_training_complete>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:safety_hazardous_waste_training_na")
                            outputStringBuilder.Append(importValues(line, "<my:safety_hazardous_waste_training_na>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:safety_bio_training_complete")
                            outputStringBuilder.Append(importValues(line, "<my:safety_bio_training_complete>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:safety_bio_training_na")
                            outputStringBuilder.Append(importValues(line, "<my:safety_bio_training_na>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:safety_radiation_training_complete")
                            outputStringBuilder.Append(importValues(line, "<my:safety_radiation_training_complete>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:safety_radiation_training_na")
                            outputStringBuilder.Append(importValues(line, "<my:safety_radiation_training_na>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:safety_laser_training_complete")
                            outputStringBuilder.Append(importValues(line, "<my:safety_laser_training_complete>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:safety_laser_training_na")
                            outputStringBuilder.Append(importValues(line, "<my:safety_laser_training_na>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:safety_removed_from_training_lists")
                            outputStringBuilder.Append(importValues(line, "<my:safety_removed_from_training_lists>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:nick_name")
                            outputStringBuilder.Append(importValues(line, "<my:nick_name>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:hire_type")
                            outputStringBuilder.Append(importValues(line, "<my:hire_type>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:incumbent_name")
                            outputStringBuilder.Append(importValues(line, "<my:incumbent_name>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:Activity_Stream_Textbox")
                            outputStringBuilder.Append(importValues(line, "<my:Activity_Stream_Textbox>"))
                            outputStringBuilder.AppendLine()
                            appendDateElement(outputStringBuilder, line, "my:arrival_date")
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:Facilities_Status")
                            outputStringBuilder.Append(importValues(line, "<my:Facilities_Status>", False, "2")) 'this should be 1 for coming versions of this method. 2 means complete
                            outputStringBuilder.AppendLine()
                            appendDateElement(outputStringBuilder, line, "my:facilities_Move_In_Date")
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:positiontype")
                            outputStringBuilder.Append(importValues(line, "<my:positiontype>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:costing_billing_code")
                            outputStringBuilder.Append(importValues(line, "<my:costing_billing_code>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:corporate_card_issue")
                            outputStringBuilder.Append(importValues(line, "<my:corporate_card_issue>", False, "N/A"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:corporate_card_revoke")
                            outputStringBuilder.Append(importValues(line, "<my:corporate_card_revoke>", False, "N/A"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:purchasing_card_issue")
                            outputStringBuilder.Append(importValues(line, "<my:purchasing_card_issue>", False, "N/A"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:purchasing_card_revoke")
                            outputStringBuilder.Append(importValues(line, "<my:purchasing_card_revoke>", False, "N/A"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:crew_issue")
                            outputStringBuilder.Append(importValues(line, "<my:crew_issue>", False, "N/A"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:gmas_issue")
                            outputStringBuilder.Append(importValues(line, "<my:gmas_issue>", False, "N/A"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:oracle_issue")
                            outputStringBuilder.Append(importValues(line, "<my:oracle_issue>", False, "N/A"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:hcom_issue")
                            outputStringBuilder.Append(importValues(line, "<my:hcom_issue>", False, "N/A"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:crew_revoke")
                            outputStringBuilder.Append(importValues(line, "<my:crew_revoke>", False, "N/A"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:gmas_revoke")
                            outputStringBuilder.Append(importValues(line, "<my:gmas_revoke>", False, "N/A"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:oracle_revoke")
                            outputStringBuilder.Append(importValues(line, "<my:oracle_revoke>", False, "N/A"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:hcom_revoke")
                            outputStringBuilder.Append(importValues(line, "<my:hcom_revoke>", False, "N/A"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:facilities_keys_returned")
                            outputStringBuilder.Append(importValues(line, "<my:facilities_keys_returned>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:facilities_phone_deactivated")
                            outputStringBuilder.Append(importValues(line, "<my:facilities_phone_deactivated>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:facilities_space_cleaned")
                            outputStringBuilder.Append(importValues(line, "<my:facilities_space_cleaned>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:goes_into_labs")
                            Dim goesIntoLabs As String = importValues(line, "<my:goes_into_labs>", True, "0")
                            outputStringBuilder.Append(goesIntoLabs)
                            If goesIntoLabs <> "" And goesIntoLabs <> " />" Then
                                outputStringBuilder.Append("</my:goes_into_labs>")
                            End If
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:autoID")
                            outputStringBuilder.Append(importValues(line, "<my:autoID>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:aries_posting_title")
                            outputStringBuilder.Append(importValues(line, "<my:aries_posting_title>", False, ""))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:aries_posting_req_number")
                            outputStringBuilder.Append(importValues(line, "<my:aries_posting_req_number>", False, ""))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:file_update_type")
                            outputStringBuilder.Append(importValues(line, "<my:file_update_type>", False, "Hire"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:it_computer_eligibility")
                            outputStringBuilder.Append(importValues(line, "<my:it_computer_eligibility>", False, "0"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:sends_email_to_finance")
                            outputStringBuilder.Append(importValues(line, "<my:sends_email_to_finance>", False, "0"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:sends_email_to_faculty_admins")
                            outputStringBuilder.Append(importValues(line, "<my:sends_email_to_faculty_admins>", False, "0"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:sends_email_to_communications")
                            outputStringBuilder.Append(importValues(line, "<my:sends_email_to_communications>", False, "0"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:sends_email_to_budget")
                            outputStringBuilder.Append(importValues(line, "<my:sends_email_to_budget>", False, "0"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:sends_email_to_events")
                            outputStringBuilder.Append(importValues(line, "<my:sends_email_to_events>", False, "0"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:sends_email_to_coi")
                            outputStringBuilder.Append(importValues(line, "<my:sends_email_to_coi>", False, "0"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:funding_type")
                            Dim funding_type As String = importValues(line, "<my:funding_type>", True, "")
                            outputStringBuilder.Append(funding_type)
                            If funding_type <> "" And funding_type <> " />" Then
                                outputStringBuilder.Append("</my:funding_type>")
                            End If
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:funds")
                            outputStringBuilder.Append(importValues(line, "<my:funds>", False, ""))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:facilities_proposed_office_building")
                            outputStringBuilder.Append(importValues(line, "<my:facilities_proposed_office_building>", False, ""))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:facilities_proposed_office_room")
                            outputStringBuilder.Append(importValues(line, "<my:facilities_proposed_office_room>", False, ""))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:office_being_manually_emailed")
                            outputStringBuilder.Append(importValues(line, "<my:office_being_manually_emailed>", False, ""))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:sends_email_to_area_admin_ap_ese")
                            Dim emails_jill As String = importValues(line, "<my:sends_email_to_area_admin_ap_ese>", True, "0")

                            Select Case position

                                Case "Assistant Professor"
                                    emails_jill = "1"
                                Case "Associate Professor"
                                    emails_jill = "1"
                                Case "Preceptor"
                                    emails_jill = "1"
                                Case "Professor of the Practice"
                                    emails_jill = "1"
                                Case "Lecturer"
                                    emails_jill = "1"
                                Case "Senior Preceptor"
                                    emails_jill = "1"
                                Case "Senior Lecturer"
                                    emails_jill = "1"
                                Case "Staff"
                                    emails_jill = "1"
                                Case "Tenured Professor"
                                    emails_jill = "1"
                                Case "Visiting Assistant Professor"
                                    emails_jill = "1"
                                Case "Visiting Associate Professor"
                                    emails_jill = "1"
                                Case "Visiting Professor"
                                    emails_jill = "1"

                            End Select

                            If area = "Applied Physics" Or area = "Environmental Science and Engineering" Then

                                Select Case position

                                    Case "Associate"
                                        emails_jill = "1"
                                    Case "Fellow"
                                        emails_jill = "1"
                                    Case "Postdoctoral Fellow - Employee"
                                        emails_jill = "1"
                                    Case "Postdoctoral Fellow - Stipend"
                                        emails_jill = "1"
                                    Case "Postdoctoral Fellow - Unpaid"
                                        emails_jill = "1"
                                    Case "Research Associate"
                                        emails_jill = "1"
                                    Case "Senior Research Fellow"
                                        emails_jill = "1"
                                    Case "Visiting Scholar"
                                        emails_jill = "1"
                                    Case "Visiting Undergraduate Research Intern"
                                        emails_jill = "1"

                                End Select

                            End If

                            outputStringBuilder.Append(emails_jill)
                            If emails_jill <> "" And emails_jill <> " />" Then
                                outputStringBuilder.Append("</my:sends_email_to_area_admin_ap_ese>")
                            End If
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:sends_email_to_area_admin_am_bio_msme")
                            Dim emails_arlene As String = importValues(line, "<my:sends_email_to_area_admin_am_bio_msme>", True, "0")

                            Select Case position

                                Case "Assistant Professor"
                                    emails_arlene = "1"
                                Case "Associate Professor"
                                    emails_arlene = "1"
                                Case "Preceptor"
                                    emails_arlene = "1"
                                Case "Professor of the Practice"
                                    emails_arlene = "1"
                                Case "Lecturer"
                                    emails_arlene = "1"
                                Case "Senior Preceptor"
                                    emails_arlene = "1"
                                Case "Senior Lecturer"
                                    emails_arlene = "1"
                                Case "Staff"
                                    emails_arlene = "1"
                                Case "Tenured Professor"
                                    emails_arlene = "1"
                                Case "Visiting Assistant Professor"
                                    emails_arlene = "1"
                                Case "Visiting Associate Professor"
                                    emails_arlene = "1"
                                Case "Visiting Professor"
                                    emails_arlene = "1"
                                Case "Visiting Lecturer"
                                    emails_arlene = "1"

                            End Select

                            If area = "Applied Math" Or area = "Bioengineering" Or area = "Materials Science and Mechanical Engineering" Then

                                Select Case position

                                    Case "Associate"
                                        emails_jill = "1"
                                    Case "Fellow"
                                        emails_jill = "1"
                                    Case "Postdoctoral Fellow - Employee"
                                        emails_jill = "1"
                                    Case "Postdoctoral Fellow - Stipend"
                                        emails_jill = "1"
                                    Case "Postdoctoral Fellow - Unpaid"
                                        emails_jill = "1"
                                    Case "Research Associate"
                                        emails_jill = "1"
                                    Case "Senior Research Fellow"
                                        emails_jill = "1"
                                    Case "Visiting Scholar"
                                        emails_jill = "1"
                                    Case "Visiting Undergraduate Research Intern"
                                        emails_jill = "1"

                                End Select

                            End If

                            outputStringBuilder.Append(emails_arlene)
                            If emails_arlene <> "" And emails_arlene <> " />" Then
                                outputStringBuilder.Append("</my:sends_email_to_area_admin_am_bio_msme>")
                            End If
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:sends_email_to_area_admin_cs_ee")
                            Dim emails_tristen As String = importValues(line, "<my:sends_email_to_area_admin_cs_ee>", True, "0")

                            If position = "Staff" Then

                                If hr_unit = "Area Administration" Or hr_unit = "Faculty and Lab Administration" Or hr_unit = "Research Staff" Then

                                    emails_tristen = "1"

                                End If

                            End If

                            If area = "Computer Science" Or area = "Electrical Engineering" Then

                                emails_tristen = "1"

                            End If

                            outputStringBuilder.Append(emails_tristen)
                            If emails_tristen <> "" And emails_tristen <> " />" Then
                                outputStringBuilder.Append("</my:sends_email_to_area_admin_cs_ee>")
                            End If
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:lab_building")
                            outputStringBuilder.Append(importValues(line, "<my:lab_building>", False, ""))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:lab_room_number")
                            outputStringBuilder.Append(importValues(line, "<my:lab_room_number>", False, ""))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:sends_email_to_research_admin")
                            Dim emails_research_admin As String = importValues(line, "<my:sends_email_to_research_admin>", True, "0")

                            If funding_type = "Sponsored" Or funding_type = "Combination" Then

                                If position <> "" Or position <> " />" Then

                                    emails_research_admin = "1"

                                End If

                            End If

                            outputStringBuilder.Append(emails_research_admin)
                            If emails_research_admin <> "" And emails_research_admin <> " />" Then
                                outputStringBuilder.Append("</my:sends_email_to_research_admin>")
                            End If
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:sends_email_to_safety")
                            Dim emails_safety As String = importValues(line, "<my:sends_email_to_safety>", True, "0")

                            If goesIntoLabs = "1" And (email_address <> "" Or email_address <> " />") Then

                                emails_safety = "1"

                            End If

                            outputStringBuilder.Append(emails_safety)
                            If emails_safety <> "" And emails_safety <> " />" Then
                                outputStringBuilder.Append("</my:sends_email_to_safety>")
                            End If
                            outputStringBuilder.AppendLine()

                            outputStringBuilder.Append("</my:myFields>")
                            outputStringBuilder.AppendLine()

                            writer.Write(outputStringBuilder.ToString)

                            outputStringBuilder = Nothing
                            filename = ""
                            line = ""

                            writer.Dispose()

                        Catch ex As Exception

                        End Try

                    Next fi

                End If ' Ends the OK of DirectoryBrowser

            End If ' Ends the OK of DirectoryBrowser

        Catch ex2 As Exception

        End Try

        btnExit.Enabled = True

    End Sub


    Private Function appendDateElement(ByRef outputStringBuilder As StringBuilder, ByRef line As String, ByVal field As String)

        outputStringBuilder.Append("<" & field)

        If line.Contains("<" & field & "/>") = True Then ' v1 empty

            'outputStringBuilder.Append(" xsi:nil=""true""></" & field & ">")\
            outputStringBuilder.Append(" xsi:nil=""true"" />")

        ElseIf line.Contains("<" & field & "></") = True Then 'v2 empty

            'outputStringBuilder.Append(" xsi:nil=""true""></" & field & ">")
            outputStringBuilder.Append(" xsi:nil=""true"" />")

        ElseIf line.Contains("<" & field & " xsi:nil") = True Then 'v2 v2 empty

            'outputStringBuilder.Append(" xsi:nil=""true""></" & field & ">")
            outputStringBuilder.Append(" xsi:nil=""true"" />")

        ElseIf line.Contains("<" & field & ">") = True Then 'Contains something

            Dim charsToCut As Integer = line.IndexOf("</" & field & ">") - line.IndexOf("<" & field & ">") - (field.Length + 2)

            outputStringBuilder.Append(">" & line.Substring(line.IndexOf("<" & field & ">") + field.Length + 2, charsToCut) & "</" & field & ">")

        ElseIf line.Contains("<" & field & ">") = False Then ' Element does not exist in file (v1 files normally)

            'outputStringBuilder.Append(" xsi:nil=""true""></" & field & ">")
            outputStringBuilder.Append(" xsi:nil=""true"" />")

        End If

    End Function


    Private Function importValues(ByRef line As String, ByVal field As String, Optional ByVal skipClosingTag As Boolean = False, Optional ByVal defaultValue As String = "") As String

        field = field.Replace("<", "").Replace(">", "").Replace("</", "")

        If line.Contains("<" & field & "/>") = True Then ' v1 empty

            If defaultValue <> "" Then

                If skipClosingTag = False Then
                    Return ">" & defaultValue & "</" & field & ">"
                Else
                    Return ">" & defaultValue
                End If

            Else

                If skipClosingTag = False Then
                    Return defaultValue & " />"
                Else
                    Return defaultValue & " />"
                End If

            End If

        ElseIf line.Contains("<" & field & "></") = True Then 'v2 empty

            If defaultValue <> "" Then

                If skipClosingTag = False Then
                    Return ">" & defaultValue & "</" & field & ">"
                Else
                    Return ">" & defaultValue
                End If

            Else

                If skipClosingTag = False Then
                    Return defaultValue & " />"
                Else
                    Return defaultValue & " />"
                End If

            End If

        ElseIf line.Contains("<" & field & ">") = True And line.Contains("<" & field & "></") = False Then 'Contains something

            Dim charsToCut As Integer = line.IndexOf("</" & field & ">") - line.IndexOf("<" & field & ">") - (field.Length + 2)

            If skipClosingTag = False Then
                Return ">" & line.Substring(line.IndexOf("<" & field & ">") + field.Length + 2, charsToCut) & "</" & field & ">"
            Else
                Return ">" & line.Substring(line.IndexOf("<" & field & ">") + field.Length + 2, charsToCut)
            End If

        ElseIf line.Contains("<" & field & ">") = False Then ' Element does not exist in file (v1 files normally)

            If defaultValue <> "" Then

                If skipClosingTag = False Then
                    Return ">" & defaultValue & "</" & field & ">"
                Else
                    Return ">" & defaultValue
                End If

            Else

                If skipClosingTag = False Then
                    Return defaultValue & " />"
                Else
                    Return defaultValue & " />"
                End If

            End If


        End If

    End Function


    Private Function importAndDivideCommentValues(ByRef outputStringBuilder As StringBuilder, ByRef line As String)

        Dim itemPastActions As String = ""
        Dim lineTemp As String = importValues(line, "<my:comments>", True)

        If lineTemp.Contains("|") = True Then 'Old versions have the | separator

            Dim tmpComments As String() = lineTemp.Split("|")

            For i = 0 To tmpComments.Length - 1

                If tmpComments(i).Contains("emailed") Then

                    itemPastActions += tmpComments(i) & ";"

                Else

                    outputStringBuilder.Append(tmpComments(i))

                End If

            Next i

        ElseIf lineTemp.Contains("201") = True And lineTemp.Contains(";") = True Then 'Action check with part of the year (2012), ";" separator case

            Dim tmpComments As String() = lineTemp.Split(";")

            For i = 0 To tmpComments.Length - 1

                If tmpComments(i).Contains("emailed") Then

                    itemPastActions += tmpComments(i) & ";"

                Else

                    outputStringBuilder.Append(tmpComments(i))

                End If

            Next i

        End If

        outputStringBuilder.Append("</my:comments>")

        lineTemp = ""

        Return itemPastActions

    End Function


    Private Sub updateForms()

        btnExit.Enabled = False

        Try

            inputFolder.ShowNewFolderButton = False
            inputFolder.Description = "Please choose the folder where the outdated forms are"

            If inputFolder.ShowDialog() = DialogResult.OK Then

                outputFolder.ShowNewFolderButton = False
                outputFolder.Description = "Please choose the folder where the UPDATED forms will be saved"

                If outputFolder.ShowDialog() = DialogResult.OK Then

                    If idFinder.ShowDialog() = DialogResult.OK Then

                        Dim idsReader As FileIO.TextFieldParser = New FileIO.TextFieldParser(idFinder.FileName)
                        Dim CurrentRow As String()
                        idsReader.TextFieldType = FileIO.FieldType.Delimited
                        idsReader.Delimiters = New String() {","}
                        idsReader.HasFieldsEnclosedInQuotes = False

                        Do While Not idsReader.EndOfData

                            Try

                                CurrentRow = idsReader.ReadFields

                                Dim filename As String = CurrentRow(0)
                                Dim itemAutoID As String = CurrentRow(1)

                                Dim hasAutoID As Boolean = False

                                rtb.AppendText("Reading " & inputFolder.SelectedPath & "\" & filename & Chr(13))

                                Dim reader As StreamReader = New StreamReader(inputFolder.SelectedPath & "\" & filename)
                                Dim writer As New StreamWriter(outputFolder.SelectedPath & "\" & itemAutoID & ".xml")

                                Dim line As String

                                Do

                                    line = reader.ReadLine()

                                    If line Is Nothing Then
                                        Exit Do
                                    End If

                                    If line.Contains("solutionVersion=") = True Then

                                        line = line.Substring(0, line.IndexOf("solutionVersion=")) & "solutionVersion=" & """1.0.0.2154""" & line.Substring(line.IndexOf("solutionVersion=") + 28)

                                    End If

                                    If line.Contains("<my:autoID></my:autoID>") = True Then

                                        line = line.Replace("<my:autoID></my:autoID>", "<my:autoID>" & itemAutoID & "</my:autoID>")
                                        hasAutoID = True

                                    End If

                                    If line.Contains("<my:autoID>") = True And line.Contains("<my:autoID></my:autoID>") = False Then

                                        line = line.Substring(0, line.IndexOf("<my:autoID>")) & "<my:autoID>" & itemAutoID & "</my:autoID>" & line.Substring(line.IndexOf("</my:autoID>") + 12)
                                        hasAutoID = True

                                    End If

                                    If line.Contains("</my:myFields>") = True And hasAutoID = False Then

                                        line = line.Replace("</my:myFields>", "<my:autoID>" & itemAutoID & "</my:autoID></my:myFields>")
                                        hasAutoID = True

                                    End If

                                    writer.WriteLine(line)


                                Loop ' next attribute until EOF

                                filename = ""
                                itemAutoID = ""
                                hasAutoID = False
                                line = ""

                                reader.Dispose()
                                writer.Dispose()

                            Catch ex As Exception

                            End Try

                        Loop

                    End If ' Ends the OK of DirectoryBrowser

                End If ' Ends the OK of FileBrowser

            End If ' Ends the OK of DirectoryBrowser

        Catch ex2 As Exception

        End Try

        btnExit.Enabled = True

    End Sub


    Private Sub updateFormsv2()

        btnExit.Enabled = False

        Try

            inputFolder.ShowNewFolderButton = False
            inputFolder.Description = "Please choose the folder where the outdated forms are"

            If inputFolder.ShowDialog() = DialogResult.OK Then

                outputFolder.ShowNewFolderButton = False
                outputFolder.Description = "Please choose the folder where the UPDATED forms will be saved"

                If outputFolder.ShowDialog() = DialogResult.OK Then

                    'If idFinder.ShowDialog() = DialogResult.OK Then

                    'Dim idsReader As FileIO.TextFieldParser = New FileIO.TextFieldParser(idFinder.FileName)
                    'Dim CurrentRow As String()
                    'idsReader.TextFieldType = FileIO.FieldType.Delimited
                    'idsReader.Delimiters = New String() {","}
                    'idsReader.HasFieldsEnclosedInQuotes = False

                    'Do While Not idsReader.EndOfData

                    'Try

                    'CurrentRow = idsReader.ReadFields

                    'Dim filename As String = CurrentRow(0)
                    'Dim itemAutoID As String = CurrentRow(1)

                    Dim dirInfo As New DirectoryInfo(inputFolder.SelectedPath)

                    For Each fi As FileInfo In dirInfo.GetFiles()

                        Dim filename As String = fi.Name

                        Try

                            rtb.AppendText("Reading " & inputFolder.SelectedPath & "\" & filename & Chr(13))

                            Dim reader As StreamReader = New StreamReader(inputFolder.SelectedPath & "\" & filename)
                            'Dim writer As New StreamWriter(outputFolder.SelectedPath & "\" & filename & ".xml")
                            Dim writer As New StreamWriter(outputFolder.SelectedPath & "\" & filename)

                            Dim line As String = reader.ReadToEnd()
                            reader.Dispose()

                            Dim outputStringBuilder As New StringBuilder

                            outputStringBuilder.Append("<?xml version=""1.0"" encoding=""utf-8""?>")
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<?mso-infoPathSolution name=""urn:schemas-microsoft-com:office:infopath:AAAD-Forms-v3:-myXSD-2009-11-02T17-20-04"" solutionVersion=""1.0.0.2384"" productVersion=""12.0.0.0"" PIVersion=""1.0.0.0"" href=""https://sp.seas.harvard.edu/sites/eforms/AAAD%20Forms%20v3/Forms/template.xsn""?>")
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<?mso-application progid=""InfoPath.Document"" versionProgid=""InfoPath.Document.2""?>")
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("<my:myFields xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:dfs=""http://schemas.microsoft.com/office/infopath/2003/dataFormSolution"" xmlns:xhtml=""http://www.w3.org/1999/xhtml"" xmlns:my=""http://schemas.microsoft.com/office/infopath/2003/myXSD/2009-11-02T17:20:04"" xmlns:xd=""http://schemas.microsoft.com/office/infopath/2003"" xml:lang=""en-US"">")
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:hr_aspire_posting_title>")
                            outputStringBuilder.Append(importValues(line, "<my:hr_aspire_posting_title>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:finance_costing_approved>")
                            outputStringBuilder.Append(importValues(line, "<my:finance_costing_approved>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:facilities_furniture_redistributed>")
                            outputStringBuilder.Append(importValues(line, "<my:facilities_furniture_redistributed>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:pi>")
                            Dim pi As String = importValues(line, "<my:pi>", True)
                            outputStringBuilder.Append(pi)
                            outputStringBuilder.Append("</my:pi>")
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:it_account_created>")
                            outputStringBuilder.Append(importValues(line, "<my:it_account_created>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:it_account_disabled>")
                            outputStringBuilder.Append(importValues(line, "<my:it_account_disabled>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:it_equipment_reclaimed>")
                            outputStringBuilder.Append(importValues(line, "<my:it_equipment_reclaimed>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:sponsresrch_costing_approval_for_academic_or_staff_appointment>")
                            outputStringBuilder.Append(importValues(line, "<my:sponsresrch_costing_approval_for_academic_or_staff_appointment>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:sponsresrch_costing_approved>")
                            outputStringBuilder.Append(importValues(line, "<my:sponsresrch_costing_approved>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:sponsresrch_gmas_access_granted>")
                            outputStringBuilder.Append(importValues(line, "<my:sponsresrch_gmas_access_granted>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:sponsresrch_pi_award_transferred>")
                            outputStringBuilder.Append(importValues(line, "<my:sponsresrch_pi_award_transferred>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:sponsresrch_pi_award_terminated>")
                            outputStringBuilder.Append(importValues(line, "<my:sponsresrch_pi_award_terminated>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:studaffrs_notified_by_area_admins_of_faculty_arrival_date>")
                            outputStringBuilder.Append(importValues(line, "<my:studaffrs_notified_by_area_admins_of_faculty_arrival_date>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:studaffrs_grad_students_assigned_to_new_faculty>")
                            outputStringBuilder.Append(importValues(line, "<my:studaffrs_grad_students_assigned_to_new_faculty>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:studaffrs_keys_returned_to_facilities>")
                            outputStringBuilder.Append(importValues(line, "<my:studaffrs_keys_returned_to_facilities>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:studaffrs_id_deactivated_and_returned>")
                            outputStringBuilder.Append(importValues(line, "<my:studaffrs_id_deactivated_and_returned>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:studaffrs_payroll_notified_of_termination>")
                            outputStringBuilder.Append(importValues(line, "<my:studaffrs_payroll_notified_of_termination>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:studaffrs_area_admins_notified_of_departure>")
                            outputStringBuilder.Append(importValues(line, "<my:studaffrs_area_admins_notified_of_departure>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:procurement_iprocurement_authorization_form_completed_by_pi>")
                            outputStringBuilder.Append(importValues(line, "<my:procurement_iprocurement_authorization_form_completed_by_pi>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:procurement_iprocurement_training_scheduled>")
                            outputStringBuilder.Append(importValues(line, "<my:procurement_iprocurement_training_scheduled>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:procurement_pcard_training_scheduled>")
                            outputStringBuilder.Append(importValues(line, "<my:procurement_pcard_training_scheduled>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:procurement_corporate_card_training_scheduled>")
                            outputStringBuilder.Append(importValues(line, "<my:procurement_corporate_card_training_scheduled>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:procurement_notified_by_sponsored_research_of_pi_departure>")
                            outputStringBuilder.Append(importValues(line, "<my:procurement_notified_by_sponsored_research_of_pi_departure>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:facilities_comments>")
                            outputStringBuilder.Append(importValues(line, "<my:facilities_comments>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:areaadmins_studaffrs_notified_of_faculty_arrival>")
                            outputStringBuilder.Append(importValues(line, "<my:areaadmins_studaffrs_notified_of_faculty_arrival>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:areaadmins_appointing_office_notified_of_departure>")
                            outputStringBuilder.Append(importValues(line, "<my:areaadmins_appointing_office_notified_of_departure>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:areaadmins_faculty_support_reassigned>")
                            outputStringBuilder.Append(importValues(line, "<my:areaadmins_faculty_support_reassigned>", False, "false"))
                            outputStringBuilder.AppendLine()
                            'outputStringBuilder.Append("	<my:info_url>https://sp.seas.harvard.edu/sites/eforms/_layouts/FormServer.aspx?XmlLocation=/sites/eforms/aaad%20forms%20v3/" & itemAutoID & ".xml&amp;OpenIn=Browser</my:info_url>")
                            outputStringBuilder.Append("	<my:info_url>")
                            outputStringBuilder.Append(importValues(line, "<my:info_url>", False, "").Replace("aaad%20forms%20v2", "aaad%20forms%20v3"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:position>")
                            Dim position As String = importValues(line, "<my:position>", True)
                            outputStringBuilder.Append(position)
                            outputStringBuilder.Append("</my:position>")
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:alternate_email_address>")
                            outputStringBuilder.Append(importValues(line, "<my:alternate_email_address>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:supervisor>")
                            Dim supervisor As String = importValues(line, "<my:supervisor>", True)
                            outputStringBuilder.Append(supervisor)
                            outputStringBuilder.Append("</my:supervisor>")
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:huid>")
                            outputStringBuilder.Append(importValues(line, "<my:huid>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:email_address>")
                            outputStringBuilder.Append(importValues(line, "<my:email_address>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:hr_req_number>")
                            outputStringBuilder.Append(importValues(line, "<my:hr_req_number>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:office_phone_number>")
                            outputStringBuilder.Append(importValues(line, "<my:office_phone_number>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:it_pc_delivered>")
                            outputStringBuilder.Append(importValues(line, "<my:it_pc_delivered>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:it_pc_set_up>")
                            outputStringBuilder.Append(importValues(line, "<my:it_pc_set_up>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:hr_unit>")
                            outputStringBuilder.Append(importValues(line, "<my:hr_unit>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:office_building>")
                            outputStringBuilder.Append(importValues(line, "<my:office_building>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:office_room_number>")
                            outputStringBuilder.Append(importValues(line, "<my:office_room_number>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:first_name>")
                            outputStringBuilder.Append(importValues(line, "<my:first_name>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:middle_name>")
                            outputStringBuilder.Append(importValues(line, "<my:middle_name>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:last_name>")
                            outputStringBuilder.Append(importValues(line, "<my:last_name>"))
                            outputStringBuilder.AppendLine()
                            appendDateElement(outputStringBuilder, line, "my:start_date")
                            outputStringBuilder.AppendLine()
                            appendDateElement(outputStringBuilder, line, "my:termination_date")
                            outputStringBuilder.AppendLine()

                            outputStringBuilder.Append("	<my:comments>")
                            Dim itemPastActions As String = ""
                            itemPastActions = importAndDivideCommentValues(outputStringBuilder, line)
                            outputStringBuilder.AppendLine()

                            outputStringBuilder.Append("	<my:acdaffrs_area>")
                            outputStringBuilder.Append(importValues(line, "<my:acdaffrs_area>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:seas_username>")
                            outputStringBuilder.Append(importValues(line, "<my:seas_username>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:office_phone_jack_number>")
                            outputStringBuilder.Append(importValues(line, "<my:office_phone_jack_number>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:office_phone_jack_activated>")
                            outputStringBuilder.Append(importValues(line, "<my:office_phone_jack_activated>", False, "false"))
                            outputStringBuilder.AppendLine()
                            appendDateElement(outputStringBuilder, line, "my:expected_pc_arrival_date")
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:pc_ordered>")
                            outputStringBuilder.Append(importValues(line, "<my:pc_ordered>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:copier_code>")
                            outputStringBuilder.Append(importValues(line, "<my:copier_code>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:assigned_mailbox>")
                            outputStringBuilder.Append(importValues(line, "<my:assigned_mailbox>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:office_key_provided>")
                            outputStringBuilder.Append(importValues(line, "<my:office_key_provided>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:swipe_access_activated>")
                            outputStringBuilder.Append(importValues(line, "<my:swipe_access_activated>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:assigned_faculty_support>")
                            outputStringBuilder.Append(importValues(line, "<my:assigned_faculty_support>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:safety_general_training_complete>")
                            outputStringBuilder.Append(importValues(line, "<my:safety_general_training_complete>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:safety_general_training_na>")
                            outputStringBuilder.Append(importValues(line, "<my:safety_general_training_na>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:safety_hazardous_waste_training_complete>")
                            outputStringBuilder.Append(importValues(line, "<my:safety_hazardous_waste_training_complete>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:safety_hazardous_waste_training_na>")
                            outputStringBuilder.Append(importValues(line, "<my:safety_hazardous_waste_training_na>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:safety_bio_training_complete>")
                            outputStringBuilder.Append(importValues(line, "<my:safety_bio_training_complete>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:safety_bio_training_na>")
                            outputStringBuilder.Append(importValues(line, "<my:safety_bio_training_na>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:safety_radiation_training_complete>")
                            outputStringBuilder.Append(importValues(line, "<my:safety_radiation_training_complete>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:safety_radiation_training_na>")
                            outputStringBuilder.Append(importValues(line, "<my:safety_radiation_training_na>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:safety_laser_training_complete>")
                            outputStringBuilder.Append(importValues(line, "<my:safety_laser_training_complete>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:safety_laser_training_na>")
                            outputStringBuilder.Append(importValues(line, "<my:safety_laser_training_na>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:safety_removed_from_training_lists>")
                            outputStringBuilder.Append(importValues(line, "<my:safety_removed_from_training_lists>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:nick_name>")
                            outputStringBuilder.Append(importValues(line, "<my:nick_name>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:hire_type>")
                            outputStringBuilder.Append(importValues(line, "<my:hire_type>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:incumbent_name>")
                            outputStringBuilder.Append(importValues(line, "<my:incumbent_name>"))
                            outputStringBuilder.AppendLine()

                            outputStringBuilder.Append("	<my:Activity_Stream_Textbox>")
                            outputStringBuilder.Append(importValues(line, "<my:Activity_Stream_Textbox>", True))
                            outputStringBuilder.Append(itemPastActions & "</my:Activity_Stream_Textbox>")
                            itemPastActions = ""
                            outputStringBuilder.AppendLine()

                            appendDateElement(outputStringBuilder, line, "my:arrival_date")
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:Facilities_Status>")
                            outputStringBuilder.Append(importValues(line, "<my:Facilities_Status>"))
                            outputStringBuilder.AppendLine()
                            appendDateElement(outputStringBuilder, line, "my:facilities_Move_In_Date")
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:positiontype>")
                            outputStringBuilder.Append(importValues(line, "<my:positiontype>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:costing_billing_code>")
                            outputStringBuilder.Append(importValues(line, "<my:costing_billing_code>"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:corporate_card_issue>")
                            outputStringBuilder.Append(importValues(line, "<my:corporate_card_issue>", False, "N/A"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:corporate_card_revoke>")
                            outputStringBuilder.Append(importValues(line, "<my:corporate_card_revoke>", False, "N/A"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:purchasing_card_issue>")
                            outputStringBuilder.Append(importValues(line, "<my:purchasing_card_issue>", False, "N/A"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:purchasing_card_revoke>")
                            outputStringBuilder.Append(importValues(line, "<my:purchasing_card_revoke>", False, "N/A"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:crew_issue>")
                            outputStringBuilder.Append(importValues(line, "<my:crew_issue>", False, "N/A"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:gmas_issue>")
                            outputStringBuilder.Append(importValues(line, "<my:gmas_issue>", False, "N/A"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:oracle_issue>")
                            outputStringBuilder.Append(importValues(line, "<my:oracle_issue>", False, "N/A"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:hcom_issue>")
                            outputStringBuilder.Append(importValues(line, "<my:hcom_issue>", False, "N/A"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:crew_revoke>")
                            outputStringBuilder.Append(importValues(line, "<my:crew_revoke>", False, "N/A"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:gmas_revoke>")
                            outputStringBuilder.Append(importValues(line, "<my:gmas_revoke>", False, "N/A"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:oracle_revoke>")
                            outputStringBuilder.Append(importValues(line, "<my:oracle_revoke>", False, "N/A"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:hcom_revoke>")
                            outputStringBuilder.Append(importValues(line, "<my:hcom_revoke>", False, "N/A"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:facilities_keys_returned>")
                            outputStringBuilder.Append(importValues(line, "<my:facilities_keys_returned>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:facilities_phone_deactivated>")
                            outputStringBuilder.Append(importValues(line, "<my:facilities_phone_deactivated>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:facilities_space_cleaned>")
                            outputStringBuilder.Append(importValues(line, "<my:facilities_space_cleaned>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:goes_into_labs>") 'used to be called in_lab, now called goes_into_labs
                            Dim goesIntoLabs As String = importValues(line, "<my:in_lab>", True, "0")

                            Select Case pi

                                Case "jaiz"
                                    goesIntoLabs = "1"
                                Case "aziz"
                                    goesIntoLabs = "1"
                                Case "bertoldi"
                                    goesIntoLabs = "1"
                                Case "capasso"
                                    goesIntoLabs = "1"
                                Case "clarke"
                                    goesIntoLabs = "1"
                                Case "kcrozier"
                                    goesIntoLabs = "1"
                                Case "dedwards"
                                    goesIntoLabs = "1"
                                Case "golovchenko"
                                    goesIntoLabs = "1"
                                Case "donhee"
                                    goesIntoLabs = "1"
                                Case "hau"
                                    goesIntoLabs = "1"
                                Case "howe"
                                    goesIntoLabs = "1"
                                Case "ehu"
                                    goesIntoLabs = "1"
                                Case "jalewis"
                                    goesIntoLabs = "1"
                                Case "loncar"
                                    goesIntoLabs = "1"
                                Case "lm"
                                    goesIntoLabs = "1"
                                Case "vnm"
                                    goesIntoLabs = "1"
                                Case "smartin"
                                    goesIntoLabs = "1"
                                Case "mazur"
                                    goesIntoLabs = "1"
                                Case "mitchell"
                                    goesIntoLabs = "1"
                                Case "mooneyd"
                                    goesIntoLabs = "1"
                                Case "kkparker"
                                    goesIntoLabs = "1"
                                Case "shriram"
                                    goesIntoLabs = "1"
                                Case "spaepen"
                                    goesIntoLabs = "1"
                                Case "vecitis"
                                    goesIntoLabs = "1"
                                Case "jvlassak"
                                    goesIntoLabs = "1"
                                Case "walsh"
                                    goesIntoLabs = "1"
                                Case "weitz"
                                    goesIntoLabs = "1"
                                Case "bmw"
                                    goesIntoLabs = "1"
                                Case "rjwood"
                                    goesIntoLabs = "1"

                            End Select

                            Select Case supervisor

                                Case "iavdagic"
                                    goesIntoLabs = "1"
                                Case "achalah"
                                    goesIntoLabs = "1"
                                Case "dclaflin"
                                    goesIntoLabs = "1"
                                Case "ldefeo"
                                    goesIntoLabs = "1"
                                Case "graham"
                                    goesIntoLabs = "1"
                                Case "habbalf"
                                    goesIntoLabs = "1"
                                Case "hollar"
                                    goesIntoLabs = "1"
                                Case "aleite"
                                    goesIntoLabs = "1"
                                Case "jmacarth"
                                    goesIntoLabs = "1"
                                Case "cnielsen"
                                    goesIntoLabs = "1"

                            End Select

                            outputStringBuilder.Append(goesIntoLabs)
                            outputStringBuilder.Append("</my:goes_into_labs>")
                            outputStringBuilder.AppendLine()
                            'outputStringBuilder.Append("	<my:autoID>" & itemAutoID & "</my:autoID>")
                            outputStringBuilder.Append("	<my:autoID>")
                            outputStringBuilder.Append(importValues(line, "<my:autoID>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:aries_posting_title>")
                            outputStringBuilder.Append(importValues(line, "<my:aries_posting_title>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:aries_posting_req_number>")
                            outputStringBuilder.Append(importValues(line, "<my:aries_posting_req_number>", False, "false"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:file_update_type>")
                            outputStringBuilder.Append(importValues(line, "<my:file_update_type>", False, "Hire"))
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:it_computer_eligibility>")
                            Dim emails_computing As String = importValues(line, "<my:it_computer_eligibility>", True, "0")

                            Select Case position

                                Case "Assistant Professor"
                                    emails_computing = "1"
                                Case "Associate Professor"
                                    emails_computing = "1"
                                Case "Postdoctoral Fellow - Employee (computer eligible)"
                                    emails_computing = "1"
                                Case "Preceptor"
                                    emails_computing = "1"
                                Case "Professor of the Practice"
                                    emails_computing = "1"
                                Case "Research Associate"
                                    emails_computing = "1"
                                Case "Senior Preceptor"
                                    emails_computing = "1"
                                Case "Senior Research Fellow"
                                    emails_computing = "1"
                                Case "Staff"
                                    emails_computing = "1"
                                Case "Tenured Professor"
                                    emails_computing = "1"

                            End Select

                            outputStringBuilder.Append(emails_computing)
                            outputStringBuilder.Append("</my:it_computer_eligibility>")
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:sends_email_to_finance>")
                            Dim emails_finance As String = importValues(line, "<my:sends_email_to_finance>", True, "0")

                            Select Case position

                                Case "Assistant Professor"
                                    emails_finance = "1"
                                Case "Associate Professor"
                                    emails_finance = "1"
                                Case "Lecturer"
                                    emails_finance = "1"
                                Case "Preceptor"
                                    emails_finance = "1"
                                Case "Professor of the Practice"
                                    emails_finance = "1"
                                Case "Senior Lecturer"
                                    emails_finance = "1"
                                Case "Senior Preceptor"
                                    emails_finance = "1"
                                Case "Staff"
                                    emails_finance = "1"
                                Case "Tenured Professor"
                                    emails_finance = "1"
                                Case "Visiting Assistant Professor"
                                    emails_finance = "1"
                                Case "Visiting Associate Professor"
                                    emails_finance = "1"
                                Case "Visiting Lecturer"
                                    emails_finance = "1"

                            End Select

                            outputStringBuilder.Append(emails_finance)
                            outputStringBuilder.Append("</my:sends_email_to_finance>")
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:sends_email_to_faculty_admins>")
                            Dim emails_facultyadmins As String = importValues(line, "<my:sends_email_to_faculty_admins>", True, "0")

                            Select Case position

                                Case "Associate"
                                    emails_facultyadmins = "1"
                                Case "Fellow"
                                    emails_facultyadmins = "1"
                                Case "Lecturer"
                                    emails_facultyadmins = "1"
                                Case "Postdoctoral Fellow - Employee (computer eligible)"
                                    emails_facultyadmins = "1"
                                Case "Postdoctoral Fellow - Stipend (not computer eligible)"
                                    emails_facultyadmins = "1"
                                Case "Postdoctoral Fellow - Unpaid (not computer eligible)"
                                    emails_facultyadmins = "1"
                                Case "Preceptor"
                                    emails_facultyadmins = "1"
                                Case "Research Associate"
                                    emails_facultyadmins = "1"
                                Case "Senior Lecturer"
                                    emails_facultyadmins = "1"
                                Case "Senior Preceptor"
                                    emails_facultyadmins = "1"
                                Case "Visiting Assistant Professor"
                                    emails_facultyadmins = "1"
                                Case "Visiting Associate Professor"
                                    emails_facultyadmins = "1"
                                Case "Visiting Lecturer"
                                    emails_facultyadmins = "1"
                                Case "Visiting Scholar"
                                    emails_facultyadmins = "1"
                                Case "Visiting Undergraduate Research Intern"
                                    emails_facultyadmins = "1"

                            End Select

                            outputStringBuilder.Append(emails_facultyadmins)
                            outputStringBuilder.Append("</my:sends_email_to_faculty_admins>")
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:sends_email_to_communications>")
                            Dim emails_communications As String = importValues(line, "<my:sends_email_to_communications>", True, "0")

                            Select Case position

                                Case "Assistant Professor"
                                    emails_communications = "1"
                                Case "Associate Professor"
                                    emails_communications = "1"
                                Case "Lecturer"
                                    emails_communications = "1"
                                Case "Preceptor"
                                    emails_communications = "1"
                                Case "Professor of the Practice"
                                    emails_communications = "1"
                                Case "Senior Lecturer"
                                    emails_communications = "1"
                                Case "Senior Preceptor"
                                    emails_communications = "1"
                                Case "Staff"
                                    emails_communications = "1"
                                Case "Tenured Professor"
                                    emails_communications = "1"
                                Case "Visiting Lecturer"
                                    emails_communications = "1"

                            End Select

                            outputStringBuilder.Append(emails_communications)
                            outputStringBuilder.Append("</my:sends_email_to_communications>")
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:sends_email_to_budget>")
                            Dim emails_budget As String = importValues(line, "<my:sends_email_to_budget>", True, "0")

                            Select Case position

                                Case "Assistant Professor"
                                    emails_budget = "1"
                                Case "Associate Professor"
                                    emails_budget = "1"
                                Case "Lecturer"
                                    emails_budget = "1"
                                Case "Preceptor"
                                    emails_budget = "1"
                                Case "Professor of the Practice"
                                    emails_budget = "1"
                                Case "Senior Lecturer"
                                    emails_budget = "1"
                                Case "Senior Preceptor"
                                    emails_budget = "1"
                                Case "Staff"
                                    emails_budget = "1"
                                Case "Tenured Professor"
                                    emails_budget = "1"
                                Case "Visiting Assistant Professor"
                                    emails_budget = "1"
                                Case "Visiting Associate Professor"
                                    emails_budget = "1"
                                Case "Visiting Lecturer"
                                    emails_budget = "1"

                            End Select

                            outputStringBuilder.Append(emails_budget)
                            outputStringBuilder.Append("</my:sends_email_to_budget>")
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:sends_email_to_events>")
                            Dim emails_events As String = importValues(line, "<my:sends_email_to_events>", True, "0")

                            Select Case position

                                Case "Assistant Professor"
                                    emails_events = "1"
                                Case "Associate Professor"
                                    emails_events = "1"
                                Case "Professor of the Practice"
                                    emails_events = "1"
                                Case "Senior Lecturer"
                                    emails_events = "1"
                                Case "Senior Research Fellow"
                                    emails_events = "1"
                                Case "Staff"
                                    emails_events = "1"
                                Case "Tenured Professor"
                                    emails_events = "1"

                            End Select

                            outputStringBuilder.Append(emails_events)
                            outputStringBuilder.Append("</my:sends_email_to_events>")
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("	<my:sends_email_to_coi>")
                            Dim emails_coi As String = importValues(line, "<my:sends_email_to_coi>", True, "0")

                            Select Case position

                                Case "Assistant Professor"
                                    emails_coi = "1"
                                Case "Associate Professor"
                                    emails_coi = "1"
                                Case "Preceptor"
                                    emails_coi = "1"
                                Case "Professor of the Practice"
                                    emails_coi = "1"
                                Case "Senior Preceptor"
                                    emails_coi = "1"
                                Case "Tenured Professor"
                                    emails_coi = "1"
                                Case "Visiting Assistant Professor"
                                    emails_coi = "1"
                                Case "Visiting Associate Professor"
                                    emails_coi = "1"
                                Case "Visiting Lecturer"
                                    emails_coi = "1"
                                Case "Visiting Professor"
                                    emails_coi = "1"

                            End Select

                            outputStringBuilder.Append(emails_coi)
                            outputStringBuilder.Append("</my:sends_email_to_coi>")
                            outputStringBuilder.AppendLine()
                            outputStringBuilder.Append("</my:myFields>")
                            outputStringBuilder.AppendLine()

                            writer.Write(outputStringBuilder.ToString)

                            outputStringBuilder = Nothing
                            filename = ""
                            'itemAutoID = ""
                            line = ""

                            writer.Dispose()

                        Catch ex As Exception

                        End Try

                        'Loop

                    Next fi

                    'End If ' Ends the OK of DirectoryBrowser

                End If ' Ends the OK of FileBrowser

            End If ' Ends the OK of DirectoryBrowser

        Catch ex2 As Exception

        End Try

        btnExit.Enabled = True

    End Sub


    Private Sub updateFormsv3()

        btnExit.Enabled = False

        Try

            inputFolder.ShowNewFolderButton = False
            inputFolder.Description = "Please choose the folder where the outdated forms are"

            If inputFolder.ShowDialog() = DialogResult.OK Then

                outputFolder.ShowNewFolderButton = False
                outputFolder.Description = "Please choose the folder where the UPDATED forms will be saved"

                If outputFolder.ShowDialog() = DialogResult.OK Then

                    If idFinder.ShowDialog() = DialogResult.OK Then

                        Dim idsReader As FileIO.TextFieldParser = New FileIO.TextFieldParser(idFinder.FileName)
                        Dim CurrentRow As String()
                        idsReader.TextFieldType = FileIO.FieldType.Delimited
                        idsReader.Delimiters = New String() {","}
                        idsReader.HasFieldsEnclosedInQuotes = False

                        Do While Not idsReader.EndOfData

                            Try

                                CurrentRow = idsReader.ReadFields

                                Dim filename As String = CurrentRow(0)
                                Dim itemAutoID As String = CurrentRow(1)

                                Dim hasAutoID As Boolean = False

                                rtb.AppendText("Reading " & inputFolder.SelectedPath & "\" & filename & Chr(13))

                                Dim reader As StreamReader = New StreamReader(inputFolder.SelectedPath & "\" & filename) 'The export from Sharepoint already has .xml at the end
                                Dim writer As New StreamWriter(outputFolder.SelectedPath & "\" & itemAutoID & ".xml")

                                Dim line As String = reader.ReadToEnd()
                                reader.Dispose()

                                If line.Contains("<my:autoID>") = True And line.Contains("<my:autoID></my:autoID>") = False Then

                                    line = line.Substring(0, line.IndexOf("<my:autoID>")) & "<my:autoID>" & itemAutoID & "</my:autoID>" & line.Substring(line.IndexOf("</my:autoID>") + 12)

                                End If
                                If line.Contains("<my:info_url>") = True Then

                                    line = line.Substring(0, line.IndexOf("<my:info_url>")) & "<my:info_url>" & "https://sp.seas.harvard.edu/sites/eforms/_layouts/FormServer.aspx?XmlLocation=/sites/eforms/aaad%20forms%20v3/" & itemAutoID & ".xml&amp;OpenIn=Browser" & "</my:info_url>" & line.Substring(line.IndexOf("</my:info_url>") + 14)

                                End If

                                writer.WriteLine(line)

                                filename = ""
                                itemAutoID = ""
                                line = ""

                                writer.Dispose()

                            Catch ex As Exception

                            End Try

                        Loop

                    End If ' Ends the OK of DirectoryBrowser

                End If ' Ends the OK of FileBrowser

            End If ' Ends the OK of DirectoryBrowser

        Catch ex2 As Exception

        End Try

        btnExit.Enabled = True

    End Sub


    Private Sub updateFormsv4()

        btnExit.Enabled = False

        Try

            inputFolder.ShowNewFolderButton = False
            inputFolder.Description = "Please choose the folder where the outdated forms are"

            If inputFolder.ShowDialog() = DialogResult.OK Then

                outputFolder.ShowNewFolderButton = False
                outputFolder.Description = "Please choose the folder where the UPDATED forms will be saved"

                If outputFolder.ShowDialog() = DialogResult.OK Then

                    Dim dirInfo As New DirectoryInfo(inputFolder.SelectedPath)

                    For Each fi As FileInfo In dirInfo.GetFiles()

                        Dim filename As String = fi.Name
                        'Dim filename As String = "555.xml"

                        Try

                            Dim reader As StreamReader = New StreamReader(inputFolder.SelectedPath & "\" & filename)



                            Dim line As String = reader.ReadToEnd()
                            reader.Dispose()

                            If line.IndexOf("<my:aries_posting_title>false</my:aries_posting_title>") > 0 And line.IndexOf("<my:aries_posting_req_number>false</my:aries_posting_req_number>") > 0 Then

                                rtb.AppendText("Writing " & outputFolder.SelectedPath & "\" & filename & Chr(13))
                                line = line.Replace("<my:aries_posting_title>false</my:aries_posting_title>", "<my:aries_posting_title></my:aries_posting_title>").Replace("<my:aries_posting_req_number>false</my:aries_posting_req_number>", "<my:aries_posting_req_number></my:aries_posting_req_number>")

                                Dim writer As New StreamWriter(outputFolder.SelectedPath & "\" & filename)
                                writer.Write(line)
                                writer.Dispose()

                            End If

                            line = ""
                            filename = ""

                        Catch ex As Exception

                            Dim test As String
                            test = "error"

                        End Try

                        'Loop

                    Next fi

                    'End If ' Ends the OK of DirectoryBrowser

                End If ' Ends the OK of FileBrowser

            End If ' Ends the OK of DirectoryBrowser

        Catch ex2 As Exception

        End Try

        btnExit.Enabled = True

    End Sub


End Class
