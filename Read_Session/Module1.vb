Imports Microsoft.Office.Interop
Imports System.Text.RegularExpressions

Module Module1
    Sub Main()
        Dim objExcel As Excel.Application       'the file we're going to read from
        Dim conn As ADODB.Connection
        Dim files_read As Integer
        Dim new_files As Integer
        Dim successful_adds As Integer
        Dim error_format As Integer
        Dim error_content As Integer
        Dim results As Integer()
        Dim sSql
        Dim rec As ADODB.Recordset
        Dim start_time As DateTime
        Dim end_time As DateTime
        Dim elapsed_time As Long
        Dim elapsed_time_hours As Long
        Dim elapsed_time_minutes As Long
        Dim elapsed_time_seconds As Long
        Dim folder_path As String
        Dim dev_mode As Boolean
        Dim demo_mode As Boolean
        Dim reset As Boolean
        Dim db As String

        objExcel = New Excel.Application
        conn = New ADODB.Connection
        files_read = 0
        new_files = 0
        successful_adds = 0
        error_format = 0
        error_content = 0
        results = {0, 0, 0, 0, 0, 0}
        sSql = ""
        rec = New ADODB.Recordset
        start_time = Now()
        end_time = Now()
        elapsed_time = 0
        elapsed_time_hours = 0
        elapsed_time_minutes = 0
        elapsed_time_seconds = 0
        folder_path = ""
        dev_mode = True
        reset = False
        demo_mode = False
        db = ""

        If dev_mode = True Then
            folder_path = "\\sharepoint.washington.edu@SSL\DavWWWRoot\oim\proj\HRPayroll\Imp\Supervisory Org Cleanup\Role_mapping_2\Submittals"
            conn.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Role_Mapping_2\Session_responses_2.accdb")
            If reset = True Then
                sSql = "DELETE * FROM " & db & " submittals"
                'Debug.WriteLine(sSql)
                'conn.Execute(sSql)
                sSql = "DELETE * FROM rejected_" & db & " submittals"
                'Debug.WriteLine(sSql)
                'conn.Execute(sSql)
                sSql = "DELETE * FROM responses"
                'Debug.WriteLine(sSql)
                'conn.Execute(sSql)
                sSql = "DELETE * FROM unit_correction_map"
                'Debug.WriteLine(sSql)
                'conn.Execute(sSql)
                sSql = "DELETE * FROM org_team_map"
                'Debug.WriteLine(sSql)
                'conn.Execute(sSql)
            End If
        Else
            folder_path = "\\sharepoint.washington.edu@SSL\DavWWWRoot\oim\proj\HRPayroll\Imp\Supervisory Org Cleanup\Role-Mapping"
            conn.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\submissions\Session_responses.accdb")
        End If

        results = process_Folder(objExcel, conn, folder_path, demo_mode)

        'test a single file
        'Process_workbook("20151110-160531 Working in Workday Session 1 - Data Collection Tool College of Arts And Sciences_Linguistics" & ".xlsm", objExcel, conn)    'Session 1 test
        'Process_workbook("20151118- 1009 HFS - Session 3 - Data Collection - Working in Workday" & ".xlsx", objExcel, conn)    'Session 3 test
        'Process_workbook("20151118- 1009 HFS - Session 2 - Data Collection - Working in Workday" & ".xlsx", objExcel, conn)    'Session 2 test

        'generate_excel_report(objExcel, conn, "Workday_Role_Mapping", "", demo_mode)  'file name, '

        'generate_error_report(objExcel, conn, "Error_report", "")

        'initiate_unit_reports(objExcel, conn, "C:\submissions\Unit Reports\", demo_mode)

        'generate_unit_report(objExcel, conn, "Housing and Food Services (Housing and Food Services)", "C:\submissions\Unit Reports\", demo_mode)  'file name, '
        'generate_unit_report(objExcel, conn, "Workday_Role_Mapping", "Applied Physics Lab (Applied Physics Lab)", 7)  'file name, 

        conn.Close()

        If demo_mode = False Then
            objExcel.Quit()
        End If

        end_time = Now()

        elapsed_time_hours = DateDiff("h", start_time, end_time)
        elapsed_time_minutes = DateDiff("m", start_time, end_time) Mod 60
        elapsed_time_seconds = DateDiff("s", start_time, end_time) Mod 60

        Debug.WriteLine("Read " & results(0) & " files in folder in " & elapsed_time_hours & ":" & elapsed_time_minutes & ":" & elapsed_time_seconds & " seconds")
        Console.WriteLine("Read " & results(0) & " files in folder in " & elapsed_time_seconds & " seconds.")
        Debug.WriteLine(results(1) & " new files identified.")
        Debug.WriteLine(results(2) & " files successfully added.")
        Debug.WriteLine(results(3) & " files Not added.")
        Debug.WriteLine(results(4) & " files contain format errors.")
        Debug.WriteLine(results(5) & " files contain content errors.")

    End Sub

    Function process_Folder(objExcel, conn, folderpath, demo_mode) As Integer()

        'results(0) = 'Total number of files read
        'results(1) = 'Total number of new files
        'results(2) = 'Files not added due to format
        'results(3) = 'Files not added due to content

        Dim sSql
        Dim file_count = 0
        Dim folder_count = 0
        Dim rec As ADODB.Recordset
        Dim FileNameWithExt
        Dim filenames
        Dim results As Integer()
        Dim workbook_results As Integer()
        Dim successful_adds As Integer = 0
        Dim error_format As Integer = 0
        Dim error_content As Integer = 0
        Dim blank_fields As Integer = 0
        Dim blank_fields_academic As Integer = 0
        Dim error_format_string = ""
        Dim error_content_string = ""
        Dim not_added = 0
        Dim debug_state = False

        results = {0, 0, 0, 0, 0, 0}         'Total number of files in folder, Files not already input, files sucessful added, files not added, Files containing format errors, Files containing content errors
        workbook_results = {0, 0, 0, 0, 0}      'successful_reads, add_attempted, not_added, error_format, error_content

        Try
            filenames = My.Computer.FileSystem.GetFiles(folderpath, FileIO.SearchOption.SearchTopLevelOnly)

        Catch ex As System.IO.IOException
            Debug.WriteLine(
            "{0}: The write operation could " &
            "not be performed because the " &
            "specified part of the file is " &
            "locked.", ex.GetType().Name)
            MsgBox("Please ensure that you have access to " & folderpath &
                    " on sharepoint.")

        End Try

        rec = New ADODB.Recordset

        For Each fileName As String In filenames
            'Debug.WriteLine(fileName)
            folder_count = folder_count + 1
            FileNameWithExt = Mid$(fileName, InStrRev(fileName, "\") + 1)
            'Debug.WriteLine(FileNameWithExt)
            sSql = "Select submittalID from " & db & " submittals where file_name = """ & FileNameWithExt & """"
            'Debug.WriteLine(sSql)
            rec.Open(sSql, conn)

            If (rec.BOF And rec.EOF) Then                                                   'if the file name has not been recorded
                Debug.WriteLine("Processing " & file_count + 1 & ":  " & folderpath & "\" & FileNameWithExt & "...")
                workbook_results = Process_workbook(FileNameWithExt, folderpath, objExcel, conn, demo_mode)

                If (workbook_results(0) > 0) Then
                    successful_adds = successful_adds + 1
                End If
                If (workbook_results(2) = 1) Then
                    not_added = not_added + 1
                End If

                If (workbook_results(3) = 1) Then
                    error_format = error_format + 1
                End If
                If (workbook_results(4) = 1) Then
                    error_content = error_content + 1
                End If
                file_count = file_count + 1
            End If
            rec.Close()

        Next

        results(0) = folder_count   'Total number of files in folder
        results(1) = file_count     'Files not already input 
        results(2) = successful_adds  'files sucessful added
        results(3) = not_added          'files not added
        results(4) = error_format        'Files containing format errors
        results(5) = error_content       'Files containing content errors


        Return results

        If debug_state = True Then
            Debug.WriteLine("Files in folder: " & results(0))
            Debug.WriteLine("Files not already input: " & results(1))
            Debug.WriteLine("Files successfully added: " & results(2))
            Debug.WriteLine("Files not added: " & results(3))
            Debug.WriteLine("Files containing formatting errors: " & results(4))
            Debug.WriteLine("Files containing content errors: " & results(5))

        End If

        sSql = Nothing
        file_count = Nothing
        folder_count = Nothing
        rec = Nothing
        FileNameWithExt = Nothing
        folderpath = Nothing
        filenames = Nothing
        results = Nothing
        workbook_results = Nothing
        successful_adds = Nothing
        error_format = Nothing
        error_content = Nothing
        blank_fields = Nothing
        blank_fields_academic = Nothing
        error_format_string = Nothing
        error_content_string = Nothing
        not_added = Nothing
        debug_state = Nothing

    End Function

    Function Process_workbook(filename, folder_path, objExcel, conn, demo_mode) As Integer()

        'workbook_results(0)    = 1 File was inserted 0 file was not inserted
        'workbook_results(1)    = There was an error in format - number of worksheets was not what was expected
        'workbook_results(2)    = There was an error of content - identifying information missing
        'workbook_results(3)    = A count of blank fields
        'workbook_results(4)    = A count of blank fields related to academic Scenarios

        Dim excelPath As String
        Dim pathToFile As String
        Dim worksheet
        Dim workbook
        Dim sSql As String
        Dim unit As String
        Dim contact As String
        Dim date_submitted As String
        Dim submittalID As Integer
        Dim worksheetCount As Integer
        Dim error_conditions As Integer
        Dim file_ext As String
        Dim session_no As Integer
        Dim Error_identifying_information = False
        Dim Error_file_type = False
        Dim successful_adds As Integer
        Dim error_format_ct As Integer
        Dim error_content_ct As Integer
        Dim error_content_bool As Boolean
        Dim error_format_bool As Boolean
        Dim rec As ADODB.Recordset
        Dim debug_state As Boolean
        Dim add_attempted As Integer
        Dim not_added As Integer
        Dim unit_map_ID As Integer
        Dim workbook_results As Integer()
        Dim process_scenario_results As Integer()

        excelPath = ""
        pathToFile = ""
        'worksheet
        'workbook
        sSql = ""
        unit = ""
        contact = ""
        date_submitted = ""
        submittalID = 0
        worksheetCount = 0
        error_conditions = 0
        file_ext = ""
        session_no = 0
        Error_identifying_information = False
        Error_file_type = False
        successful_adds = 0
        error_format_ct = 0
        error_content_ct = 0
        error_content_bool = False
        error_format_bool = False
        rec = New ADODB.Recordset
        debug_state = False
        add_attempted = 0
        not_added = 0
        unit_map_ID = 0
        workbook_results = {0, 0, 0, 0, 0}      'successful_reads, add_attempted, not_added, error_format, error_content
        process_scenario_results = {0, 0, 0}

        pathToFile = folder_path & "\" & filename

        file_ext = Mid$(filename, InStrRev(filename, ".") + 1)
        'Debug.WriteLine(file_ext)

        If debug_state = True Then
            Debug.WriteLine("Process_workbook()...")
            objExcel.Visible = True
            objExcel.DisplayAlerts = 0 ' Don't display any messages about conversion and so forth
        End If

        If demo_mode = True Then
            objExcel.Visible = True
            System.Threading.Thread.CurrentThread.Sleep(500)
        End If

        Try
            workbook = objExcel.Workbooks.Open(pathToFile)
        Catch ex As Exception
            Debug.WriteLine(pathToFile & " couldn't be opened.")
        End Try

        If Not IsNothing(workbook) Then
            'workbook = objExcel.ActiveWorkbook
            worksheet = workbook.Worksheets(1)
            If demo_mode Then
                worksheet.Activate
                System.Threading.Thread.CurrentThread.Sleep(500)
            End If

            'Collect Identifying info
            Dim start_row = 4
            Dim start_col = 2

            If demo_mode Then
                worksheet.Cells(start_row, start_col).Select
                System.Threading.Thread.CurrentThread.Sleep(500)
            End If
            Do Until worksheet.Cells(start_row, start_col).Value = ""
                If worksheet.Cells(start_row, start_col).Value = "Unit: " _
                    Or worksheet.Cells(start_row, start_col).Value = "Organization for which this was completed" Then
                    start_col = start_col + 1
                    If demo_mode Then
                        worksheet.Cells(start_row, start_col).Select
                        System.Threading.Thread.CurrentThread.Sleep(500)
                    End If
                Else
                    unit = worksheet.Cells(4, start_col).Value
                    If demo_mode Then
                        worksheet.Cells(4, start_col).Select
                        System.Threading.Thread.CurrentThread.Sleep(250)
                    End If
                    contact = worksheet.Cells(6, start_col).Value
                    If demo_mode Then
                        worksheet.Cells(6, start_col).Select
                        System.Threading.Thread.CurrentThread.Sleep(250)
                    End If
                    date_submitted = worksheet.Cells(8, start_col).Value
                    If demo_mode Then
                        worksheet.Cells(8, start_col).Select
                        System.Threading.Thread.CurrentThread.Sleep(250)
                    End If
                    start_col = start_col + 1
                End If
                If demo_mode Then
                    System.Threading.Thread.CurrentThread.Sleep(500)
                End If
            Loop

            If IsNothing(unit) Then
                error_content_bool = True
                error_content_ct = error_content_ct + 1
            End If

            session_no = 0

            worksheetCount = workbook.Worksheets.Count

            If worksheetCount > 1 Then
                worksheet = objExcel.ActiveWorkbook.Worksheets(2)
                If demo_mode Then
                    worksheet.Activate
                    System.Threading.Thread.CurrentThread.Sleep(500)
                End If
                If Left(worksheet.Name, 2) = "1 " Then                                             'Session 1
                    session_no = 1
                    file_ext = "xlsm"
                ElseIf Left(worksheet.Name, 2) = "5A" Then                           'Session 2
                    session_no = 2
                    'Debug.WriteLine(Left(workbook.Worksheets(4).Name, 2))
                    If Left(workbook.Worksheets(4).Name, 2) = "6A" Then
                        file_ext = "xlsm"
                    Else
                        file_ext = "xlsx"
                    End If
                ElseIf Left(worksheet.Name, 2) = "9A" Then                                          'Session 3
                    session_no = 3
                    If Left(workbook.Worksheets(3).Name, 2) = "9B" Then
                        file_ext = "xlsx"
                    Else
                        file_ext = "xlsm"
                    End If
                End If
            Else
                error_format_bool = True
                error_format_ct = error_format_ct + 1
            End If

            'Debug.WriteLine(file_ext)
            'Debug.WriteLine(session_no)

            If error_content_bool = False And error_format_bool = False And session_no <> 0 Then
                'Debug.WriteLine("Format and content check complete.  processing scenarios")

                'Prepare Unit for verification
                sSql = "SELECT unit_map_ID from unit_correction_map where reported_unit = """ & unit & """"
                'Debug.WriteLine(sSql)
                rec.Open(sSql, conn)

                Dim i = 0
                If (rec.BOF And rec.EOF) Then
                    rec.Close()
                    sSql = "INSERT INTO unit_correction_map (reported_unit) VALUES (""" & unit & """)"
                    'Debug.WriteLine(sSql)
                    conn.Execute(sSql)

                    'Identify the new record's ID
                    sSql = "Select max(unit_map_ID) FROM unit_correction_map"
                    rec.Open(sSql, conn)
                    For Each x In rec.Fields
                        unit_map_ID = x.value
                    Next
                    rec.Close()
                Else
                    For Each x In rec.Fields
                        If Not IsDBNull(x.value) Then
                            unit_map_ID = x.value
                        End If
                    Next
                    rec.Close()
                End If


                'Add a record
                sSql = "INSERT INTO " & db & " submittals (reported_unit, contact, date_submitted, date_recorded, file_name, session_no, unit_map_ID) VALUES (""" &
                    unit & """, """ & contact & """, """ & date_submitted & """, """ & Format(Now, "MM/dd/yyyy") & """, """ & filename & """,""" & session_no.ToString & """,""" & unit_map_ID & """)"
                'Debug.WriteLine(sSql)
                conn.Execute(sSql)

                'Identify the new record's ID
                sSql = "Select max(submittalID) FROM " & db & " submittals"
                rec.Open(sSql, conn)
                For Each x In rec.Fields
                    submittalID = x.value
                Next
                rec.Close()

                process_scenario_results = Process_scenarios(objExcel, workbook, conn, submittalID, file_ext, session_no, demo_mode)

                workbook_results(0) = process_scenario_results(0)   'Successful_field_reads
                'workbook_results(3) = process_scenario_results(1)   'blank_field_ct_non-academic
                'workbook_results(4) = process_scenario_results(2)   'blank_field_ct_academic
                add_attempted = add_attempted + 1

                sSql = "DELETE * FROM " & db & " rejected_submittals WHERE filename = """ & filename & """"
                'Debug.WriteLine(sSql)
                conn.Execute(sSql)

            Else
                Debug.WriteLine("...Format and content check returned errors.  See error log for more information.")
                not_added = not_added + 1

                sSql = "Select rejected_submittalID from " & db & " rejected_submittals where filename = """ & filename & """"
                'Debug.WriteLine(sSql)
                rec.Open(sSql, conn)

                If (rec.BOF And rec.EOF) Then                                                   'if the file name has not been recorded
                    sSql = "INSERT INTO " & db & " rejected_submittals (filename, content_error, format_error) values (""" & filename & """, " & error_content_bool & ", " & error_format_bool & ")"
                    Debug.WriteLine(sSql)
                    conn.Execute(sSql)
                End If
                rec.Close()

            End If

            If demo_mode = True Then
                System.Threading.Thread.CurrentThread.Sleep(500)
            End If

            Try
                workbook.Close()
            Catch ex As Exception
                Debug.WriteLine("Could not close workbook.")
            End Try

            workbook = Nothing
            worksheet = Nothing

        End If

        workbook_results(1) = add_attempted
        workbook_results(2) = not_added
        workbook_results(3) = error_format_ct                  'There was an error in format - number of worksheets was not what was expected
        workbook_results(4) = error_content_ct                 'There was an error of content - identifying information missing

        If debug_state = True Then
            Debug.WriteLine("Successful_field_reads: " & workbook_results(0))
            Debug.WriteLine("add_attempted: " & workbook_results(1))
            Debug.WriteLine("not_added: " & workbook_results(2))
            Debug.WriteLine("error_format: " & workbook_results(3))
            Debug.WriteLine("error_content: " & workbook_results(4))
        End If

        'Debug.WriteLine("..." & filename & " processing completed.")

        Return workbook_results

        excelPath = Nothing
        pathToFile = Nothing
        worksheet = Nothing
        workbook = Nothing
        sSql = Nothing
        unit = Nothing
        contact = Nothing
        date_submitted = Nothing
        submittalID = Nothing
        worksheetCount = Nothing
        error_conditions = Nothing
        file_ext = Nothing
        session_no = Nothing
        Error_identifying_information = Nothing
        Error_file_type = Nothing
        process_scenario_results = Nothing
        successful_adds = Nothing
        error_format_ct = Nothing
        error_content_ct = Nothing
        error_content_bool = Nothing
        error_format_bool = Nothing
        workbook_results = Nothing
        rec = Nothing
        debug_state = Nothing
        add_attempted = Nothing
        not_added = Nothing

    End Function

    Function Process_scenarios(objExcel, workbook, conn, submittalID, file_ext, session_no, demo_mode) As Integer()

        'process_scenarios(0) = Count of successful scenarios
        'process_scenarios(1) = Blank Field Count
        'process_scenarios(1) = Blank Field Count Academic


        Dim successful_field_reads = 0
        Dim blank_field_txt = ""
        Dim blank_field_txt_academic = ""
        Dim blank_field_ct = 0
        Dim blank_field_ct_academic = 0
        Dim file_structure_issue = ""
        Dim sSql
        Dim process_scenarios_results As Integer()
        Dim read_field_results As String()
        Dim worksheet_name_error = ""
        Dim worksheet_orient_error = ""
        Dim index = 0
        Dim debug_state = False

        process_scenarios_results = {0, 0, 0}
        read_field_results = {"", "", "", "", "", "", ""} 'Data_found_ct, blank_field_txt, blank_field_ct, blank_field_academic_txt, blank_field_ct_academic, worksheet_name_error, worksheet_orient_error

        If debug_state = True Then
            Debug.WriteLine("Processing Scenarios...")
            objExcel.Visible = True
        End If

        If file_ext = "xlsm" Then                                                                  'non-mac formatted
            If session_no = 1 Then                                                                 'Session 1
                Dim field_definition(0 To 6) As String                              'xlsm session 1
                field_definition(0) = "2,1 W,1,2C,2C,6,3"
                field_definition(1) = "2,1 W,1,3B,3B,12,3"
                field_definition(2) = "3,2 T,2,2A,2A,6,3"
                'Worksheet 4 (3-Time Off) is blank
                field_definition(3) = "5,4A ,4A,2A,2A,6,3"
                field_definition(4) = "6,4B ,4B,2A,2A,6,3"
                field_definition(5) = "6,4B ,4B,3A,3A,12,3"
                field_definition(6) = "7,4C ,4C,3A,3A,6,3"
                read_field_results = read_field(objExcel, field_definition, workbook, conn, submittalID, demo_mode)
            ElseIf session_no = 2 Then                                                             'Session 2
                Dim field_definition(0 To 20) As String                             'xlsm session 2
                field_definition(0) = "2,5A ,5A,2B,2B,5,3"
                field_definition(1) = "2,5A ,5A,4A,4A,11,3"
                field_definition(2) = "3,5B ,5B,2B,2B,5,3"
                field_definition(3) = "3,5B ,5B,3A,3A,11,3"
                field_definition(4) = "3,5B ,5B,4A,4A,17,3"
                field_definition(5) = "4,6A ,6A,2A,2A,5,3"
                field_definition(6) = "4,6A ,6A,3A,3A,11,3"
                field_definition(7) = "5,6B ,6B,1A,1A,5,3"
                field_definition(8) = "5,6B ,6B,2A,2A,11,3"
                field_definition(9) = "5,6B ,6B,3A,3A,17,3"
                field_definition(10) = "6,7 O,7,4A,4A,5,3"
                field_definition(11) = "7,8A ,8A,1A,1A,5,3"
                field_definition(12) = "7,8A ,8A,2A,2A,11,3"
                field_definition(13) = "7,8A ,8A,2B,2B,17,3"
                field_definition(14) = "7,8A ,8A,3A,3A,23,3"
                field_definition(15) = "8,8B ,8B,1B,1B,5,3"
                field_definition(16) = "8,8B ,8B,2A,2A,11,3"
                field_definition(17) = "8,8B ,8B,3A,3A,17,3"
                field_definition(18) = "8,8B ,8B,4A,4A,23,3"
                field_definition(19) = "9,8C ,8C,2A,2A,5,3"
                field_definition(20) = "9,8C ,8C,3A,3A,11,3"
                read_field_results = read_field(objExcel, field_definition, workbook, conn, submittalID, demo_mode)
            ElseIf session_no = 3 Then                                                             'Session 3
                Dim field_definition(0 To 14) As String                             'xlsm session 3
                field_definition(0) = "2,9A ,9A,2A,2A,4,3"
                field_definition(1) = "2,9A ,9A,3A,3A,10,3"
                'worksheet 3 ("3 Time Off") is blank
                field_definition(2) = "4,9B ,9B,2B,2B,4,3"
                field_definition(3) = "5,10 ,10,2A,2A,4,3"
                field_definition(4) = "5,10 ,10,3A,3A,10,3"
                field_definition(5) = "6,11A,11A,2A,2A,4,3"
                field_definition(6) = "6,11A,11A,3A,3A,10,3"
                field_definition(7) = "6,11A,11A,3B,3B,16,3"
                field_definition(8) = "6,11A,11A,4A,4A,22,3"
                field_definition(9) = "7,11B,11B,2A,2A,4,3"
                field_definition(10) = "7,11B,11B,3A,3A,10,3"
                field_definition(11) = "7,11B,11B,3B,3B,16,3"
                field_definition(12) = "8,12 ,12,2A,2A,4,3"
                field_definition(13) = "8,12 ,12,3A,3A,10,3"
                field_definition(14) = "8,12 ,12,4A,4A,16,3"
                read_field_results = read_field(objExcel, field_definition, workbook, conn, submittalID, demo_mode)
            ElseIf session_no = 0 Then
                file_structure_issue = "x"
            End If
        ElseIf file_ext = "xlsx" Then
            If session_no = 1 Then                                                                 'Session 1
                Dim field_definition(0 To 6) As String                          'xlsx session 1
                field_definition(0) = "2,1 W,1,2C,2C,6,3"
                field_definition(1) = "2,1 W,1,3B,3B,12,3"
                field_definition(2) = "3,2 T,2,2A,2A,6,3"
                'Worksheet 4 (3-Time Off) is blank
                field_definition(3) = "5,4A ,4A,2A,2A,6,3"
                field_definition(4) = "6,4B ,4B,2A,2A,6,3"
                field_definition(5) = "6,4B ,4B,3A,3A,12,3"
                field_definition(6) = "7,4C ,4C,3A,3A,6,3"
                read_field_results = read_field(objExcel, field_definition, workbook, conn, submittalID, demo_mode)
            ElseIf session_no = 2 Then                                                             'Session 2
                Dim field_definition(0 To 20) As String                         'xlsx session 2
                field_definition(0) = "2,5A ,5A,2B,2B,5,3"
                field_definition(1) = "2,5A ,5A,4A,4A,11,3"

                field_definition(2) = "3,5B ,5B,2B,2B,6,3"
                field_definition(3) = "3,5B ,5B,3A,3A,12,3"
                field_definition(4) = "3,5B ,5B,4A,4A,18,3"

                'worksheet 4 contains time off information,is blank

                field_definition(5) = "5,6B ,6A,2A,2A,6,3"    'Typo on tab of data collection tool
                field_definition(6) = "5,6B ,6A,3A,3A,12,3"   'Typo on tab of data collection tool

                field_definition(7) = "6,6B ,6B,1A,1A,6,3"
                field_definition(8) = "6,6B ,6B,2A,2A,12,3"
                field_definition(9) = "6,6B ,6B,3A,3A,18,3"

                field_definition(10) = "7,7 O,7,4A,3A,6,3"

                field_definition(11) = "8,8A ,8A,1A,1A,6,3"
                field_definition(12) = "8,8A ,8A,2A,2A,12,3"
                field_definition(13) = "8,8A ,8A,2B,2B,18,3"
                field_definition(14) = "8,8A ,8A,3A,3A,24,3"

                field_definition(15) = "9,8B ,8B,1B,1B,6,3"
                field_definition(16) = "9,8B ,8B,2A,2A,12,3"
                field_definition(17) = "9,8B ,8B,3A,3A,18,3"
                field_definition(18) = "9,8B ,8B,4A,4A,24,3"

                field_definition(19) = "9,8C ,8C,2A,2A,5,3"
                field_definition(20) = "9,8C ,8C,3A,3A,11,3"
                read_field_results = read_field(objExcel, field_definition, workbook, conn, submittalID, demo_mode)
            ElseIf session_no = 3 Then                                                             'Session 3
                Dim field_definition(0 To 14) As String                      'xlsx Session 3
                field_definition(0) = "2,9A ,9A,2A,2A,5,3"
                field_definition(1) = "2,9A ,9A,3A,3A,11,3"

                field_definition(2) = "3,9B ,9B,2B,2A,6,3"

                'worksheet 5 ("3 Time off") is blank

                field_definition(3) = "3,10 ,10,2A,2A,6,3"
                field_definition(4) = "3,10 ,10,3A,3A,12,3"

                field_definition(5) = "6,11A,11A,2A,2A,6,3"
                field_definition(6) = "6,11A,11A,3A,3A,12,3"
                field_definition(7) = "6,11A,11A,3B,3B,18,3"
                field_definition(8) = "6,11A,11A,4A,4A,24,3"

                field_definition(9) = "7,11B,11B,2A,2A,6,3"
                field_definition(10) = "7,11B,11B,3A,3A,12,3"
                field_definition(11) = "7,11B,11B,3B,3B,18,3"

                field_definition(12) = "8,12 ,12,2A,2A,6,3"
                field_definition(13) = "8,12 ,12,3A,3A,12,3"
                field_definition(14) = "8,12 ,12,4A,4A,18,3"
                read_field_results = read_field(objExcel, field_definition, workbook, conn, submittalID, demo_mode)
            ElseIf session_no = 0 Then
                file_structure_issue = "x"
            End If
        Else
            'Debug.WriteLine("The file was either Not an xlsm Or xlsx.")
        End If

        successful_field_reads = CInt(read_field_results(0))
        blank_field_txt = read_field_results(1)
        blank_field_ct = CInt(read_field_results(2))
        blank_field_txt_academic = read_field_results(3)
        blank_field_ct_academic = CInt(read_field_results(4))
        worksheet_name_error = read_field_results(5)
        worksheet_orient_error = read_field_results(6)

        If debug_state = True Then
            Debug.WriteLine("Sucessful Reads:" & successful_field_reads)
            Debug.WriteLine("blank_field_txt:" & blank_field_txt)
            Debug.WriteLine("blank_field_ct:" & blank_field_ct)
            Debug.WriteLine("blank_field_txt_academic:" & blank_field_txt_academic)
            Debug.WriteLine("blank_field_ct_academic: " & blank_field_ct_academic)
            Debug.WriteLine("worksheet_name_error:" & worksheet_name_error)
            Debug.WriteLine("worksheet_orient_error: " & worksheet_orient_error)
        End If

        If blank_field_ct > 0 Then
            sSql = "UPDATE " & db & " submittals SET blank_fields_non_academic = """ & blank_field_txt & """ WHERE submittalID = " & submittalID
            'Debug.WriteLine(sSql)
            conn.Execute(sSql)

        End If

        If blank_field_ct_academic > 0 Then
            sSql = "UPDATE " & db & " submittals SET blank_fields_academic = """ & blank_field_txt_academic & """ WHERE submittalID = " & submittalID
            'Debug.WriteLine(sSql)
            conn.Execute(sSql)

        End If

        If worksheet_name_error <> "Worksheet name errors: (expected):(encountered);" Then
            sSql = "UPDATE " & db & " submittals SET worksheet_name_error = """ & worksheet_name_error & """ WHERE submittalID = " & submittalID
            'Debug.WriteLine(sSql)
            conn.Execute(sSql)
        End If

        If worksheet_orient_error <> "Orient cell errors: (s):(f):(oc);" Then
            sSql = "UPDATE " & db & " submittals SET worksheet_orient_error = """ & worksheet_orient_error & """ WHERE submittalID = " & submittalID
            'Debug.WriteLine(sSql)
            conn.Execute(sSql)
        End If


        process_scenarios_results(0) = successful_field_reads
        process_scenarios_results(1) = blank_field_ct
        process_scenarios_results(2) = blank_field_ct_academic

        Return process_scenarios_results

        successful_field_reads = Nothing
        blank_field_txt = Nothing
        blank_field_txt_academic = Nothing
        blank_field_ct = Nothing
        blank_field_ct_academic = Nothing
        file_structure_issue = Nothing
        sSql = Nothing
        process_scenarios_results = Nothing
        read_field_results = Nothing
        worksheet_name_error = Nothing
        worksheet_orient_error = Nothing
        index = Nothing
        debug_state = Nothing

    End Function

    Function read_field(ObjExcel, field_definition, workbook, conn, submittalID, demo_mode) As String()

        'Returns a file field string array

        'read_field_results(0) = data_found_ct.ToString             'The number of field entries found
        'read_field_results(1) = blank_field_txt                    'A string of non-academic blank fields
        'read_field_results(2) = blank_field_ct.ToString            'a count of non-academic fields
        'read_field_results(3) = blank_field_txt_academic           'a string of academic blank fields
        'read_field_results(4) = blank_field_ct_academic.ToString   'a count of academic blank fields

        Dim foo
        Dim index As Integer
        Dim worksheet As Integer
        Dim worksheetName As String
        Dim scenario As String
        Dim orient_cell As String
        Dim startRow As Integer
        Dim startCol As Integer
        Dim data_found_ct As Integer
        Dim blank_field_txt_academic As String
        Dim blank_field_txt As String
        Dim blank_field_ct As Integer
        Dim blank_field_ct_academic As Integer
        Dim worksheet_name_error As String
        Dim worksheet_orient_error As String
        Dim read_field_results As String()
        Dim collect_field_results As String()

        index = 0
        worksheet = 0
        worksheetName = ""
        scenario = ""
        orient_cell = ""
        startRow = 0
        startCol = 0
        data_found_ct = 0
        blank_field_txt_academic = ""
        blank_field_txt = ""
        blank_field_ct = 0
        blank_field_ct_academic = 0
        worksheet_name_error = "Worksheet name errors: (expected):(encountered);"
        worksheet_orient_error = "Orient cell errors: (s):(f):(oc);"
        read_field_results = {"", "", "", "", "", "", ""} 'Data_found_ct, blank_field_txt, blank_field_ct, blank_field_academic_txt, blank_field_ct_academic, worksheet_name_error, worksheet_orient_error
        collect_field_results = {"", "", ""}
        Dim debug_state = False

        If debug_state = True Then
            Debug.WriteLine("Reading fields for submittalID " & submittalID)
        End If

        For Each field In field_definition
            foo = Split(field, ",")

            worksheet = CInt(foo(0))
            worksheetName = foo(1)
            scenario = foo(2)
            field = foo(3)
            orient_cell = foo(4)
            startRow = CInt(foo(5))
            startCol = CInt(foo(6))

            collect_field_results = collect_field(ObjExcel, workbook, conn, submittalID, worksheet, worksheetName, scenario, field, orient_cell, startRow, startCol, demo_mode)

            If CInt(collect_field_results(0)) = 0 Then
                If scenario = "4C" _
                    Or scenario = "5B" _
                       Or scenario = "6B" _
                       Or scenario = "8B" _
                       Or scenario = "11A" _
                       Or scenario = "11B" Then
                    blank_field_txt_academic = blank_field_txt_academic & " " & foo(2) & ":" & foo(3) & ";"
                    blank_field_ct_academic = blank_field_ct_academic + 1
                Else
                    blank_field_txt = blank_field_txt & " " & foo(2) & ":" & foo(3) & ";"
                    blank_field_ct = blank_field_ct + 1
                End If


            Else
                data_found_ct = data_found_ct + CInt(collect_field_results(0))
            End If

            worksheet_name_error = worksheet_name_error & collect_field_results(1)
            worksheet_orient_error = worksheet_orient_error & collect_field_results(1)

            index = index + 1
        Next

        read_field_results(0) = data_found_ct.ToString
        read_field_results(1) = blank_field_txt
        read_field_results(2) = blank_field_ct.ToString
        read_field_results(3) = blank_field_txt_academic
        read_field_results(4) = blank_field_ct_academic.ToString
        read_field_results(5) = worksheet_name_error
        read_field_results(6) = worksheet_orient_error

        If debug_state = True Then
            Debug.WriteLine("Read Field: Cound of Data Found: " & read_field_results(0))
            Debug.WriteLine("Read Field: Blank_field_txt: " & read_field_results(1))
            Debug.WriteLine("Read Field: Blank_field_ct: " & read_field_results(2))
            Debug.WriteLine("Read Field: blank_field_academic: " & read_field_results(3))
            Debug.WriteLine("Read Field: blank_field_ct_academic: " & read_field_results(4))
            Debug.WriteLine("Read Field: worksheet name error: " & read_field_results(5))
            Debug.WriteLine("Read Field: worksheet orient error: " & read_field_results(6))
        End If

        Return read_field_results

        foo = Nothing
        index = Nothing
        worksheet = Nothing
        worksheetName = Nothing
        scenario = Nothing
        orient_cell = Nothing
        startRow = Nothing
        startCol = Nothing
        collect_field_results = Nothing
        blank_field_txt_academic = Nothing
        blank_field_txt = Nothing
        blank_field_ct = Nothing
        blank_field_ct_academic = Nothing
        worksheet_name_error = Nothing
        worksheet_orient_error = Nothing
        read_field_results = Nothing
        data_found_ct = Nothing
        debug_state = Nothing

    End Function

    Private Function collect_field(objExcel, workbook, conn, submittalID, worksheet, worksheetName, scenario, field, orient_cell, startRow, startCol, demo_mode) As String()

        'Returns the number of field entries encountered.  Blank if 0

        Dim curRow
        Dim curCol
        Dim currentWorkSheet
        Dim first_name
        Dim last_name
        Dim eid
        Dim org_team As String
        Dim budget_no
        Dim sSql
        Dim rec As ADODB.Recordset
        Dim responseID As Integer
        Dim org_team_mapID
        Dim index
        Dim debug_state = False
        Dim worksheet_name_error = ""
        Dim worksheet_orient_error = ""
        Dim results = {"", "", ""}   'index, worksheet_name_error, worksheet_orient_error
        Dim r
        Dim i

        index = 0
        org_team_mapID = 0

        rec = New ADODB.Recordset

        Try
            currentWorkSheet = workbook.Worksheets(worksheet)
        Catch ex As Exception
            Debug.WriteLine("worksheet " & worksheet & "Not found.")
        End Try

        If debug_state = True Then
            objExcel.Visible = True
        End If

        If demo_mode = True Then
            objExcel.Visible = True
        End If


        'Debug.WriteLine("Reading " & scenario & ":" & field & " data from worksheet " & currentWorkSheet.Name)

        If Not IsNothing(currentWorkSheet) Then
            'Debug.WriteLine(Left(currentWorkSheet.name, 3) & " " & worksheetName)
            If demo_mode Then
                currentWorkSheet.Activate
                currentWorkSheet.Cells(4, 2).Activate
                System.Threading.Thread.CurrentThread.Sleep(500)
            End If

            If Left(currentWorkSheet.name, 3) = worksheetName Then
                r = currentWorkSheet.Cells.Find(What:=orient_cell)
                If Not IsNothing(r) Then
                    If demo_mode Then
                        currentWorkSheet.Cells(r.row, r.column).Activate
                        System.Threading.Thread.CurrentThread.Sleep(500)
                    End If
                    'Debug.WriteLine("Column: " & r.column)
                    'Debug.WriteLine("Row: " & r.row)
                    'startRow = startRow
                    'startCol = startCol
                    startRow = r.row + 1
                    startCol = r.column
                    If demo_mode Then
                        currentWorkSheet.Cells(startRow, startCol).Activate
                        System.Threading.Thread.CurrentThread.Sleep(500)
                    End If
                    curRow = startRow
                    curCol = startCol
                    'Debug.WriteLine( "Start RC " & startRow &","& startCol
                    'Debug.WriteLine( "Current RC " & curRow &", "& curCol

                    Do Until currentWorkSheet.Cells(curRow, curCol).Value = ""
                        If currentWorkSheet.Cells(curRow, curCol).Value = "Ex: Elizabeth" _
                            Or currentWorkSheet.Cells(curRow, curCol).Value = "EXAMPLE: Peter" _
                            Or currentWorkSheet.Cells(curRow, curCol).Value = "EXAMPLE: Smith" _
                            Or currentWorkSheet.Cells(curRow, curCol).Value = "N/A" _
                            Or currentWorkSheet.Cells(curRow, curCol).Value = "n/a" _
                            Or currentWorkSheet.Cells(curRow, curCol).Value = "First Name(s)" Then
                            curCol = curCol + 1
                            If demo_mode Then
                                currentWorkSheet.Cells(curRow, curCol).Activate
                                System.Threading.Thread.CurrentThread.Sleep(500)
                            End If
                        Else
                            first_name = Trim(currentWorkSheet.Cells(curRow, curCol).Value)
                            If demo_mode Then
                                currentWorkSheet.Cells(curRow, curCol).Activate
                                System.Threading.Thread.CurrentThread.Sleep(500)
                            End If
                            curRow = curRow + 1
                            'Debug.WriteLine(first_name)

                            last_name = Trim(currentWorkSheet.Cells(curRow, curCol).Value)
                            If demo_mode Then
                                currentWorkSheet.Cells(curRow, curCol).Activate
                                System.Threading.Thread.CurrentThread.Sleep(500)
                            End If
                            curRow = curRow + 1
                            'Debug.WriteLine(last_name)


                            eid = Trim(currentWorkSheet.Cells(curRow, curCol).Value)
                            If Not IsNothing(eid) Then
                                eid = eid.ToString()
                                eid = eid.replace("-", "")
                            End If
                            If demo_mode Then
                                currentWorkSheet.Cells(curRow, curCol).Activate
                                System.Threading.Thread.CurrentThread.Sleep(500)
                            End If
                            curRow = curRow + 1
                            'Debug.WriteLine(eid)

                            'Org Team
                            org_team = Trim(currentWorkSheet.Cells(curRow, curCol).Value)
                            'Debug.WriteLine(org_team)
                            If demo_mode Then
                                currentWorkSheet.Cells(curRow, curCol).Activate
                                System.Threading.Thread.CurrentThread.Sleep(500)
                            End If
                            sSql = "SELECT org_team_mapID from org_team_map where org_team = """ & org_team & """"
                            'Debug.WriteLine(sSql)
                            rec.Open(sSql, conn)

                            If (rec.BOF And rec.EOF) Then
                                sSql = "INSERT INTO org_team_map (org_team) VALUES (""" & org_team & """)"
                                'Debug.WriteLine(sSql)
                                conn.Execute(sSql)
                            End If

                            rec.Close()

                            sSql = "SELECT org_team_mapID from org_team_map where org_team = """ & org_team & """"
                            'Debug.WriteLine(sSql)
                            rec.Open(sSql, conn)

                            If (rec.BOF And rec.EOF) Then
                                org_team_mapID = 0
                            Else
                                For Each x In rec.Fields
                                    If Not IsDBNull(x.value) Then
                                        org_team_mapID = x.value
                                    End If
                                Next
                            End If

                            rec.Close()
                            curRow = curRow + 1

                            budget_no = currentWorkSheet.Cells(curRow, curCol).Value
                            If demo_mode Then
                                currentWorkSheet.Cells(curRow, curCol).Activate
                                System.Threading.Thread.CurrentThread.Sleep(500)
                            End If
                            curRow = curRow + 1

                            sSql = "SELECT max(responseID) from responses where org_team = """ & org_team & """ AND EID = """ & eid & """" & " AND submittalID = " & submittalID
                            'Debug.WriteLine(sSql)
                            rec.Open(sSql, conn)

                            i = 0
                            responseID = 0
                            If (rec.BOF And rec.EOF) Then
                                'Debug.WriteLine("The line Is empty")
                            Else
                                For Each x In rec.Fields
                                    If Not IsDBNull(x.value) Then

                                        responseID = x.value
                                    End If
                                Next
                            End If

                            rec.Close()

                            'Debug.WriteLine(responseID)

                            If responseID > 0 Then
                                sSql = "UPDATE responses Set " & scenario & "_" & field & " = 'x' WHERE responseID = " & responseID
                            Else
                                sSql = "INSERT INTO responses (first_name, last_name, EID, org_team_mapID, org_team, budget_no, " & scenario & "_" & field & ", date_recorded, submittalID) values (""" &
                                        first_name & """, """ & last_name & """, """ & eid & """, " & org_team_mapID & ", """ & org_team & """, """ & budget_no & """, ""x"",""" & Format(Now, "MM/dd/yyyy") & """, " & submittalID & ")"
                            End If
                            'Debug.WriteLine(sSql)
                            conn.Execute(sSql)
                            curRow = startRow
                            curCol = curCol + 1
                            index = index + 1
                            'Debug.WriteLine( "RC: " & curRow &","& curCol

                        End If
                    Loop
                Else
                    'Debug.WriteLine("Orient Cell  " & orient_cell & " not encountered.")
                    worksheet_orient_error = " " & scenario & ":" & field & ":" & orient_cell & ";"
                End If
            Else
                worksheet_name_error = " " & worksheetName & ":" & Left(currentWorkSheet.Name, 3) & ";"
                'Debug.WriteLine("Sheetname starting With " & worksheetName & " expected, found worksheet starting With " & Left(currentWorkSheet.Name, 2) & ".")
            End If
        End If

        If index = 0 Then
            'Debug.WriteLine("The program found no records For the " & scenario & ":" & field & " field.")
        End If

        'Try
        'currentWorkSheet.close()
        'Catch ex As Exception
        'Debug.WriteLine("Couldn't Close worksheet")
        'End Try

        results(0) = index
        results(1) = worksheet_name_error
        results(2) = worksheet_orient_error

        Return results

        curRow = Nothing
        curCol = Nothing
        currentWorkSheet = Nothing
        first_name = Nothing
        last_name = Nothing
        eid = Nothing
        org_team = Nothing
        budget_no = Nothing
        sSql = Nothing
        rec = Nothing
        responseID = Nothing
        org_team_mapID = Nothing
        index = Nothing
        debug_state = Nothing
        worksheet_name_error = Nothing
        worksheet_orient_error = Nothing
        results = Nothing
        r = Nothing
        i = Nothing

    End Function

    Function TransposeDim(v As Object) As Object
        ' Custom Function to Transpose a 0-based array (v)

        Dim X As Long, Y As Long, Xupper As Long, Yupper As Long
        Dim tempArray As Object

        Xupper = UBound(v, 2)
        Yupper = UBound(v, 1)

        ReDim tempArray(Xupper, Yupper)
        For X = 0 To Xupper
            For Y = 0 To Yupper
                tempArray(X, Y) = v(Y, X)
            Next Y
        Next X

        TransposeDim = tempArray
    End Function

    Function generate_excel_report(objExcel, conn, file_name, where_clause, demo_mode)
        Dim file_path = ""
        Dim folder = "C:\submissions\"
        Dim file_ext = ".xlsx"
        Dim workbook
        Dim worksheet
        Dim record_count = 1960

        file_path = folder & file_name & file_ext

        Debug.WriteLine("Generating " & file_name & file_ext & "...")

        objExcel.Visible = False
        objExcel.DisplayAlerts = 0 ' Don't display any messages about conversion and so forth
        workbook = objExcel.Workbooks.Add
        workbook.Sheets.Add
        workbook.Sheets.Add
        workbook.Sheets.Add
        worksheet = workbook.Worksheets("Sheet4")
        worksheet.Name = "Roles"
        worksheet = workbook.Worksheets("Sheet3")
        worksheet.Name = "Field by Role"
        worksheet = workbook.Worksheets("Sheet2")
        worksheet.Name = "Field by Scenario"
        worksheet = workbook.Worksheets("Sheet1")
        worksheet.Name = "Role Confirmation Tool"
        workbook.SaveAs(FileName:=file_path)

        generate_field_report(objExcel, conn, where_clause, file_path, "Field by Role", record_count, demo_mode, workbook)
        generate_field_report(objExcel, conn, where_clause, file_path, "Field by Scenario", record_count, demo_mode, workbook)
        generate_by_role_report(objExcel, conn, where_clause, file_path, "Roles", record_count, demo_mode, workbook)

        workbook.Close()

        workbook = Nothing
        worksheet = Nothing
        folder = Nothing
        file_ext = Nothing
        file_path = Nothing

    End Function

    Function generate_error_report(objExcel, conn, file_name, where_clause)
        Dim file_path = ""
        Dim folder = "C: \submissions\"
        Dim file_ext = ".xlsx"
        Dim workbook
        Dim worksheet

        file_path = folder & file_name & file_ext

        Debug.WriteLine("Generating " & file_name & file_ext & "...")

        objExcel.Visible = False
        objExcel.DisplayAlerts = 0 ' Don't display any messages about conversion and so forth
        workbook = objExcel.Workbooks.Add
        ' default sheet       'Identifying information not as expected
        workbook.Sheets.Add   'Couldn't identify session
        workbook.Sheets.Add   'Blank EIDs
        workbook.Sheets.Add   'Malformed EIDs
        workbook.Sheets.Add   'Unexpected Tab Name
        workbook.Sheets.Add   'Can't center cursor on Start Row, Start Column

        Dim Report_definition(0 To 5) As String
        Report_definition(0) = "Sheet6,Missing identifying information,Select * from Errors_no_identifying_information"
        Report_definition(1) = "Sheet5,Session Not identified,Select * from Errors_session_not_identified"
        Report_definition(2) = "Sheet4,Blank EIDs,Select * from Errors_EID_blank"
        Report_definition(3) = "Sheet3,Malformed EIDs,Select * from Errors_EID_malformed"
        Report_definition(4) = "Sheet2,Unexpected tab name,Select * from Errors_worksheet_name"
        Report_definition(5) = "Sheet1,Orient Cell Not encountered,Select * from Errors_worksheet_orient"

        For Each report In Report_definition
            Dim foo = Split(report, ",")
            Dim sheet_name = foo(0)
            Dim new_sheet_name = foo(1)
            Dim sSql = foo(2)
            worksheet = workbook.Worksheets(sheet_name)
            worksheet.Name = new_sheet_name
        Next

        workbook.SaveAs(FileName:=file_path)
        workbook.Close()

        For Each report In Report_definition
            Dim foo = Split(report, ",")
            Dim sheet_name = foo(0)
            Dim new_sheet_name = foo(1)
            Dim sSql = foo(2)
            generate_generic_report(objExcel, conn, sSql, file_path, new_sheet_name)
        Next

        workbook = Nothing
        worksheet = Nothing
        folder = Nothing
        file_ext = Nothing
        file_path = Nothing

    End Function

    Function initiate_unit_reports(objExcel, conn, demo_mode)
        Dim sSql
        Dim rec As ADODB.Recordset
        Dim Unit As String
        Dim record_count As Integer
        Dim file_name As String
        Dim where_clause As String
        Dim i As Integer
        Dim j As Integer
        Dim fld
        Dim start_time
        Dim end_time

        rec = New ADODB.Recordset

        sSql = "SELECT * FROM Workday_Role_mapping_summary"
        Debug.WriteLine(sSql)
        rec.Open(sSql, conn)

        Debug.WriteLine("   Generating Unit reports...")

        Try
            MkDir("C:\submissions\Unit Reports\")
        Catch ex As Exception
            'Debug.WriteLine("Folder already Exists")
        End Try

        j = 0
        If (rec.BOF And rec.EOF) Then
            Debug.WriteLine("No records found.")
        Else
            Do While Not rec.EOF
                i = 0
                start_time = Now()
                For Each fld In rec.Fields
                    If i = 0 Then
                        Unit = fld.value
                    Else
                        record_count = CInt(fld.value)
                    End If
                    i = i + 1
                Next fld
                'Debug.WriteLine("Record Count: " & record_count & "; Unit: " & Unit)
                'file_name = "Working_in_Workday_" & Unit
                where_clause = Unit
                generate_unit_report(objExcel, conn, where_clause, record_count, demo_mode)
                end_time = Now()
                Dim elapsed_time = DateDiff("s", start_time, end_time)
                Debug.WriteLine("Processed " & j & ": " & Unit & " in " & elapsed_time & " seconds")
                rec.MoveNext()
                i = 0
                j = j + 1
            Loop
        End If
        rec.Close()

    End Function

    Function generate_unit_report(objExcel, conn, where_clause, folder, demo_mode)
        Dim file_path = ""
        Dim rec As ADODB.Recordset
        Dim file_ext = ".xlsx"
        Dim workbook
        Dim worksheet
        Dim file_name_append As String
        Dim sSql As String
        Dim Condition As String
        Dim i As Integer
        Dim j As Integer
        Dim unit As String
        Dim record_count As Integer
        Dim file_name As String
        Dim debug_mode As Boolean

        file_path = ""
        rec = New ADODB.Recordset
        file_ext = ".xlsx"
        'workbook
        'worksheet
        file_name_append = ""
        sSql = ""
        Condition = ""
        i = 0
        j = 0
        unit = ""
        record_count = 0
        file_name = "Working In Workday"
        debug_mode = True

        If Not IsNothing(where_clause) Then

            Condition = " WHERE unit = """ & where_clause & """ "
            sSql = "SELECT * FROM Workday_Role_mapping_summary" & Condition

            rec.Open(sSql, conn)
            j = 0
            If (rec.BOF And rec.EOF) Then
                Debug.WriteLine("No records found.")
            Else
                Do While Not rec.EOF
                    i = 0
                    For Each fld In rec.Fields
                        If i = 0 Then
                            unit = fld.value
                        Else
                            record_count = CInt(fld.value)
                        End If
                        i = i + 1
                    Next fld
                    i = 0
                    j = j + 1
                    rec.MoveNext()
                Loop
            End If
            rec.Close()
            file_name_append = "_" & unit
        Else
            file_name_append = ""
            Condition = ""
        End If

        file_path = folder & file_name & file_name_append & file_ext
        Debug.WriteLine(file_path)

        If debug_mode = True Then
            objExcel.Visible = True
        End If
        objExcel.DisplayAlerts = 0 ' Don't display any messages about conversion and so forth
        workbook = objExcel.Workbooks.Add
        workbook.Sheets.Add
        workbook.Sheets.Add
        workbook.Sheets.Add
        worksheet = workbook.Worksheets("Sheet4")
        worksheet.Name = "Roles"
        worksheet = workbook.Worksheets("Sheet3")
        worksheet.Name = "Field By Role"
        worksheet = workbook.Worksheets("Sheet2")
        worksheet.Name = "Field By Scenario"
        worksheet = workbook.Worksheets("Sheet1")
        worksheet.Name = "Role Confirmation Tool"
        workbook.SaveAs(FileName:=file_path)

        workbook = generate_by_role_report(objExcel, conn, where_clause, file_path, "Roles", record_count, demo_mode, workbook)
        workbook = generate_field_report(objExcel, conn, where_clause, file_path, "Field by Role", record_count, demo_mode, workbook)
        workbook = generate_field_report(objExcel, conn, where_clause, file_path, "Field by Scenario", record_count, demo_mode, workbook)
        workbook = generate_role_confirmation_tool(objExcel, conn, where_clause, file_path, "Role Confirmation Tool", record_count, demo_mode, workbook)

        If demo_mode = False Then
            workbook.Close()
        End If


        workbook = Nothing
        worksheet = Nothing
        folder = Nothing
        file_ext = Nothing
        file_path = Nothing

    End Function

    Function generate_by_role_report(objExcel, conn, where_clause, file_path, worksheet_name, record_count, demo_mode, workbook)
        Dim sSql As String
        Dim rec As ADODB.Recordset
        Dim worksheet
        Dim condition As String
        Dim index As Integer
        Dim code As String
        Dim title As String
        Dim i As Integer
        Dim j As Integer
        Dim debug_mode As Boolean
        Dim role_count As Integer

        sSql = ""
        rec = New ADODB.Recordset
        'workbook
        'worksheet
        condition = ""
        index = 0
        code = ""
        title = ""
        i = 0
        j = 0
        debug_mode = False
        role_count = 0

        If where_clause = "" Then
            condition = ""
        Else
            condition = " WHERE unit = """ & where_clause & """"
        End If

        sSql = "SELECT * FROM Workday_Role_Mapping_by_role" & condition & " ORDER BY  `employee last name` asc"
        rec.Open(sSql, conn)
        workbook = generate_worksheet(objExcel, rec, file_path, worksheet_name, workbook)
        rec.Close()

        worksheet = workbook.Worksheets(worksheet_name)

        If debug_mode = True Then
            objExcel.Visible = True
            worksheet.Activate
        End If

        If demo_mode = True Then
            objExcel.Visible = True
            worksheet.Activate
        End If


        worksheet.Columns("A:A").ColumnWidth = 38
        worksheet.Columns("B:B").ColumnWidth = 8
        worksheet.Columns("C:C").ColumnWidth = 15
        worksheet.Columns("D:D").ColumnWidth = 15
        worksheet.Columns("E:E").ColumnWidth = 10
        worksheet.Columns("F:F").ColumnWidth = 46
        worksheet.Columns("G:S").ColumnWidth = 4

        worksheet.Rows("1").Insert

        sSql = "SELECT role_code, role_title FROM roles WHERE role_order is not null ORDER BY  `role_order` asc"
        'Debug.WriteLine(sSql)
        rec.Open(sSql, conn)
        If (rec.BOF And rec.EOF) Then
            Debug.WriteLine("No records found.")
        Else
            Do While Not rec.EOF
                i = 0
                For Each fld In rec.Fields
                    If i = 0 Then
                        code = fld.value
                    Else
                        title = fld.value
                    End If
                    i = i + 1
                Next fld
                i = 0
                j = j + 1
                worksheet.cells(1, j + 6).Value = code
                worksheet.cells(2, j + 6).Value = title
                rec.MoveNext()
            Loop
        End If
        rec.Close()
        role_count = j


        Dim max_column_txt = worksheet.Cells(1, 19).Address
        Dim max_cell_txt = worksheet.Cells(record_count + 3, 19).Address
        Dim max_row = record_count + 3

        index = 6
        Do
            If worksheet.Cells(1, index).Value = "I9" Then
                worksheet.Columns(index).Interior.Color = RGB(253, 228, 207)
                worksheet.Range(worksheet.Cells(1, index), worksheet.Cells(2, index)).Interior.Color = RGB(247, 150, 70)
            ElseIf worksheet.Cells(1, index).Value = "ABP" Then
                worksheet.Columns(index).Interior.Color = RGB(218, 231, 246)
                worksheet.Range(worksheet.Cells(1, index), worksheet.Cells(2, index)).Interior.Color = RGB(83, 141, 213)
            ElseIf worksheet.Cells(1, index).Value = "ACP" Then
                worksheet.Columns(index).Interior.Color = RGB(246, 230, 230)
                worksheet.Range(worksheet.Cells(1, index), worksheet.Cells(2, index)).Interior.Color = RGB(218, 150, 148)
            ElseIf worksheet.Cells(1, index).Value = "CP" Then
                worksheet.Columns(index).Interior.Color = RGB(238, 234, 242)
                worksheet.Range(worksheet.Cells(1, index), worksheet.Cells(2, index)).Interior.Color = RGB(128, 100, 162)
            ElseIf worksheet.Cells(1, index).Value = "CAC" Then
                worksheet.Columns(index).Interior.Color = RGB(228, 223, 236)
                worksheet.Range(worksheet.Cells(1, index), worksheet.Cells(2, index)).Interior.Color = RGB(228, 223, 236)
            ElseIf worksheet.Cells(1, index).Value = "HRC" Then
                worksheet.Columns(index).Interior.Color = RGB(228, 228, 228)
                worksheet.Range(worksheet.Cells(1, index), worksheet.Cells(2, index)).Interior.Color = RGB(178, 178, 178)
            ElseIf worksheet.Cells(1, index).Value = "HRP" Then
                worksheet.Columns(index).Interior.Color = RGB(205, 233, 239)
                worksheet.Range(worksheet.Cells(1, index), worksheet.Cells(2, index)).Interior.Color = RGB(49, 134, 155)
            ElseIf worksheet.Cells(1, index).Value = "TC" Then
                worksheet.Columns(index).Interior.Color = RGB(241, 245, 231)
                worksheet.Range(worksheet.Cells(1, index), worksheet.Cells(2, index)).Interior.Color = RGB(196, 215, 155)
            End If
            index = index + 1
        Loop Until index > role_count + 6

        worksheet.Range("A1:S2").Font.Bold = True

        worksheet.Range("A2:S2").Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous

        worksheet.Range("A3:" & max_cell_txt).Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlDot
        worksheet.Range("A3:" & max_cell_txt).Borders(Excel.XlBordersIndex.xlInsideHorizontal).ThemeColor = 1
        worksheet.Range("A3:" & max_cell_txt).Borders(Excel.XlBordersIndex.xlInsideHorizontal).TintAndShade = -0.14996795556505

        With worksheet.Range("A1:" & max_cell_txt).Font
            .Size = 10
        End With

        worksheet.Columns("G:S").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

        worksheet.Range("G2:S2").Orientation = 90

        worksheet.Columns("A:A").WrapText = True
        worksheet.Columns("F:F").WrapText = True

        worksheet.Range("A2:" & max_cell_txt).Autofilter

        worksheet.Autofilter.Sort.SortFields.Add(Key:=worksheet.Range("F3:F" & max_row), SortOn:=Excel.XlSortOn.xlSortOnValues, Order:=Excel.XlSortOrder.xlAscending, DataOption:=Excel.XlSortDataOption.xlSortNormal)
        'worksheet.Sort.SortFields.Add(Key:=worksheet.Range("D3:D" & max_row), SortOn:=Excel.XlSortOn.xlSortOnValues, Order:=Excel.XlSortOrder.xlAscending, DataOption:=Excel.XlSortDataOption.xlSortNormal)

        With worksheet.Autofilter.Sort
            .Header = Excel.XlYesNoGuess.xlYes
            .MatchCase = False
            .Orientation = Excel.Constants.xlTopToBottom
            .SortMethod = Excel.XlSortMethod.xlPinYin
            .Apply
        End With

        worksheet.PageSetup.PrintArea = "$A$1:" & max_cell_txt
        worksheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape
        worksheet.PageSetup.PaperSize = Excel.XlPaperSize.xlPaper11x17
        worksheet.PageSetup.PrintTitleRows = "$1:$2"
        worksheet.PageSetup.PrintTitleColumns = "$A:$F"
        worksheet.PageSetup.CenterHeader = where_clause & Chr(10) & worksheet_name
        worksheet.PageSetup.RightHeader = "&D"

        workbook.Save()

        If demo_mode = True Then
            System.Threading.Thread.CurrentThread.Sleep(500)
        End If

        Return workbook

        sSql = Nothing
        rec = Nothing
        workbook = Nothing
        worksheet = Nothing

    End Function

    Function generate_field_report(objExcel, conn, where_clause, file_path, worksheet_name, record_count, demo_mode, workbook)
        Dim sSql As String
        Dim rec As ADODB.Recordset
        Dim index As Integer
        Dim worksheet
        Dim i As Integer
        Dim j As Integer
        Dim condition As String
        Dim role_code As String
        Dim field_description As String
        Dim data_column_ct As Integer
        Dim column_offset As Integer
        Dim row_offset As Integer
        Dim dataSql As String
        Dim headerSql As String
        Dim debug_mode As Boolean

        sSql = ""
        rec = New ADODB.Recordset
        index = 0
        condition = ""
        role_code = ""
        field_description = ""
        row_offset = 3
        column_offset = 6
        data_column_ct = 0
        dataSql = ""
        headerSql = ""
        debug_mode = False

        If where_clause = "" Then
            condition = ""
        Else
            condition = " WHERE unit = """ & where_clause & """"
        End If

        If worksheet_name = "Field By Role" Then
            dataSql = "SELECT * FROM Workday_Role_Mapping_by_field_in_order_of_role" & condition
            headerSql = "SELECT role_code, field_description  FROM fields WHERE order_field_by_role_asc is not null ORDER BY  `order_field_by_role_asc` asc"
        Else
            dataSql = "SELECT * FROM Workday_Role_Mapping_by_field_in_order_of_scenario" & condition
            headerSql = "SELECT role_code, field_description  FROM fields WHERE order_field_by_scenario_asc is not null ORDER BY  `order_field_by_scenario_asc` asc"
        End If

        rec = New ADODB.Recordset
        'Debug.WriteLine(dataSql)
        rec.Open(dataSql, conn)
        workbook = generate_worksheet(objExcel, rec, file_path, worksheet_name, workbook)
        rec.Close()

        worksheet = workbook.Worksheets(worksheet_name)

        If debug_mode = True Then
            objExcel.Visible = True
            worksheet.Activate
        End If

        If demo_mode = True Then
            objExcel.Visible = True
            worksheet.Activate
        End If

        worksheet.Columns("A:A").ColumnWidth = 38
        worksheet.Columns("B:B").ColumnWidth = 8
        worksheet.Columns("C:C").ColumnWidth = 15
        worksheet.Columns("D:D").ColumnWidth = 15
        worksheet.Columns("E:E").ColumnWidth = 10
        worksheet.Columns("F:F").ColumnWidth = 48
        worksheet.Columns("G:AW").ColumnWidth = 3

        column_offset = 6 ' the number of fields before data

        worksheet.Rows("1").Insert
        worksheet.Rows("1").Insert

        'Debug.WriteLine(headerSql)
        rec.Open(headerSql, conn)
        If (rec.BOF And rec.EOF) Then
            Debug.WriteLine("No records found.")
        Else
            Do While Not rec.EOF
                i = 0
                For Each fld In rec.Fields
                    If i = 0 Then
                        role_code = fld.value
                    Else
                        field_description = fld.value
                    End If
                    i = i + 1
                Next fld
                i = 0
                j = j + 1
                worksheet.cells(1, j + column_offset).Value = role_code
                worksheet.cells(2, j + column_offset).Value = field_description
                rec.MoveNext()
            Loop
        End If
        rec.Close()
        data_column_ct = j

        Dim total_rows = record_count + row_offset
        Dim total_columns = column_offset + data_column_ct
        Dim max_column_txt = worksheet.Cells(1, column_offset + data_column_ct).Address
        Dim max_cell_txt = worksheet.Cells(total_rows, total_columns).Address

        index = column_offset
        Do
            If worksheet.Cells(1, index).Value = "I9" Then
                worksheet.Columns(index).Interior.Color = RGB(253, 228, 207)
                worksheet.Range(worksheet.Cells(1, index), worksheet.Cells(3, index)).Interior.Color = RGB(247, 150, 70)
            ElseIf worksheet.Cells(1, index).Value = "ABP" Then
                worksheet.Columns(index).Interior.Color = RGB(218, 231, 246)
                worksheet.Range(worksheet.Cells(1, index), worksheet.Cells(3, index)).Interior.Color = RGB(83, 141, 213)
            ElseIf worksheet.Cells(1, index).Value = "ACP" Then
                worksheet.Columns(index).Interior.Color = RGB(246, 230, 230)
                worksheet.Range(worksheet.Cells(1, index), worksheet.Cells(3, index)).Interior.Color = RGB(218, 150, 148)
            ElseIf worksheet.Cells(1, index).Value = "CP" Then
                worksheet.Columns(index).Interior.Color = RGB(238, 234, 242)
                worksheet.Range(worksheet.Cells(1, index), worksheet.Cells(3, index)).Interior.Color = RGB(128, 100, 162)
            ElseIf worksheet.Cells(1, index).Value = "CAC" Then
                worksheet.Columns(index).Interior.Color = RGB(228, 223, 236)
                worksheet.Range(worksheet.Cells(1, index), worksheet.Cells(3, index)).Interior.Color = RGB(228, 223, 236)
            ElseIf worksheet.Cells(1, index).Value = "HRC" Then
                worksheet.Columns(index).Interior.Color = RGB(228, 228, 228)
                worksheet.Range(worksheet.Cells(1, index), worksheet.Cells(3, index)).Interior.Color = RGB(178, 178, 178)
            ElseIf worksheet.Cells(1, index).Value = "HRP" Then
                worksheet.Columns(index).Interior.Color = RGB(205, 233, 239)
                worksheet.Range(worksheet.Cells(1, index), worksheet.Cells(3, index)).Interior.Color = RGB(49, 134, 155)
            ElseIf worksheet.Cells(1, index).Value = "TC" Then
                worksheet.Columns(index).Interior.Color = RGB(241, 245, 231)
                worksheet.Range(worksheet.Cells(1, index), worksheet.Cells(3, index)).Interior.Color = RGB(196, 215, 155)
            End If
            index = index + 1
        Loop Until index > column_offset + data_column_ct

        worksheet.Range("A1:AW3").Font.Bold = True
        worksheet.Range("A3:AW3").Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous

        Dim Dataset = worksheet.Range("A4:" & max_cell_txt)
        Dim entire_sheet = worksheet.Range("A1:" & max_cell_txt)

        With worksheet.Range("A1:" & max_cell_txt).Font
            .Size = 10
        End With

        worksheet.Range("A4:" & max_cell_txt).Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlDot
        worksheet.Range("A4:" & max_cell_txt).Borders(Excel.XlBordersIndex.xlInsideHorizontal).ThemeColor = 1
        worksheet.Range("A4:" & max_cell_txt).Borders(Excel.XlBordersIndex.xlInsideHorizontal).TintAndShade = -0.14996795556505

        worksheet.Columns("G:AW").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

        worksheet.Range("G2:AW2").Orientation = 90
        worksheet.Range("A3:" & max_cell_txt).Autofilter
        worksheet.Columns("A:A").WrapText = True
        worksheet.Columns("F:F").WrapText = True


        worksheet.Autofilter.Sort.SortFields.Add(Key:=worksheet.Range("F3:F" & total_rows), SortOn:=Excel.XlSortOn.xlSortOnValues, Order:=Excel.XlSortOrder.xlAscending, DataOption:=Excel.XlSortDataOption.xlSortNormal)
        'worksheet.Sort.SortFields.Add(Key:=worksheet.Range("D3:D" & max_row), SortOn:=Excel.XlSortOn.xlSortOnValues, Order:=Excel.XlSortOrder.xlAscending, DataOption:=Excel.XlSortDataOption.xlSortNormal)

        With worksheet.Autofilter.Sort
            .Header = Excel.XlYesNoGuess.xlYes
            .MatchCase = False
            .Orientation = Excel.Constants.xlTopToBottom
            .SortMethod = Excel.XlSortMethod.xlPinYin
            .Apply
        End With

        worksheet.PageSetup.PrintArea = "$A$1:" & max_cell_txt
        worksheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape
        worksheet.PageSetup.PaperSize = Excel.XlPaperSize.xlPaper11x17
        worksheet.PageSetup.PrintTitleRows = "$1:$3"
        worksheet.PageSetup.PrintTitleColumns = "$A:$G"
        worksheet.PageSetup.CenterHeader = where_clause & Chr(10) & worksheet_name
        worksheet.PageSetup.RightHeader = "&D"

        workbook.Save()

        If demo_mode = True Then
            System.Threading.Thread.CurrentThread.Sleep(500)
        End If

        Return workbook

        sSql = Nothing
        rec = Nothing
        index = Nothing
        workbook = Nothing
        worksheet = Nothing

    End Function

    Function generate_role_confirmation_tool(objExcel, conn, where_clause, file_path, worksheet_name, record_count, demo_mode, workbook)
        Dim sSql As String
        Dim rec As ADODB.Recordset
        Dim worksheet
        Dim condition As String
        Dim debug_mode As Boolean
        Dim role_title As String
        Dim role_description As String
        Dim role_code As String
        Dim field_count As Integer
        Dim i As Integer
        Dim j As Integer

        sSql = ""
        rec = New ADODB.Recordset
        'workbook 
        'worksheet
        condition = ""
        debug_mode = True
        role_title = ""
        role_description = ""
        role_code = ""
        field_count = 0
        i = 0
        j = 0

        If where_clause = "" Then
            condition = ""
        Else
            condition = " WHERE unit = """ & where_clause & """"
        End If

        sSql = "SELECT * FROM Workday_Role_Mapping_By_Role_Transpose" & condition
        rec.Open(sSql, conn)

        field_count = rec.Fields.Count
        workbook = generate_transposed_worksheet(objExcel, rec, file_path, worksheet_name, record_count, workbook)
        rec.Close()

        worksheet = workbook.Worksheets(worksheet_name)

        If debug_mode = True Then
            objExcel.Visible = True
            worksheet.Activate
            worksheet.Cells(1, 1).Activate
        End If

        If demo_mode = True Then
            objExcel.Visible = True
            worksheet.Activate
            worksheet.Cells(1, 1).Select
        End If

        worksheet.Columns(1).Insert
        worksheet.Columns(1).Insert
        worksheet.Columns(1).Insert

        worksheet.Columns("A:A").ColumnWidth = 6        'Workday Code
        worksheet.Columns("B:B").ColumnWidth = 30       'Workday Role
        worksheet.Columns("C:C").ColumnWidth = 75       'Workday Role Description

        'worksheet.Range("G2:N2").Cut

        sSql = "SELECT role_code, role_title, role_description FROM roles WHERE role_order is not null ORDER BY  `role_order` asc"
        Debug.WriteLine(sSql)
        rec.Open(sSql, conn)
        If (rec.BOF And rec.EOF) Then
            Debug.WriteLine("No records found.")
        Else
            Do While Not rec.EOF
                i = 0
                For Each fld In rec.Fields
                    If i = 0 Then
                        role_code = fld.value
                    ElseIf i = 1 Then
                        role_title = fld.value
                    Else
                        role_description = fld.value
                    End If
                    i = i + 1
                Next fld
                i = 0
                j = j + 1
                worksheet.cells(j + 1, 1).Value = role_code
                worksheet.cells(j + 1, 2).Value = role_title
                worksheet.cells(j + 1, 3).Value = role_description
                rec.MoveNext()
            Loop
        End If
        rec.Close()

        Dim max_column_txt = worksheet.Cells(1, record_count + 3).Address
        Dim max_cell_txt = worksheet.Cells(14, record_count + 3).Address

        worksheet.Range("A1:" & max_column_txt).Font.Bold = True

        'worksheet.Range("A2:N2").Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous

        'worksheet.Range("A3:N2000").Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlDot
        'worksheet.Range("A3:N2000").Borders(Excel.XlBordersIndex.xlInsideHorizontal).ThemeColor = 1
        'worksheet.Range("A3:N2000").Borders(Excel.XlBordersIndex.xlInsideHorizontal).TintAndShade = -0.14996795556505

        'worksheet.Columns("D:N").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

        worksheet.Range("D1:" & max_column_txt).Orientation = 90
        worksheet.Range("D1:" & max_column_txt).RowHeight = 100
        worksheet.Range("D1:" & max_column_txt).ColumnWidth = 9
        worksheet.Range("D1:" & max_cell_txt).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
        worksheet.Range("D1:" & max_cell_txt).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        worksheet.Range("D1:" & max_column_txt).WrapText = True

        Dim selection = worksheet.Range("A1:" & max_cell_txt)

        With worksheet.Range("A1:" & max_cell_txt).Font
            .Size = 10
        End With

        With selection.Borders(Excel.XlBordersIndex.xlEdgeLeft)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = Excel.XlBorderWeight.xlThin
        End With
        With selection.Borders(Excel.XlBordersIndex.xlEdgeTop)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = Excel.XlBorderWeight.xlThin
        End With
        With selection.Borders(Excel.XlBordersIndex.xlEdgeBottom)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = Excel.XlBorderWeight.xlThin
        End With
        With selection.Borders(Excel.XlBordersIndex.xlEdgeRight)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = Excel.XlBorderWeight.xlThin
        End With
        With selection.Borders(Excel.XlBordersIndex.xlInsideVertical)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = Excel.XlBorderWeight.xlThin
        End With
        With selection.Borders(Excel.XlBordersIndex.xlInsideHorizontal)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = Excel.XlBorderWeight.xlThin
        End With

        selection = worksheet.Range("A2:" & max_cell_txt)

        With selection.Borders(Excel.XlBordersIndex.xlEdgeLeft)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = Excel.XlBorderWeight.xlThin
        End With
        With selection.Borders(Excel.XlBordersIndex.xlEdgeTop)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = Excel.XlBorderWeight.xlThin
        End With
        With selection.Borders(Excel.XlBordersIndex.xlEdgeBottom)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = Excel.XlBorderWeight.xlThin
        End With
        With selection.Borders(Excel.XlBordersIndex.xlEdgeRight)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = Excel.XlBorderWeight.xlThin
        End With
        With selection.Borders(Excel.XlBordersIndex.xlInsideVertical)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = Excel.XlBorderWeight.xlThin
        End With
        With selection.Borders(Excel.XlBordersIndex.xlInsideHorizontal)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = Excel.XlBorderWeight.xlThin
        End With

        worksheet.Columns("B:C").WrapText = True

        worksheet.Range("A1:" & max_cell_txt).Autofilter

        worksheet.Range("A1:" & max_column_txt).Interior.Color = RGB(216, 191, 235)

        worksheet.PageSetup.PrintArea = "$A$1:" & max_cell_txt
        worksheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape
        worksheet.PageSetup.PaperSize = Excel.XlPaperSize.xlPaper11x17
        worksheet.PageSetup.PrintTitleRows = "$1:$1"
        worksheet.PageSetup.PrintTitleColumns = "$A:$C"
        worksheet.PageSetup.CenterHeader = where_clause & Chr(10) & worksheet_name
        worksheet.PageSetup.RightHeader = "&D"

        workbook.Save()

        If demo_mode = True Then
            System.Threading.Thread.CurrentThread.Sleep(500)
        End If

        Return workbook

        sSql = Nothing
        rec = Nothing
        workbook = Nothing
        worksheet = Nothing

    End Function

    Function generate_worksheet(objExcel, recordset, file_path, worksheet_name, workbook)
        Dim Worksheet
        Dim fieldCount
        Dim recArray
        Dim recCount
        Dim debug_mode

        debug_mode = False

        If debug_mode = True Then
            objExcel.Visible = True
        End If


        objExcel.DisplayAlerts = 0 ' Don't display any messages about conversion and so forth
        Worksheet = workbook.Worksheets(worksheet_name)

        ' Copy field names to the first row of the worksheet
        fieldCount = recordset.Fields.Count
        For iCol = 1 To fieldCount
            Worksheet.Cells(1, iCol).Value = recordset.Fields(iCol - 1).Name
        Next

        ' Check version of Excel
        If Val(Mid(objExcel.Version, 1, InStr(1, objExcel.Version, ".") - 1)) > 8 Then
            'EXCEL 2000,2002,2003, or 2007: Use CopyFromRecordset

            ' Copy the recordset to the worksheet, starting in cell A2
            Worksheet.Cells(2, 1).CopyFromRecordset(recordset)
            'Note: CopyFromRecordset will fail if the recordset
            'contains an OLE object field or array data such
            'as hierarchical recordsets

        Else
            'EXCEL 97 or earlier: Use GetRows then copy array to Excel

            ' Copy recordset to an array
            recArray = recordset.GetRows
            'Note: GetRows returns a 0-based array where the first
            'dimension contains fields and the second dimension
            'contains records. We will transpose this array so that
            'the first dimension contains records, allowing the
            'data to appears properly when copied to Excel

            ' Determine number of records

            recCount = UBound(recArray, 2) + 1 '+ 1 since 0-based array


            ' Check the array for contents that are not valid when
            ' copying the array to an Excel worksheet
            For iCol = 0 To fieldCount - 1
                For iRow = 0 To recCount - 1
                    ' Take care of Date fields
                    If IsDate(recArray(iCol, iRow)) Then
                        recArray(iCol, iRow) = Format(recArray(iCol, iRow))
                        ' Take care of OLE object fields or array fields
                    ElseIf IsArray(recArray(iCol, iRow)) Then
                        recArray(iCol, iRow) = "Array Field"
                    End If
                Next iRow 'next record
            Next iCol 'next field

            ' Transpose and Copy the array to the worksheet,
            ' starting in cell A2
            Worksheet.Cells(2, 1).Resize(recCount, fieldCount).Value =
                TransposeDim(recArray)
        End If

        ' Auto-fit the column widths and row heights
        objExcel.Selection.CurrentRegion.Columns.AutoFit
        objExcel.Selection.CurrentRegion.Rows.AutoFit

        workbook.SaveAs(FileName:=file_path)

        Return workbook

        workbook = Nothing
        Worksheet = Nothing

    End Function

    Function generate_transposed_worksheet(objExcel, recordset, file_path, worksheet_name, record_count, workbook)
        Dim Worksheet
        Dim fieldCount
        Dim recArray
        Dim recCount
        Dim debug_mode

        debug_mode = False

        If debug_mode = True Then
            objExcel.Visible = True
        End If
        objExcel.DisplayAlerts = 0 ' Don't display any messages about conversion and so forth
        Worksheet = workbook.Worksheets(worksheet_name)
        Worksheet.Select

        ' Copy field names to the first row of the worksheet
        fieldCount = recordset.Fields.Count

        For iCol = 1 To fieldCount
            Worksheet.Cells(1, iCol).Value = recordset.Fields(iCol - 1).Name
        Next

        ' Check version of Excel
        If Val(Mid(objExcel.Version, 1, InStr(1, objExcel.Version, ".") - 1)) > 8 Then
            'EXCEL 2000,2002,2003, or 2007: Use CopyFromRecordset

            ' Copy the recordset to the worksheet, starting in cell A2
            Worksheet.Range("A1").CopyFromRecordset(recordset)
            Dim max_cell_txt = Worksheet.cells(record_count, fieldCount).Address
            Dim min_cell_txt = Worksheet.cells(1, 1).Address
            Dim range = min_cell_txt & ":" & max_cell_txt
            Dim t_range = Worksheet.cells(record_count + 1, 1).Address

            Worksheet.Range(range).Copy
            Debug.WriteLine(t_range)
            Worksheet.Range(t_range).Select
            Worksheet.Range(t_range).PasteSpecial(Paste:=Excel.XlPasteType.xlPasteValues,
                    Transpose:=True)
            'Note: CopyFromRecordset will fail if the recordset
            'contains an OLE object field or array data such
            'as hierarchical recordsets

        Else
            'EXCEL 97 or earlier: Use GetRows then copy array to Excel

            ' Copy recordset to an array
            recArray = recordset.GetRows
            'Note: GetRows returns a 0-based array where the first
            'dimension contains fields and the second dimension
            'contains records. We will transpose this array so that
            'the first dimension contains records, allowing the
            'data to appears properly when copied to Excel

            ' Determine number of records

            recCount = UBound(recArray, 2) + 1 '+ 1 since 0-based array

            ' Check the array for contents that are not valid when
            ' copying the array to an Excel worksheet
            For iCol = 0 To fieldCount - 1
                For iRow = 0 To recCount - 1
                    ' Take care of Date fields
                    If IsDate(recArray(iCol, iRow)) Then
                        recArray(iCol, iRow) = Format(recArray(iCol, iRow))
                        ' Take care of OLE object fields or array fields
                    ElseIf IsArray(recArray(iCol, iRow)) Then
                        recArray(iCol, iRow) = "Array Field"
                    End If
                Next iRow 'next record
            Next iCol 'next field

            ' Transpose and Copy the array to the worksheet,
            ' starting in cell A2
            Worksheet.Cells(2, 1).Resize(recCount, fieldCount).Value =
                TransposeDim(recArray)
        End If

        Worksheet.Rows("1:" & record_count.ToString).delete

        Worksheet.Rows("2:2").delete

        ' Auto-fit the column widths and row heights
        objExcel.Selection.CurrentRegion.Columns.AutoFit
        objExcel.Selection.CurrentRegion.Rows.AutoFit

        workbook.SaveAs(FileName:=file_path)

        Return workbook

        workbook = Nothing
        Worksheet = Nothing

    End Function

    Function generate_generic_report(objExcel, conn, sSql, file_path, worksheet_name)
        Dim rec As ADODB.Recordset
        Dim index As Integer
        Dim workbook
        Dim worksheet
        Dim MaxCol = 0
        Dim MaxRow = 3000

        rec = New ADODB.Recordset
        rec.Open(sSql, conn)
        generate_worksheet(objExcel, rec, file_path, worksheet_name, workbook)

        Debug.WriteLine("FindRecordCount :" & rec.Fields.Count)
        rec.Close()

        MaxCol = rec.Fields.Count
        Dim MaxColTxt = MaxCol.ToString
        Dim MaxRowTxt = MaxRow.ToString

        workbook = objExcel.Workbooks.Open(file_path)
        objExcel.Visible = True

        worksheet = workbook.Worksheets(worksheet_name)

        Dim MaxCell = worksheet.Cells(MaxRow, MaxCol)
        Dim lastColumnCell = worksheet.Cells(1, MaxCol)
        Dim StartCell = worksheet.Cells(1, 1)
        Dim lastRowCell = worksheet.Cells(MaxRow, 1)
        Dim firstDataCell = worksheet.Cells(2, 1)
        Dim Full_set = worksheet.Range(StartCell, MaxCell)
        Dim Dataset = worksheet.Range(firstDataCell, MaxCell)

        worksheet.Range("$1:$1").Font.Bold = True

        worksheet.Range(StartCell, lastColumnCell).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous

        Dataset.Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlDot
        Dataset.Borders(Excel.XlBordersIndex.xlInsideHorizontal).ThemeColor = 1
        Dataset.Borders(Excel.XlBordersIndex.xlInsideHorizontal).TintAndShade = -0.14996795556505

        'worksheet.Columns("G:AW").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

        Full_set.Autofilter

        'worksheet.PageSetup.PrintArea = Full_set
        worksheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape
        worksheet.PageSetup.PaperSize = Excel.XlPaperSize.xlPaper11x17
        worksheet.PageSetup.PrintTitleRows = "$1:$1"
        'worksheet.PageSetup.PrintTitleColumns = "$A:$G"
        worksheet.PageSetup.CenterHeader = worksheet_name
        worksheet.PageSetup.RightHeader = "&D"

        workbook.Save()
        workbook.Close

        sSql = Nothing
        rec = Nothing
        index = Nothing
        workbook = Nothing
        worksheet = Nothing

    End Function


End Module
