Option Explicit

' === Config ===
Const TARGET_WORKSHEETS As String = "AKH, WSK, MAG"
Const PREFIX_LOOKUP As String = "LT-"
Const HEADER_ROW As Long = 1
Const TARGET_COLUMNS_FW As String = "#, interne/externe Verbindung, Kategorie, Beschreibung, Kostenstelle, Quelle, Ziel, Serviceprotokoll, Ports"
Const TARGET_COLUMNS_SFW As String = "#, Servername, Betriebsystem, Art des Geschäftsfalls, Protokoll, Ports, Quelle, Kostenstelle"

Const SAVE_SUBPATH As String = "\Export"

Const COLUMN_NO As Long = 1                          ' #
Const COLUMN_INDICATOR As Long = 2                   ' Beantragt
Const COLUMN_DATE As Long = 3                        ' Beantragungsdatum
Const COLUMN_STATUS As Long = 4                      ' Status
Const COLUMN_ENV As Long = 5                         ' Ebene
Const COLUMN_INTEXT As Long = 6                      ' interne/externe Verbindung
Const COLUMN_CATEGORY As Long = 7                    ' Kategorie
Const COLUMN_DESC As Long = 8                        ' Beschreibung
Const COLUMN_IP As Long = 11                         ' Gegenstelle
Const COLUMN_DIRECTION As Long = 12                  ' Richtung
Const COLUMN_PROTOCOL As Long = 15                   ' Protokoll
Const COLUMN_PORTS As Long = 16                      ' Ports
Const COLUMN_COSTCENTER As Long = 19                 ' Kostenstelle
Const COLUMN_CONNECTIONSTATUS As Long = 21           ' Ergebnis Verbindung
Const COLUMN_CONNECTIONSTATUSDATE As Long = 22       ' Zeitpunkt Test


Const LTCOLUMN_ENV As Long = 1                       ' Ebene
Const LTCOLUMN_HOST As Long = 2                      ' Hostname
Const LTCOLUMN_ALIAS As Long = 3                     ' Alias
Const LTCOLUMN_IP As Long = 4                        ' IP
Const LTCOLUMN_VIP As Long = 5                       ' Virtuelle IP
Const LTCOLUMN_OS As Long = 6                        ' Betriebsystem
Const LTCOLUMN_DIRECTION As Long = 7                 ' Richtung

' === Main ===
' Summary: Creates export workbooks and marks rows as processed.
Sub Process()
    Dim targetWorksheets() As String
    Dim targetColumns() As String
    Dim overviewWorkbook As Workbook
    Dim currentDateTime As Date
    Dim sourceSheet As Worksheet
    Dim lookupSheet As Worksheet
    Dim unprocessedRows As Collection
    Dim lastRow As Long
    Dim dataArea As Range
    Dim cell As Range
    Dim wbNew As Workbook
    Dim wsNew As Worksheet
    Dim targetColIndex As Long
    Dim newFileName As String
    Dim savePath As String
    Dim dataRow As Range
    Dim targetIndex As Long
    Dim targetRow As Long
    Dim direction As String
    Dim dict As Variant
    Dim env As Variant
    Dim ips As Variant
    Dim ip As String
    Dim i As Long
    Dim kind As Variant
    
    ' === Initialization ===
    Set overviewWorkbook = ActiveWorkbook
    currentDateTime = Now()
    targetWorksheets = SplitAndTrim(TARGET_WORKSHEETS)
    savePath = overviewWorkbook.Path & SAVE_SUBPATH & Application.PathSeparator
    EnsureFolderExists savePath
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    
    ' === Main Processing Loop, iterates over AKH/WSK/MAG ===
    For Each sourceSheet In overviewWorkbook.Worksheets
        If Not IsInArray(sourceSheet.Name, targetWorksheets) Then GoTo NextIteration
        
        Set lookupSheet = overviewWorkbook.Worksheets(PREFIX_LOOKUP & sourceSheet.Name)
        
        ' --- For each type of target excel file
        For Each kind In Array("SFW", "FW")
        
            ' --- Find Unprocessed Rows, group by environment ---
            Set dict = CreateObject("Scripting.Dictionary")
            lastRow = 0 ' Reset lastRow
            On Error Resume Next
            lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, COLUMN_INDICATOR).End(xlUp).Row
            On Error GoTo 0
    
            If lastRow > HEADER_ROW Then
                Set dataArea = sourceSheet.Range(sourceSheet.Cells(HEADER_ROW + 1, COLUMN_INDICATOR), sourceSheet.Cells(lastRow, COLUMN_INDICATOR))
                For Each cell In dataArea
                    If Trim(CStr(cell.Value)) = "Nein" And (kind = "FW" Or sourceSheet.Cells(cell.Row, COLUMN_DIRECTION).Value = "IN") Then ' skip OUT for server firewall
                    
                        env = sourceSheet.Cells(cell.Row, COLUMN_ENV).Value
                        If dict.Exists(env) Then
                            Set unprocessedRows = dict(env)
                        Else
                            Set unprocessedRows = New Collection
                            dict.Add env, unprocessedRows
                        End If
                        
                        unprocessedRows.Add sourceSheet.Rows(cell.Row)
                    End If
                Next cell
            End If
            
            ' --- Iterate over environments (DEV, MIG, PROD)
            For Each env In dict.Keys
                Set unprocessedRows = dict(env)
                
                ' --- Init new workbook
                targetColumns = SplitAndTrim(IIf(kind = "FW", TARGET_COLUMNS_FW, TARGET_COLUMNS_SFW))
                
                ' --- New workbook
                Set wbNew = Workbooks.Add(xlWBATWorksheet)
                Set wsNew = wbNew.Worksheets(1)
                wsNew.Name = sourceSheet.Name
                
                ' --- Write headers
                For targetColIndex = 0 To UBound(targetColumns)
                    wsNew.Cells(HEADER_ROW, targetColIndex + 1).Value = Trim(targetColumns(targetColIndex))
                Next targetColIndex
                wsNew.Rows(HEADER_ROW).Font.Bold = True
                        
                ' --- Write content
                targetRow = HEADER_ROW + 1
                For Each dataRow In unprocessedRows ' dataRow is a Range object representing the source row
                
                    direction = sourceSheet.Cells(dataRow.Row, COLUMN_DIRECTION).Value
                    ips = FindIPs(direction, env, lookupSheet)
                    
                    ' --- Iterate over IPs from Lookup sheet
                    For i = LBound(ips) To UBound(ips)
                        ' --- Copy values
                        CopyValue sourceSheet, wsNew, targetColumns, targetRow, dataRow, COLUMN_NO, "#", "000"                          ' FW, SFW
                        CopyValue sourceSheet, wsNew, targetColumns, targetRow, dataRow, COLUMN_INTEXT, "interne/externe Verbindung"    ' FW
                        CopyValue sourceSheet, wsNew, targetColumns, targetRow, dataRow, COLUMN_CATEGORY, "Kategorie"                   ' FW
                        CopyValue sourceSheet, wsNew, targetColumns, targetRow, dataRow, COLUMN_DESC, "Beschreibung"                    ' FW
                        CopyValue sourceSheet, wsNew, targetColumns, targetRow, dataRow, COLUMN_COSTCENTER, "Kostenstelle"              ' FW, SFW
                        CopyValue sourceSheet, wsNew, targetColumns, targetRow, dataRow, COLUMN_PROTOCOL, "Serviceprotokoll"            ' FW
                        CopyValue sourceSheet, wsNew, targetColumns, targetRow, dataRow, COLUMN_PROTOCOL, "Protokoll"                   ' SFW
                        CopyValue sourceSheet, wsNew, targetColumns, targetRow, dataRow, COLUMN_PORTS, "Ports"                          ' FW, SFW
                        WriteValue wsNew, targetColumns, targetRow, "Art des Geschäftsfalls", "Anforderung"                             ' SFW
                            
                        ip = CStr(ips(i)(0)) & ";" & CStr(ips(i)(1))
                        If direction = "OUT" Then
                            WriteValue wsNew, targetColumns, targetRow, "Quelle", ip                                                    ' FW
                            CopyValue sourceSheet, wsNew, targetColumns, targetRow, dataRow, COLUMN_IP, "Ziel"                          ' FW
                        Else
                            CopyValue sourceSheet, wsNew, targetColumns, targetRow, dataRow, COLUMN_IP, "Quelle"                        ' FW
                            WriteValue wsNew, targetColumns, targetRow, "Ziel", ip                                                      ' FW
                            WriteValue wsNew, targetColumns, targetRow, "Servername", ip                                                ' SFW
                            WriteValue wsNew, targetColumns, targetRow, "Betriebsystem", CStr(ips(i)(2))                                ' SFW
                        End If
                        
                        ' (next)
                        targetRow = targetRow + 1
                    Next i
                    
                    
                Next dataRow
                        
                ' --- Autofit Target Columns ---
                wsNew.Columns(1).Resize(, UBound(targetColumns) + 1).AutoFit
        
                ' --- Save New Workbook ---
                newFileName = savePath & IIf(kind = "FW", "Firewall", "Server Firewall") & " Freischaltung " & sourceSheet.Name & " " & env & " " & Format(currentDateTime, "yyyy-mm-dd hh-mm-ss") & ".xlsx"
                
                On Error Resume Next
                wbNew.SaveAs Filename:=newFileName, FileFormat:=xlOpenXMLWorkbook
                If Err.Number <> 0 Then
                    MsgBox "Could not save the file: " & newFileName & vbCrLf & _
                           "Error: " & Err.Description & vbCrLf & _
                           "The original data on sheet '" & sourceSheet.Name & "' has NOT been marked as processed.", vbExclamation, "Save Failed"
                    wbNew.Close SaveChanges:=False
                    Err.Clear
                Else
                    On Error GoTo 0
                    wbNew.Close SaveChanges:=False
                    
                    ' --- Update Original Sheet Indicator Column. only after FW ---
                    If kind = "FW" Then
                        For Each dataRow In unprocessedRows
                            sourceSheet.Cells(dataRow.Row, COLUMN_INDICATOR).Value = "Ja"
                            sourceSheet.Cells(dataRow.Row, COLUMN_DATE).Value = currentDateTime ' Use date & time
                            sourceSheet.Cells(dataRow.Row, COLUMN_DATE).numberFormat = "yyyy-mm-dd hh:mm:ss"
                            sourceSheet.Cells(dataRow.Row, COLUMN_STATUS).Value = "offen"
                            
                        Next dataRow
                    End If
                End If
                
                ' Clean up object variables for this sheet
                Set unprocessedRows = Nothing
                Set wsNew = Nothing
                Set wbNew = Nothing
                Set dataArea = Nothing
            Next env
        Next kind
NextIteration:
    Next sourceSheet
    
End Sub

' Summary: Looks up host/IP/OS entries for the given environment and direction.
Function FindIPs(direction As String, env As Variant, lookupSheet As Worksheet) As Variant
    Dim lastRow    As Long
    Dim i          As Long
    Dim results()  As Variant
    Dim matchCount As Long
    
    lastRow = lookupSheet.Cells(lookupSheet.Rows.Count, LTCOLUMN_ENV).End(xlUp).Row
                      
    matchCount = 0
    
    For i = 2 To lastRow
       If StrComp(lookupSheet.Cells(i, LTCOLUMN_ENV).Value, env, vbTextCompare) = 0 And _
            (StrComp(lookupSheet.Cells(i, LTCOLUMN_DIRECTION).Value, direction, vbTextCompare) = 0 Or StrComp(lookupSheet.Cells(i, LTCOLUMN_DIRECTION).Value, "IN/OUT", vbTextCompare) = 0) Then
           
            Dim ipAddr As String
            Dim osName As String
            Dim host   As String
            
            host = CStr(lookupSheet.Cells(i, LTCOLUMN_HOST).Value)
            osName = CStr(lookupSheet.Cells(i, LTCOLUMN_OS).Value)
            
            ' pick the correct IP column
            If CStr(lookupSheet.Cells(i, LTCOLUMN_VIP).Value) <> "" Then
                ipAddr = CStr(lookupSheet.Cells(i, LTCOLUMN_VIP).Value)
            Else
                ipAddr = CStr(lookupSheet.Cells(i, LTCOLUMN_IP).Value)
            End If
            
            If Len(Trim(ipAddr)) > 0 Then
                ReDim Preserve results(0 To matchCount)
                ' each element is a 3-item array: Hostname, IP, OS
                results(matchCount) = Array(host, ipAddr, osName)
                matchCount = matchCount + 1
            End If
        End If
    Next i
    
    ' if no matches, return an empty array
    If matchCount = 0 Then
        FindIPs = Array()
    Else
        FindIPs = results
    End If
End Function

' Summary: Copies a cell value to the target sheet when the destination column exists.
Sub CopyValue(sourceSheet As Worksheet, targetSheet As Worksheet, targetColumns() As String, targetRow As Long, dataRow As Range, sourceCol As Long, targetCol As String, Optional numberFormat As String = "")
    Dim targetIndex As Long
    
    targetIndex = IndexOfInArray(targetColumns, targetCol)
    If targetIndex <> -1 Then
        targetSheet.Cells(targetRow, targetIndex + 1).Value = sourceSheet.Cells(dataRow.Row, sourceCol).Value
    
        If (numberFormat <> "") Then
            targetSheet.Cells(targetRow, targetIndex + 1).numberFormat = numberFormat
        End If
    End If
    
End Sub

' Summary: Writes a fixed value to the target sheet if the column exists.
Sub WriteValue(targetSheet As Worksheet, targetColumns() As String, targetRow As Long, targetCol As String, fixedValue As Variant, Optional numberFormat As String = "")
    Dim targetIndex As Long
    
    targetIndex = IndexOfInArray(targetColumns, targetCol)
    If targetIndex <> -1 Then
        targetSheet.Cells(targetRow, targetIndex + 1).Value = fixedValue
    
        If (numberFormat <> "") Then
            targetSheet.Cells(targetRow, targetIndex + 1).numberFormat = numberFormat
        End If
    End If
    
End Sub

' === Helpers ===
' Summary: Returns True if val exists in arr.
Function IsInArray(val As String, arr As Variant) As Boolean
    Dim item As Variant
    For Each item In arr
        If item = val Then
            IsInArray = True
            Exit Function
        End If
    Next
    IsInArray = False
End Function

' Summary: Splits a comma-separated string, trims each value and removes empties.
' Returns a zero-based array of strings.
Public Function SplitAndTrim(inputStr As String) As Variant
    Dim parts As Variant
    Dim col As Collection
    Dim item As Variant
    Dim t As String
    Dim result() As String
    Dim i As Long
    
    ' Split the string by comma
    parts = Split(inputStr, ",")
    
    ' Use a Collection to gather non-empty, trimmed items
    Set col = New Collection
    For Each item In parts
        t = Trim$(CStr(item))
        If Len(t) > 0 Then
            col.Add t
        End If
    Next item
    
    ' If nothing matched, return an empty array
    If col.Count = 0 Then
        SplitAndTrim = Array()  ' zero-length array
        Exit Function
    End If
    
    ' Build a zero-based array from the collection
    ReDim result(0 To col.Count - 1)
    For i = 1 To col.Count
        result(i - 1) = col(i)
    Next i
    
    SplitAndTrim = result
End Function

' Summary: Returns the zero-based index of searchText within the string array arr.
' If not found, returns -1.
'
' arr          A 1-D array of strings (can be 0-based or 1-based).
' searchText   The string to locate.
' ignoreCase   Optional. If True, does case-insensitive match (default = False).
Public Function IndexOfInArray(arr As Variant, _
                               searchText As String, _
                               Optional ignoreCase As Boolean = False) As Long
    Dim i As Long
    Dim lowerBound As Long, upperBound As Long
    
    ' Determine array bounds
    lowerBound = LBound(arr)
    upperBound = UBound(arr)
    
    ' Loop through each element
    For i = lowerBound To upperBound
        If ignoreCase Then
            If StrComp(CStr(arr(i)), searchText, vbTextCompare) = 0 Then
                IndexOfInArray = i - lowerBound   ' normalize to zero-based
                Exit Function
            End If
        Else
            If CStr(arr(i)) = searchText Then
                IndexOfInArray = i - lowerBound   ' normalize to zero-based
                Exit Function
            End If
        End If
    Next i
    
    ' Not found
    IndexOfInArray = -1
End Function

' Summary: Ensures that a folder exists, creating it if necessary.
Function EnsureFolderExists(folderPath As String) As Boolean
    Dim fso As Object
    On Error GoTo ErrHandler
    
    ' Create FileSystemObject (late binding)
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' If the folder doesn't exist, create it (and any necessary parent folders)
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
    
    EnsureFolderExists = True
    Exit Function

ErrHandler:
    ' Something went wrong (invalid path, permissions, etc.)
    EnsureFolderExists = False
End Function

' === Import connection check results ===
' Summary: Reads connection result CSV files and updates connection status columns.
Sub ImportConnectionResults()
    Dim overviewWorkbook As Workbook
    Dim importPath As String
    Dim targetWorksheets() As String
    Dim csvCache As Object
    Dim sourceSheet As Worksheet
    Dim lookupSheet As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim env As String
    Dim direction As String
    Dim hosts As Variant
    Dim hostItem As Variant
    Dim ipValue As String
    Dim portValue As String
    Dim ipParts As Variant
    Dim portParts As Variant
    Dim host As String
    Dim csvData As Object
    Dim ipPart As Variant
    Dim portPart As Variant
    Dim result As Variant
    Dim total As Long
    Dim successCount As Long
    Dim latestTs As String
    Dim statusText As String

    Set overviewWorkbook = ActiveWorkbook
    importPath = overviewWorkbook.Path & "\Import\"
    targetWorksheets = SplitAndTrim(TARGET_WORKSHEETS)
    Set csvCache = CreateObject("Scripting.Dictionary")

    For Each sourceSheet In overviewWorkbook.Worksheets
        If Not IsInArray(sourceSheet.Name, targetWorksheets) Then GoTo NextSheet

        Set lookupSheet = overviewWorkbook.Worksheets(PREFIX_LOOKUP & sourceSheet.Name)
        lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, COLUMN_ENV).End(xlUp).Row

        For rowIndex = HEADER_ROW + 1 To lastRow
            direction = CStr(sourceSheet.Cells(rowIndex, COLUMN_DIRECTION).Value)
            If StrComp(direction, "OUT", vbTextCompare) <> 0 Then GoTo NextRow
            If StrComp(CStr(sourceSheet.Cells(rowIndex, COLUMN_CONNECTIONSTATUS).Value), "OK", vbTextCompare) = 0 Then GoTo NextRow

            env = CStr(sourceSheet.Cells(rowIndex, COLUMN_ENV).Value)
            hosts = FindIPs("OUT", env, lookupSheet)
            On Error Resume Next
            If UBound(hosts) < LBound(hosts) Then
                On Error GoTo 0
                GoTo NextRow
            End If
            On Error GoTo 0

            ipValue = CStr(sourceSheet.Cells(rowIndex, COLUMN_IP).Value)
            portValue = CStr(sourceSheet.Cells(rowIndex, COLUMN_PORTS).Value)

            ipParts = SplitAndTrim(Replace(ipValue, ";", ","))
            portParts = SplitAndTrim(Replace(portValue, ";", ","))

            total = 0
            successCount = 0
            latestTs = ""

            For Each hostItem In hosts
                host = CStr(hostItem(0))
                If Not csvCache.Exists(host) Then
                    csvCache.Add host, LoadCsvFile(importPath & host & ".csv")
                End If
                Set csvData = csvCache(host)

                For Each ipPart In ipParts
                    For Each portPart In portParts
                        total = total + 1
                        result = CheckCsvForHostPort(csvData, ipPart, portPart)
                        If result(0) Then
                            successCount = successCount + 1
                        End If
                        If latestTs = "" Or result(1) > latestTs Then latestTs = result(1)
                    Next portPart
                Next ipPart
            Next hostItem

            If total > 0 And latestTs <> "" Then
                If successCount = total Then
                    statusText = "OK"
                ElseIf successCount > 0 Then
                    statusText = "partially"
                Else
                    statusText = "NOK"
                End If

                sourceSheet.Cells(rowIndex, COLUMN_CONNECTIONSTATUS).Value = statusText
                sourceSheet.Cells(rowIndex, COLUMN_CONNECTIONSTATUSDATE).Value = latestTs
                sourceSheet.Cells(rowIndex, COLUMN_CONNECTIONSTATUSDATE).NumberFormat = "yyyy-mm-dd hh:mm:ss"
            End If

NextRow:
        Next rowIndex

NextSheet:
    Next sourceSheet
End Sub

' Summary: Loads a CSV file of connection test results into a dictionary keyed by host and port.
Private Function LoadCsvFile(filePath As String) As Object
    Dim dict As Object
    Dim fso As Object
    Dim ts As Object
    Dim line As String
    Dim parts As Variant
    Dim headerSkipped As Boolean

    Set dict = CreateObject("Scripting.Dictionary")
    If Dir(filePath) = "" Then
        Set LoadCsvFile = dict
        Exit Function
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, 1)
    headerSkipped = False
    Do Until ts.AtEndOfStream
        line = ts.ReadLine
        If Not headerSkipped Then
            headerSkipped = True
        Else
            parts = Split(line, ",")
            If UBound(parts) >= 4 Then
                dict(LCase(Trim(parts(1))) & "|" & Trim(parts(2))) = Array(Trim(parts(0)), LCase(Trim(parts(4))))
            End If
        End If
    Loop
    ts.Close
    Set LoadCsvFile = dict
End Function

' Summary: Checks CSV data for a host or IP and port and returns success flag and timestamp.
Private Function CheckCsvForHostPort(csvData As Object, hostOrIp As String, port As String) As Variant
    Dim key As String
    Dim arr As Variant

    key = LCase(Trim(hostOrIp)) & "|" & Trim(port)
    If csvData.Exists(key) Then
        arr = csvData(key)
        CheckCsvForHostPort = Array(LCase(arr(1)) = "open", arr(0))
    Else
        CheckCsvForHostPort = Array(False, "")
    End If
End Function


