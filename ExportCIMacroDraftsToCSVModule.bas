Attribute VB_Name = "ExportCIMacroDraftsToCSVModule"
Sub ExportCIMacroDraftsToCSV()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rng As Range
    Dim desc As String, txt As String
    Dim folder As String
    Dim filePath As String
    Dim timestamp As String
    Dim i As Long
    Dim exportCount As Long
    Dim category As String
    Dim subCategory As String
    Dim specialChars As Object

    ' Set worksheet and table
    Set ws = ThisWorkbook.Worksheets("CI Macro Drafts")
    Set tbl = ws.ListObjects("CIMacroDrafts")

    ' Generate timestamp
    timestamp = Format(Now, "yyyymmddhhss")

    ' Set CSV file path
    filePath = "C:\Users\phobrla\OneDrive - Carilion\Documents\PhraseExpress\CIMacroDrafts_" & timestamp & ".csv"

    ' Create CSV file
    Open filePath For Output As #1

    ' Initialize export count
    exportCount = 0

    ' Initialize dictionary for special characters
    Set specialChars = CreateObject("Scripting.Dictionary")
    specialChars.Add "ENTER", "{#ENTER}"
    specialChars.Add "SPACE", "{#SPACE}"
    specialChars.Add "TAB", "{#TAB}"
    specialChars.Add "DEL", "{#DEL -count 15}"
    specialChars.Add "SLEEP", "{#sleep 1000}"
    specialChars.Add "BKSP", "{#BKSP}"
    specialChars.Add "INSERT", "{#insert -id 1F4D85EA-7001-48CF-88F7-F9E7012C27FE -variablename SNOW Acknowledged}"

    ' Loop through table rows
    For i = 1 To tbl.ListRows.Count
        Set rng = tbl.ListRows(i).Range

        ' Check if any of the required fields are blank
        If Trim(rng.Cells(1, tbl.ListColumns("SNOW Configuration item").Index).Value) = "" Or _
           Trim(rng.Cells(1, tbl.ListColumns("SNOW Category and Subcategory").Index).Value) = "" Or _
           Trim(rng.Cells(1, tbl.ListColumns("SNOW Description").Index).Value) = "" Then
            ' Skip this row if any required field is blank
            GoTo NextIteration
        End If

        ' Check if "SNOW Short description" is populated
        If rng.Cells(1, tbl.ListColumns("SNOW Short description").Index).Value <> "" Then
            desc = rng.Cells(1, tbl.ListColumns("SNOW Configuration item").Index).Value & ": " & rng.Cells(1, tbl.ListColumns("SNOW Short description").Index).Value
        Else
            desc = ""
        End If

        ' Only proceed if Description is not blank
        If Trim(desc) <> "" Then
            ' Split "SNOW Category and Subcategory" into category and subcategory
            If InStr(1, rng.Cells(1, tbl.ListColumns("SNOW Category and Subcategory").Index).Value, ">") > 0 Then
                category = Trim(Left(rng.Cells(1, tbl.ListColumns("SNOW Category and Subcategory").Index).Value, _
                InStr(1, rng.Cells(1, tbl.ListColumns("SNOW Category and Subcategory").Index).Value, ">") - 1))
                subCategory = Trim(Mid(rng.Cells(1, tbl.ListColumns("SNOW Category and Subcategory").Index).Value, _
                InStr(1, rng.Cells(1, tbl.ListColumns("SNOW Category and Subcategory").Index).Value, ">") + 1))
            Else
                category = rng.Cells(1, tbl.ListColumns("SNOW Category and Subcategory").Index).Value
                subCategory = ""
            End If

            ' Check if "SNOW Configuration item" starts with "ASSET TAG"
            If Left(rng.Cells(1, tbl.ListColumns("SNOW Configuration item").Index).Value, 9) <> "ASSET TAG" Then
                txt = rng.Cells(1, tbl.ListColumns("SNOW Configuration item").Index).Value & specialChars("TAB") & specialChars("TAB") & specialChars("TAB") & specialChars("TAB") & _
                      category & specialChars("SLEEP") & specialChars("TAB") & subCategory & _
                      specialChars("SLEEP") & specialChars("TAB") & specialChars("TAB") & specialChars("TAB") & specialChars("TAB") & specialChars("TAB") & specialChars("DEL") & "TSG_TSC1" & specialChars("SLEEP") & specialChars("TAB") & "phobrla" & specialChars("SLEEP") & _
                      specialChars("INSERT") & _
                      rng.Cells(1, tbl.ListColumns("SNOW Short description").Index).Value & specialChars("TAB") & _
                      rng.Cells(1, tbl.ListColumns("SNOW Description").Index).Value
            Else
                txt = specialChars("SPACE") & specialChars("BKSP") & specialChars("TAB") & specialChars("TAB") & specialChars("TAB") & specialChars("TAB") & specialChars("TAB") & _
                      category & specialChars("SLEEP") & specialChars("TAB") & subCategory & _
                      specialChars("SLEEP") & specialChars("TAB") & specialChars("TAB") & specialChars("TAB") & specialChars("TAB") & specialChars("TAB") & specialChars("DEL") & "TSG_TSC1" & specialChars("SLEEP") & specialChars("TAB") & "phobrla" & specialChars("SLEEP") & _
                      specialChars("INSERT") & _
                      rng.Cells(1, tbl.ListColumns("SNOW Short description").Index).Value & specialChars("TAB") & _
                      rng.Cells(1, tbl.ListColumns("SNOW Description").Index).Value
            End If

            folder = "_SNOW Macros (SERVICES)"

            ' Replace new lines with "{#ENTER}"
            desc = Replace(desc, vbLf, specialChars("ENTER"))
            txt = Replace(txt, vbLf, specialChars("ENTER"))
            desc = Replace(desc, vbCrLf, specialChars("ENTER"))
            txt = Replace(txt, vbCrLf, specialChars("ENTER"))
            desc = Replace(desc, vbCrLf, specialChars("ENTER"))
            txt = Replace(txt, vbCrLf, specialChars("ENTER"))

            ' Escape double quotes
            desc = Replace(desc, """", """""")
            txt = Replace(txt, """", """""")

            ' Escape commas by enclosing the field in double quotes
            If InStr(desc, ",") > 0 Then desc = """" & desc & """"
            If InStr(txt, ",") > 0 Then txt = """" & txt & """"

            ' Write to CSV file with comma delimiter
            Print #1, desc & "," & txt & "," & folder

            ' Mark as Exported
            rng.Cells(1, tbl.ListColumns("Exported").Index).Value = "Yes"
            ' Increment export count
            exportCount = exportCount + 1
        End If

NextIteration:
    Next i

    ' Close CSV file
    Close #1

    MsgBox "CSV file created successfully with " & exportCount & " rows exported: " & filePath

End Sub

