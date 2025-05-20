Sub ExportCIMacroDraftsToCSV()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim data As Variant
    Dim outputData As String
    Dim specialChars As Object
    Dim FilePath As String
    Dim timestamp As String
    Dim i As Long
    Dim exportCount As Long
    Dim category As String, subCategory As String
    Dim desc As String, txt As String

    ' Disable screen updating and calculations to improve performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Set worksheet and table
    Set ws = ThisWorkbook.Worksheets("CI Macro Drafts")
    Set tbl = ws.ListObjects("CIMacroDrafts")
    data = tbl.DataBodyRange.Value ' Read all table data into an array

    ' Generate timestamp and set CSV file path
    timestamp = Format(Now, "yyyymmddhhss")
    FilePath = "C:\Users\phobrla\OneDrive - Carilion\Documents\PhraseExpress\CIMacroDrafts_" & timestamp & ".csv"

    ' Initialize variables
    outputData = ""
    exportCount = 0

    ' Initialize dictionary for special characters
    Set specialChars = CreateObject("Scripting.Dictionary")
    specialChars.Add "ENTER_BEFORE", "{#ENTER}"
    specialChars.Add "ENTER_AFTER", "{#ENTER -variablename New Line}"
    specialChars.Add "TAB", "{#TAB}"
    specialChars.Add "DEL", "{#DEL -count 15}"
    specialChars.Add "SLEEP", "{#sleep 1000}"
    specialChars.Add "INSERT", "{#insert -id 1F4D85EA-7001-48CF-88F7-F9E7012C27FE -variablename SNOW Acknowledged}"

    ' Cache column indices
    Dim colNeedsWork As Long, colConfigItem As Long
    Dim colCategory As Long, colDescription As Long
    Dim colShortDesc As Long, colExported As Long
    colNeedsWork = tbl.ListColumns("Needs Work").Index
    colConfigItem = tbl.ListColumns("SNOW Configuration item").Index
    colCategory = tbl.ListColumns("SNOW Category and Subcategory").Index
    colDescription = tbl.ListColumns("SNOW Description").Index
    colShortDesc = tbl.ListColumns("SNOW Short description").Index
    colExported = tbl.ListColumns("Exported").Index

    ' Loop through rows in array
    For i = 1 To UBound(data, 1)
        ' Skip rows if "Needs Work" is populated or required fields are empty
        If Trim(data(i, colNeedsWork)) = "" And _
           Trim(data(i, colConfigItem)) <> "" And _
           Trim(data(i, colCategory)) <> "" And _
           Trim(data(i, colDescription)) <> "" Then

            ' Build description
            If Trim(data(i, colShortDesc)) <> "" Then
                desc = data(i, colConfigItem) & ": " & data(i, colShortDesc)
            Else
                desc = ""
            End If

            ' Split "SNOW Category and Subcategory" into category and subcategory
            If InStr(1, data(i, colCategory), ">") > 0 Then
                category = Trim(Left(data(i, colCategory), InStr(1, data(i, colCategory), ">") - 1))
                subCategory = Trim(Mid(data(i, colCategory), InStr(1, data(i, colCategory), ">") + 1))
            Else
                category = data(i, colCategory)
                subCategory = ""
            End If

            ' Build txt string to match the requested format
            txt = data(i, colConfigItem) & _
                  specialChars("TAB") & specialChars("TAB") & specialChars("TAB") & specialChars("TAB") & specialChars("TAB") & _
                  category & specialChars("SLEEP") & specialChars("TAB") & subCategory & _
                  specialChars("SLEEP") & specialChars("TAB") & specialChars("TAB") & specialChars("TAB") & specialChars("TAB") & specialChars("TAB") & _
                  specialChars("DEL") & "TSG_TSC1" & specialChars("SLEEP") & specialChars("ENTER_BEFORE") & _
                  specialChars("TAB") & specialChars("TAB") & "Hobrla, Phil (Phil)" & specialChars("ENTER_BEFORE") & _
                  specialChars("SLEEP") & specialChars("TAB") & specialChars("INSERT") & _
                  data(i, colShortDesc) & specialChars("TAB") & data(i, colDescription)

            ' Replace new lines in description with "{#ENTER -variablename New Line}"
            desc = Replace(desc, vbLf, specialChars("ENTER_AFTER"))
            desc = Replace(desc, vbCrLf, specialChars("ENTER_AFTER"))

            ' Replace new lines in txt with "{#ENTER_BEFORE}" or "{#ENTER_AFTER}" based on position
            Dim txtParts() As String
            txtParts = Split(txt, specialChars("INSERT")) ' Split around INSERT marker
            txtParts(0) = Replace(txtParts(0), vbLf, specialChars("ENTER_BEFORE"))
            txtParts(0) = Replace(txtParts(0), vbCrLf, specialChars("ENTER_BEFORE"))
            txtParts(1) = Replace(txtParts(1), vbLf, specialChars("ENTER_AFTER"))
            txtParts(1) = Replace(txtParts(1), vbCrLf, specialChars("ENTER_AFTER"))
            txt = txtParts(0) & specialChars("INSERT") & txtParts(1)

            ' Escape double quotes by doubling them
            desc = Replace(desc, """", """""")
            txt = Replace(txt, """", """""")

            ' Enclose txt in quotes only if it contains a comma
            If InStr(txt, ",") > 0 Then txt = """" & txt & """"

            ' Append to output data
            If outputData <> "" Then outputData = outputData & vbCrLf ' Add newline only between rows
            outputData = outputData & desc & "," & txt & "," & "_SNOW Macros (SERVICES)"

            ' Mark as exported
            data(i, colExported) = "Yes"
            exportCount = exportCount + 1
        End If
    Next i

    ' Write all data to CSV file
    Open FilePath For Output As #1
    Print #1, outputData
    Close #1

    ' Write updated data back to the worksheet
    tbl.DataBodyRange.Value = data

    ' Restore Excel settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    ' Notify user
    MsgBox "CSV file created successfully with " & exportCount & " rows exported: " & FilePath
End Sub
