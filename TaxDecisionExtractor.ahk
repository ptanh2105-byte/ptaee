#NoEnv
SetBatchLines -1
#Include %A_ScriptDir%

; Configuration
global resultsArray := []
global selectedFolder := ""
global isProcessing := false

; Create GUI
Gui, Add, Button, x10 y10 w150 h30, SelectFolder
Gui, Add, Button, x170 y10 w150 h30, StartExtract
Gui, Add, Button, x330 y10 w100 h30, CopyAll
Gui, Add, Button, x440 y10 w120 h30, ExportExcel
Gui, Add, Button, x570 y10 w100 h30, ExportCSV
Gui, Add, Button, x680 y10 w80 h30, Clear

Gui, Add, Text, x10 y50 w760 h20 cBlue vFolderText, Folder: (Not selected)
Gui, Add, ListBox, x10 y75 w760 h350 Multi vResultsList, 

Gui, Add, StatusBar, x10 y430 w760 h20
SB_SetText("Ready...")

Gui, Show, w780 h455, Tax Decision Extractor v1.0
return

ButtonSelectFolder:
FileSelectFolder, folder, , 1, Select folder containing Word files:
if (folder != "")
{
    selectedFolder := folder
    GuiControl, , FolderText, Folder: %folder%
    SB_SetText("Selected: " . folder)
}
return

ButtonStartExtract:
if (selectedFolder = "")
{
    MsgBox, 48, Notice, Please select a folder first!
    return
}
ProcessFiles()
return

ButtonCopyAll:
if (resultsArray.Length() = 0)
{
    MsgBox, 48, Notice, No data to copy!
    return
}
CopyToClipboard()
return

ButtonExportExcel:
if (resultsArray.Length() = 0)
{
    MsgBox, 48, Notice, No data to export!
    return
}
ExportToExcel()
return

ButtonExportCSV:
if (resultsArray.Length() = 0)
{
    MsgBox, 48, Notice, No data to export!
    return
}
ExportToCSV()
return

ButtonClear:
GuiControl, , ResultsList, |
resultsArray := []
SB_SetText("Data cleared")
return

ProcessFiles()
{
    global selectedFolder, resultsArray, isProcessing
    
    isProcessing := true
    GuiControl, Disable, Button2
    resultsArray := []
    GuiControl, , ResultsList, |
    
    fileCount := 0
    processedCount := 0
    
    ; Count total files first
    Loop, Files, %selectedFolder%\*.docx
        fileCount++
    Loop, Files, %selectedFolder%\*.doc
        fileCount++
    
    SB_SetText("Found " . fileCount . " Word files. Starting extraction...")
    Sleep, 500
    
    ; Process .docx files
    Loop, Files, %selectedFolder%\*.docx
    {
        SB_SetText("Processing (.docx): " . A_LoopFileName . " (" . (processedCount+1) . "/" . fileCount . ")")
        Sleep, 100
        
        data := ExtractFromWord(A_LoopFilePath)
        if (data.name != "" || data.organization != "")
        {
            resultsArray.Push(data)
            AddToListView(data)
        }
        processedCount++
    }
    
    ; Process .doc files
    Loop, Files, %selectedFolder%\*.doc
    {
        SB_SetText("Processing (.doc): " . A_LoopFileName . " (" . (processedCount+1) . "/" . fileCount . ")")
        Sleep, 100
        
        data := ExtractFromWord(A_LoopFilePath)
        if (data.name != "" || data.organization != "")
        {
            resultsArray.Push(data)
            AddToListView(data)
        }
        processedCount++
    }
    
    GuiControl, Enable, Button2
    isProcessing := false
    SB_SetText("Completed! Extracted " . resultsArray.Length() . " records from " . processedCount . " files")
}

ExtractFromWord(filePath)
{
    data := {}
    data.filename := A_LoopFileName
    data.name := ""
    data.birthdate := ""
    data.nationality := ""
    data.cccd := ""
    data.taxid := ""
    data.workplace := ""
    data.address := ""
    data.organization := ""
    data.orgaddress := ""
    data.representative := ""
    data.position := ""
    data.violation := ""
    data.regulation := ""
    data.aggravating := "None"
    data.mitigating := "None"
    data.penalty := ""
    data.penaltytype := ""
    data.account := ""
    data.bienban := ""
    data.ngayBienban := ""
    
    try
    {
        objWord := ComObjCreate("Word.Application")
        objWord.Visible := false
        objWord.ScreenUpdating := false
        
        objDoc := objWord.Documents.Open(filePath)
        docText := objDoc.Content.Text
        
        ; Extract data from document
        if (InStr(docText, "Tên cá nhân:"))
        {
            data.name := ExtractValue(docText, "Tên cá nhân:", "Giới tính:")
            data.birthdate := ExtractValue(docText, "Ngày, tháng, năm sinh:", "Quốc tịch:")
            data.nationality := ExtractValue(docText, "Quốc tịch:", "Nơi làm việc:")
            data.workplace := ExtractValue(docText, "Nơi làm việc:", "Nơi ở hiện tại:")
            data.address := ExtractValue(docText, "Nơi ở hiện tại:", "Số CCCD:")
            data.cccd := ExtractValue(docText, "Số CCCD:", "Mã số thuế:")
            data.taxid := ExtractValue(docText, "Mã số thuế:", A_LF)
        }
        else if (InStr(docText, "Tên tổ chức:"))
        {
            data.organization := ExtractValue(docText, "Tên tổ chức:", "Địa chỉ")
            data.orgaddress := ExtractValue(docText, "Địa chỉ trụ sở chính:", "Mã số thuế:")
            data.taxid := ExtractValue(docText, "Mã số thuế:", "Số GCN")
            data.representative := ExtractValue(docText, "Người đại diện theo pháp luật:", "Chức danh:")
            data.position := ExtractValue(docText, "Chức danh:", A_LF)
        }
        
        data.bienban := ExtractValue(docText, "Biên bản vi phạm hành chính về thuế số", "lập")
        data.violation := ExtractValue(docText, "thực hiện hành vi vi phạm hành chính:", "Quy định tại:")
        data.regulation := ExtractValue(docText, "Quy định tại:", "Các tình tiết")
        data.aggravating := ExtractValue(docText, "Các tình tiết tăng nặng (nếu có):", "Các tình tiết giảm nhẹ")
        data.mitigating := ExtractValue(docText, "Các tình tiết giảm nhẹ (nếu có):", "Bị áp dụng")
        
        if (InStr(docText, "Phạt tiền"))
            data.penaltytype := "Phat tien"
        else if (InStr(docText, "Phạt cảnh cáo"))
            data.penaltytype := "Phat canh cao"
        
        data.penalty := ExtractValue(docText, "Mức phạt:", "đồng")
        data.account := ExtractValue(docText, "Tài khoản thu NSNN:", ",")
        
        objDoc.Close()
        objWord.Quit()
    }
    catch e
    {
        SB_SetText("Error processing: " . A_LoopFileName . " - " . e.Message)
    }
    
    return data
}

ExtractValue(text, startMarker, endMarker)
{
    startPos := InStr(text, startMarker)
    if (startPos = 0)
        return ""
    
    startPos += StrLen(startMarker)
    endPos := InStr(text, endMarker, false, startPos)
    
    if (endPos = 0)
        endPos := StrLen(text)
    
    result := SubStr(text, startPos, endPos - startPos)
    result := Trim(result)
    
    ; Clean up result
    result := StrReplace(result, A_Tab, " ")
    result := StrReplace(result, A_CR, " ")
    result := StrReplace(result, A_LF, " ")
    
    Loop
    {
        if (InStr(result, "  "))
            result := StrReplace(result, "  ", " ")
        else
            break
    }
    
    return Trim(result)
}

AddToListView(data)
{
    global ResultsList
    
    name := data.name != "" ? data.name : data.organization
    violation := SubStr(data.violation, 1, 30) . (StrLen(data.violation) > 30 ? "..." : "")
    penalty := data.penaltytype . (data.penalty != "" ? " - " . SubStr(data.penalty, 1, 15) : "")
    
    GuiControl, , ResultsList, % name . " | " . penalty . " | " . data.bienban
}

CopyToClipboard()
{
    text := ""
    for index, item in resultsArray
    {
        text .= "=== Record " . index . " ===" . A_LF
        text .= "File Name: " . item.filename . A_LF
        
        if (item.name != "")
        {
            text .= "Name: " . item.name . A_LF
            text .= "Birth Date: " . item.birthdate . A_LF
            text .= "Nationality: " . item.nationality . A_LF
            text .= "ID: " . item.cccd . A_LF
            text .= "Workplace: " . item.workplace . A_LF
            text .= "Address: " . item.address . A_LF
        }
        else if (item.organization != "")
        {
            text .= "Organization: " . item.organization . A_LF
            text .= "Address: " . item.orgaddress . A_LF
            text .= "Representative: " . item.representative . A_LF
            text .= "Position: " . item.position . A_LF
        }
        
        text .= "Tax ID: " . item.taxid . A_LF
        text .= "Decision No: " . item.bienban . A_LF
        text .= "Violation: " . item.violation . A_LF
        text .= "Regulation: " . item.regulation . A_LF
        text .= "Penalty Type: " . item.penaltytype . A_LF
        text .= "Penalty Amount: " . item.penalty . A_LF
        text .= "Account: " . item.account . A_LF
        text .= A_LF
    }
    
    A_Clipboard := text
    SB_SetText("Copied " . resultsArray.Length() . " records to clipboard")
    MsgBox, 64, Success, Data copied to clipboard!
}

ExportToExcel()
{
    FileSelectFile, outputFile, S16, , Export to Excel, *.xlsx
    if (outputFile = "")
        return
    
    try
    {
        objExcel := ComObjCreate("Excel.Application")
        objExcel.Visible := true
        objWorkbook := objExcel.Workbooks.Add()
        objSheet := objWorkbook.Sheets(1)
        
        ; Create headers
        headers := ["STT", "File Name", "Name/Org", "Tax ID", "Decision No", "Penalty Type", "Amount", "Violation"]
        for col, header in headers
            objSheet.Cells(1, col).Value := header
        
        ; Add data
        row := 2
        for index, item in resultsArray
        {
            name := item.name != "" ? item.name : item.organization
            objSheet.Cells(row, 1).Value := index
            objSheet.Cells(row, 2).Value := item.filename
            objSheet.Cells(row, 3).Value := name
            objSheet.Cells(row, 4).Value := item.taxid
            objSheet.Cells(row, 5).Value := item.bienban
            objSheet.Cells(row, 6).Value := item.penaltytype
            objSheet.Cells(row, 7).Value := item.penalty
            objSheet.Cells(row, 8).Value := item.violation
            row++
        }
        
        objSheet.Columns("A:H").AutoFit()
        objWorkbook.SaveAs(outputFile)
        objWorkbook.Close()
        objExcel.Quit()
        
        SB_SetText("Exported to Excel: " . outputFile)
        MsgBox, 64, Success, Data exported to Excel!
    }
    catch e
    {
        MsgBox, 48, Error, Cannot export to Excel: %e%
    }
}

ExportToCSV()
{
    FileSelectFile, outputFile, S16, , Export to CSV, *.csv
    if (outputFile = "")
        return
    
    try
    {
        file := FileOpen(outputFile, "w", "UTF-8-RAW")
        if !IsObject(file)
        {
            MsgBox, 48, Error, Cannot create CSV file
            return
        }
        
        file.Write(Chr(0xFEFF))
        
        headers := "STT,File Name,Name/Org,Tax ID,Decision No,Penalty Type,Amount,Violation"
        file.Write(headers . "`n")
        
        for index, item in resultsArray
        {
            name := item.name != "" ? item.name : item.organization
            line := index . "," . CSVEscape(item.filename) . "," . CSVEscape(name) . "," . CSVEscape(item.taxid) . "," . CSVEscape(item.bienban) . "," . CSVEscape(item.penaltytype) . "," . CSVEscape(item.penalty) . "," . CSVEscape(item.violation)
            file.Write(line . "`n")
        }
        
        file.Close()
        SB_SetText("Exported to CSV: " . outputFile)
        MsgBox, 64, Success, Data exported to CSV!
    }
    catch e
    {
        MsgBox, 48, Error, Cannot export to CSV: %e%
    }
}

CSVEscape(text)
{
    if (InStr(text, ",") || InStr(text, """") || InStr(text, "`n"))
    {
        text := StrReplace(text, """", """""")
        return """" . text . """"
    }
    return text
}

GuiClose:
ExitApp
