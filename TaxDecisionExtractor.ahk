#NoEnv
SetBatchLines -1
#Include %A_ScriptDir%

; Configuration
global fileList := []
global fileStatus := {}
global currentFileIndex := 0
global currentData := {}
global isEditing := false

; Create GUI
Gui, Add, Button, x10 y10 w120 h35 cWhite, Chọn File
Gui, Add, Button, x140 y10 w120 h35 cWhite, Đọc File Word
Gui, Add, Button, x270 y10 w120 h35 cWhite, Chỉnh Sửa
Gui, Add, Button, x400 y10 w120 h35 cWhite, Tự Động Điền Web
Gui, Add, Button, x530 y10 w120 h35 cWhite, Đóng

; Left column - File list
Gui, Add, Text, x10 y60 w250 h20 cBlack, KET QUA:
Gui, Add, ListBox, x10 y85 w250 h400 vFileListBox gFileSelected, 

; Right column - Data display
Gui, Add, Text, x270 y60 w490 h20 cBlack, CHỈNH SỬA:
Gui, Add, Edit, x270 y85 w490 h20 vFieldName, 
Gui, Add, Text, x270 y110 w490 h15, Tên cá nhân / Tên tổ chức:

Gui, Add, Edit, x270 y130 w490 h20 vFieldBirthdate, 
Gui, Add, Text, x270 y155 w490 h15, Ngày, tháng, năm sinh:

Gui, Add, Edit, x270 y175 w490 h20 vFieldNationality, 
Gui, Add, Text, x270 y200 w490 h15, Quốc tịch:

Gui, Add, Edit, x270 y220 w490 h20 vFieldCCCD, 
Gui, Add, Text, x270 y245 w490 h15, Số CCCD:

Gui, Add, Edit, x270 y265 w490 h20 vFieldTaxID, 
Gui, Add, Text, x270 y290 w490 h15, Mã số thuế:

Gui, Add, Edit, x270 y310 w490 h20 vFieldViolation, 
Gui, Add, Text, x270 y335 w490 h15, Hành vi vi phạm:

Gui, Add, Edit, x270 y355 w490 h20 vFieldPenaltyType, 
Gui, Add, Text, x270 y380 w490 h15, Loại phạt:

Gui, Add, Edit, x270 y400 w490 h20 vFieldPenalty, 
Gui, Add, Text, x270 y425 w490 h15, Mức phạt:

; Buttons for edit
Gui, Add, Button, x270 y450 w120 h25, Lưu
Gui, Add, Button, x400 y450 w120 h25, Hủy

Gui, Add, StatusBar, x10 y490 w800 h20
SB_SetText("Sẵn sàng...")

Gui, Show, w800 h520, Bảng Điều Khiển - Tax Decision Extractor
return

ButtonChọnFile:
FileSelectFile, files, Multi, , Chọn file Word để đọc:
if (files != "")
{
    ; Parse multiple files
    Loop, Parse, files, `n
    {
        if (A_Index = 1 && A_LoopField != "")
        {
            basePath := A_LoopField
        }
        else if (A_Index > 1 && A_LoopField != "")
        {
            fullPath := basePath . "\" . A_LoopField
            if (!InArray(fileList, fullPath))
            {
                fileList.Push(fullPath)
                fileStatus[fullPath] := "Chưa đọc"
            }
        }
    }
    
    RefreshFileList()
    SB_SetText("Đã chọn " . fileList.Length() . " file")
}
return

ButtonĐọcFileWord:
if (currentFileIndex = 0 || currentFileIndex > fileList.Length())
{
    MsgBox, 48, Thông báo, Vui lòng chọn file từ danh sách trước!
    return
}

filePath := fileList[currentFileIndex]
SB_SetText("Đang đọc: " . GetFileName(filePath) . "...")

data := ExtractFromWord(filePath)
currentData := data

if (data.name != "" || data.organization != "")
{
    fileStatus[filePath] := "Đã đọc"
    RefreshFileList()
    DisplayData(data)
    SB_SetText("Đã đọc: " . GetFileName(filePath))
}
else
{
    fileStatus[filePath] := "Lỗi"
    RefreshFileList()
    SB_SetText("Lỗi: Không thể trích xuất dữ liệu từ " . GetFileName(filePath))
    MsgBox, 48, Lỗi, Không thể trích xuất dữ liệu từ file này!
}
return

ButtonChỉnhSửa:
isEditing := true
GuiControl, Enable, FieldName
GuiControl, Enable, FieldBirthdate
GuiControl, Enable, FieldNationality
GuiControl, Enable, FieldCCCD
GuiControl, Enable, FieldTaxID
GuiControl, Enable, FieldViolation
GuiControl, Enable, FieldPenaltyType
GuiControl, Enable, FieldPenalty
SB_SetText("Bạn có thể chỉnh sửa dữ liệu. Nhấn 'Lưu' để lưu hoặc 'Hủy' để bỏ qua")
return

ButtonLưu:
if (!isEditing)
    return

; Get edited values
GuiControlGet, fieldName, , FieldName
GuiControlGet, fieldBirthdate, , FieldBirthdate
GuiControlGet, fieldNationality, , FieldNationality
GuiControlGet, fieldCCCD, , FieldCCCD
GuiControlGet, fieldTaxID, , FieldTaxID
GuiControlGet, fieldViolation, , FieldViolation
GuiControlGet, fieldPenaltyType, , FieldPenaltyType
GuiControlGet, fieldPenalty, , FieldPenalty

; Update current data
currentData.name := fieldName
currentData.birthdate := fieldBirthdate
currentData.nationality := fieldNationality
currentData.cccd := fieldCCCD
currentData.taxid := fieldTaxID
currentData.violation := fieldViolation
currentData.penaltytype := fieldPenaltyType
currentData.penalty := fieldPenalty

; Disable editing
isEditing := false
GuiControl, Disable, FieldName
GuiControl, Disable, FieldBirthdate
GuiControl, Disable, FieldNationality
GuiControl, Disable, FieldCCCD
GuiControl, Disable, FieldTaxID
GuiControl, Disable, FieldViolation
GuiControl, Disable, FieldPenaltyType
GuiControl, Disable, FieldPenalty

SB_SetText("Đã lưu thay đổi")
MsgBox, 64, Thành công, Dữ liệu đã được lưu!
return

ButtonHủy:
if (!isEditing)
    return

; Reload previous data
DisplayData(currentData)

isEditing := false
GuiControl, Disable, FieldName
GuiControl, Disable, FieldBirthdate
GuiControl, Disable, FieldNationality
GuiControl, Disable, FieldCCCD
GuiControl, Disable, FieldTaxID
GuiControl, Disable, FieldViolation
GuiControl, Disable, FieldPenaltyType
GuiControl, Disable, FieldPenalty

SB_SetText("Hủy chỉnh sửa")
return

ButtonTựĐộngĐiềnWeb:
MsgBox, 48, Thông báo, Tính năng này sẽ được cập nhật sau!
return

ButtonĐóng:
ExitApp
return

FileSelected:
if (FileListBox = "")
    return

; Parse selected item to get file index
Loop, Parse, FileListBox, `n
{
    currentFileIndex := A_Index
    break
}

if (currentFileIndex > 0 && currentFileIndex <= fileList.Length())
{
    filePath := fileList[currentFileIndex]
    
    ; Try to load data if already read
    if (fileStatus[filePath] = "Đã đọc")
    {
        data := ExtractFromWord(filePath)
        if (data.name != "" || data.organization != "")
        {
            currentData := data
            DisplayData(data)
        }
    }
    else
    {
        ClearDisplay()
    }
}
return

RefreshFileList()
{
    list := ""
    Loop, % fileList.Length()
    {
        fileName := GetFileName(fileList[A_Index])
        status := fileStatus[fileList[A_Index]]
        list .= A_Index . ". " . fileName . " (" . status . ")`n"
    }
    GuiControl, , FileListBox, |%list%
}

DisplayData(data)
{
    name := data.name != "" ? data.name : data.organization
    GuiControl, , FieldName, %name%
    GuiControl, , FieldBirthdate, %data.birthdate%
    GuiControl, , FieldNationality, %data.nationality%
    GuiControl, , FieldCCCD, %data.cccd%
    GuiControl, , FieldTaxID, %data.taxid%
    GuiControl, , FieldViolation, %data.violation%
    GuiControl, , FieldPenaltyType, %data.penaltytype%
    GuiControl, , FieldPenalty, %data.penalty%
    
    ; Disable editing by default
    isEditing := false
    GuiControl, Disable, FieldName
    GuiControl, Disable, FieldBirthdate
    GuiControl, Disable, FieldNationality
    GuiControl, Disable, FieldCCCD
    GuiControl, Disable, FieldTaxID
    GuiControl, Disable, FieldViolation
    GuiControl, Disable, FieldPenaltyType
    GuiControl, Disable, FieldPenalty
}

ClearDisplay()
{
    GuiControl, , FieldName, 
    GuiControl, , FieldBirthdate, 
    GuiControl, , FieldNationality, 
    GuiControl, , FieldCCCD, 
    GuiControl, , FieldTaxID, 
    GuiControl, , FieldViolation, 
    GuiControl, , FieldPenaltyType, 
    GuiControl, , FieldPenalty, 
    
    isEditing := false
    GuiControl, Disable, FieldName
    GuiControl, Disable, FieldBirthdate
    GuiControl, Disable, FieldNationality
    GuiControl, Disable, FieldCCCD
    GuiControl, Disable, FieldTaxID
    GuiControl, Disable, FieldViolation
    GuiControl, Disable, FieldPenaltyType
    GuiControl, Disable, FieldPenalty
}

GetFileName(filePath)
{
    SplitPath, filePath, fileName
    return fileName
}

InArray(arr, value)
{
    for index, item in arr
    {
        if (item = value)
            return true
    }
    return false
}

ExtractFromWord(filePath)
{
    data := {}
    data.filename := GetFileName(filePath)
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
    data.aggravating := "Không"
    data.mitigating := "Không"
    data.penalty := ""
    data.penaltytype := ""
    data.account := ""
    data.bienban := ""
    
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
            data.name := ExtractValue(docText, "Tên tổ chức:", "Địa chỉ")
            data.organization := data.name
            data.orgaddress := ExtractValue(docText, "Địa chỉ trụ sở chính:", "Mã số thuế:")
            data.taxid := ExtractValue(docText, "Mã số thuế:", "Số GCN")
            data.representative := ExtractValue(docText, "Người đại diện theo pháp luật:", "Chức danh:")
            data.position := ExtractValue(docText, "Chức danh:", A_LF)
        }
        
        data.bienban := ExtractValue(docText, "Biên bản vi phạm hành chính về thuế số", "lập")
        data.violation := ExtractValue(docText, "thực hiện hành vi vi phạm hành chính:", "Quy định tại:")
        data.regulation := ExtractValue(docText, "Quy định tại:", "Các tình tiết")
        
        if (InStr(docText, "Phạt tiền"))
            data.penaltytype := "Phạt tiền"
        else if (InStr(docText, "Phạt cảnh cáo"))
            data.penaltytype := "Phạt cảnh cáo"
        
        data.penalty := ExtractValue(docText, "Mức phạt:", "đồng")
        data.account := ExtractValue(docText, "Tài khoản thu NSNN:", ",")
        
        objDoc.Close()
        objWord.Quit()
    }
    catch e
    {
        SB_SetText("Lỗi: " . e.Message)
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

GuiClose:
ExitApp
