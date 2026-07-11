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
Gui, Add, Button, x10 y10 w120 h35, Chon File
Gui, Add, Button, x140 y10 w120 h35, Doc File Word
Gui, Add, Button, x270 y10 w120 h35, Chinh Sua
Gui, Add, Button, x400 y10 w120 h35, Tu Dong Dien Web
Gui, Add, Button, x530 y10 w120 h35, Dong

; Left column - File list
Gui, Add, Text, x10 y60 w250 h20, KET QUA:
Gui, Add, ListBox, x10 y85 w250 h400 vFileListBox gFileSelected, 

; Right column - Data display
Gui, Add, Text, x270 y60 w490 h20, CHINH SUA:
Gui, Add, Edit, x270 y85 w490 h20 vFieldName Disabled, 
Gui, Add, Text, x270 y110 w490 h15, Ten ca nhan / Ten to chuc:

Gui, Add, Edit, x270 y130 w490 h20 vFieldBirthDate Disabled, 
Gui, Add, Text, x270 y155 w490 h15, Ngay sinh:

Gui, Add, Edit, x270 y175 w490 h20 vFieldNationality Disabled, 
Gui, Add, Text, x270 y200 w490 h15, Quoc tich:

Gui, Add, Edit, x270 y220 w490 h20 vFieldCCCD Disabled, 
Gui, Add, Text, x270 y245 w490 h15, So CCCD:

Gui, Add, Edit, x270 y265 w490 h20 vFieldTaxID Disabled, 
Gui, Add, Text, x270 y290 w490 h15, Ma so thue:

Gui, Add, Edit, x270 y310 w490 h20 vFieldViolation Disabled, 
Gui, Add, Text, x270 y335 w490 h15, Hanh vi vi pham:

Gui, Add, Edit, x270 y355 w490 h20 vFieldPenaltyType Disabled, 
Gui, Add, Text, x270 y380 w490 h15, Loai phat:

Gui, Add, Edit, x270 y400 w490 h20 vFieldPenalty Disabled, 
Gui, Add, Text, x270 y425 w490 h15, Muc phat:

; Buttons for edit
Gui, Add, Button, x270 y450 w120 h25, Luu
Gui, Add, Button, x400 y450 w120 h25, Huy

Gui, Add, StatusBar, x10 y490 w800 h20
SB_SetText("San sang...")

Gui, Show, w800 h520, Bang Dieu Khien - Tax Decision Extractor
return

ButtonChonFile:
FileSelectFile, files, Multi, , Chon file Word de doc:
if (files != "")
{
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
                fileStatus[fullPath] := "Chua doc"
            }
        }
    }
    
    RefreshFileList()
    SB_SetText("Da chon " . fileList.Length() . " file")
}
return

ButtonDocFileWord:
if (currentFileIndex = 0 || currentFileIndex > fileList.Length())
{
    MsgBox, 48, Thong bao, Vui long chon file tu danh sach truoc!
    return
}

filePath := fileList[currentFileIndex]
SB_SetText("Dang doc: " . GetFileName(filePath) . "...")

data := ExtractFromWord(filePath)
currentData := data

if (data.name != "" || data.organization != "")
{
    fileStatus[filePath] := "Da doc"
    RefreshFileList()
    DisplayData(data)
    SB_SetText("Da doc: " . GetFileName(filePath))
}
else
{
    fileStatus[filePath] := "Loi"
    RefreshFileList()
    SB_SetText("Loi: Khong the trich xuat du lieu tu " . GetFileName(filePath))
    MsgBox, 48, Loi, Khong the trich xuat du lieu tu file nay!
}
return

ButtonChinhSua:
isEditing := true
GuiControl, Enable, FieldName
GuiControl, Enable, FieldBirthDate
GuiControl, Enable, FieldNationality
GuiControl, Enable, FieldCCCD
GuiControl, Enable, FieldTaxID
GuiControl, Enable, FieldViolation
GuiControl, Enable, FieldPenaltyType
GuiControl, Enable, FieldPenalty
SB_SetText("Ban co the chinh sua du lieu. Nhan 'Luu' de luu hoac 'Huy' de bo qua")
return

ButtonLuu:
if (!isEditing)
    return

GuiControlGet, fieldName, , FieldName
GuiControlGet, fieldBirthDate, , FieldBirthDate
GuiControlGet, fieldNationality, , FieldNationality
GuiControlGet, fieldCCCD, , FieldCCCD
GuiControlGet, fieldTaxID, , FieldTaxID
GuiControlGet, fieldViolation, , FieldViolation
GuiControlGet, fieldPenaltyType, , FieldPenaltyType
GuiControlGet, fieldPenalty, , FieldPenalty

currentData.name := fieldName
currentData.birth_date := fieldBirthDate
currentData.nationality := fieldNationality
currentData.cccd := fieldCCCD
currentData.tax_id := fieldTaxID
currentData.violation := fieldViolation
currentData.penalty_type := fieldPenaltyType
currentData.penalty := fieldPenalty

isEditing := false
GuiControl, Disable, FieldName
GuiControl, Disable, FieldBirthDate
GuiControl, Disable, FieldNationality
GuiControl, Disable, FieldCCCD
GuiControl, Disable, FieldTaxID
GuiControl, Disable, FieldViolation
GuiControl, Disable, FieldPenaltyType
GuiControl, Disable, FieldPenalty

SB_SetText("Da luu thay doi")
MsgBox, 64, Thanh cong, Du lieu da duoc luu!
return

ButtonHuy:
if (!isEditing)
    return

DisplayData(currentData)

isEditing := false
GuiControl, Disable, FieldName
GuiControl, Disable, FieldBirthDate
GuiControl, Disable, FieldNationality
GuiControl, Disable, FieldCCCD
GuiControl, Disable, FieldTaxID
GuiControl, Disable, FieldViolation
GuiControl, Disable, FieldPenaltyType
GuiControl, Disable, FieldPenalty

SB_SetText("Huy chinh sua")
return

ButtonTuDongDienWeb:
MsgBox, 48, Thong bao, Tinh nang nay se duoc cap nhat sau!
return

ButtonDong:
ExitApp
return

FileSelected:
if (FileListBox = "")
    return

Loop, Parse, FileListBox, `n
{
    currentFileIndex := A_Index
    break
}

if (currentFileIndex > 0 && currentFileIndex <= fileList.Length())
{
    filePath := fileList[currentFileIndex]
    
    if (fileStatus[filePath] = "Da doc")
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
    GuiControl, , FieldBirthDate, %data.birth_date%
    GuiControl, , FieldNationality, %data.nationality%
    GuiControl, , FieldCCCD, %data.cccd%
    GuiControl, , FieldTaxID, %data.tax_id%
    GuiControl, , FieldViolation, %data.violation%
    GuiControl, , FieldPenaltyType, %data.penalty_type%
    GuiControl, , FieldPenalty, %data.penalty%
    
    isEditing := false
    GuiControl, Disable, FieldName
    GuiControl, Disable, FieldBirthDate
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
    GuiControl, , FieldBirthDate, 
    GuiControl, , FieldNationality, 
    GuiControl, , FieldCCCD, 
    GuiControl, , FieldTaxID, 
    GuiControl, , FieldViolation, 
    GuiControl, , FieldPenaltyType, 
    GuiControl, , FieldPenalty, 
    
    isEditing := false
    GuiControl, Disable, FieldName
    GuiControl, Disable, FieldBirthDate
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
    data.birth_date := ""
    data.nationality := ""
    data.cccd := ""
    data.tax_id := ""
    data.workplace := ""
    data.address := ""
    data.organization := ""
    data.org_address := ""
    data.representative := ""
    data.position := ""
    data.violation := ""
    data.regulation := ""
    data.aggravating := "Khong"
    data.mitigating := "Khong"
    data.penalty := ""
    data.penalty_type := ""
    data.account := ""
    data.bienban := ""
    
    try
    {
        objWord := ComObjCreate("Word.Application")
        objWord.Visible := false
        objWord.ScreenUpdating := false
        
        objDoc := objWord.Documents.Open(filePath)
        docText := objDoc.Content.Text
        
        if (InStr(docText, "Ten ca nhan:") || InStr(docText, "Tên cá nhân:"))
        {
            data.name := ExtractValue(docText, "Ten ca nhan:", "Gioi tinh:")
            if (data.name = "")
                data.name := ExtractValue(docText, "Tên cá nhân:", "Giới tính:")
            data.birth_date := ExtractValue(docText, "Ngay, thang, nam sinh:", "Quoc tich:")
            if (data.birth_date = "")
                data.birth_date := ExtractValue(docText, "Ngày, tháng, năm sinh:", "Quốc tịch:")
            data.nationality := ExtractValue(docText, "Quoc tich:", "Noi lam viec:")
            if (data.nationality = "")
                data.nationality := ExtractValue(docText, "Quốc tịch:", "Nơi làm việc:")
            data.workplace := ExtractValue(docText, "Noi lam viec:", "Noi o hien tai:")
            if (data.workplace = "")
                data.workplace := ExtractValue(docText, "Nơi làm việc:", "Nơi ở hiện tại:")
            data.address := ExtractValue(docText, "Noi o hien tai:", "So CCCD:")
            if (data.address = "")
                data.address := ExtractValue(docText, "Nơi ở hiện tại:", "Số CCCD:")
            data.cccd := ExtractValue(docText, "So CCCD:", "Ma so thue:")
            if (data.cccd = "")
                data.cccd := ExtractValue(docText, "Số CCCD:", "Mã số thuế:")
            data.tax_id := ExtractValue(docText, "Ma so thue:", A_LF)
            if (data.tax_id = "")
                data.tax_id := ExtractValue(docText, "Mã số thuế:", A_LF)
        }
        else if (InStr(docText, "Ten to chuc:") || InStr(docText, "Tên tổ chức:"))
        {
            data.name := ExtractValue(docText, "Ten to chuc:", "Dia chi")
            if (data.name = "")
                data.name := ExtractValue(docText, "Tên tổ chức:", "Địa chỉ")
            data.organization := data.name
            data.org_address := ExtractValue(docText, "Dia chi tru so chinh:", "Ma so thue:")
            if (data.org_address = "")
                data.org_address := ExtractValue(docText, "Địa chỉ trụ sở chính:", "Mã số thuế:")
            data.tax_id := ExtractValue(docText, "Ma so thue:", "So GCN")
            if (data.tax_id = "")
                data.tax_id := ExtractValue(docText, "Mã số thuế:", "Số GCN")
            data.representative := ExtractValue(docText, "Nguoi dai dien theo phap luat:", "Chuc danh:")
            if (data.representative = "")
                data.representative := ExtractValue(docText, "Người đại diện theo pháp luật:", "Chức danh:")
            data.position := ExtractValue(docText, "Chuc danh:", A_LF)
            if (data.position = "")
                data.position := ExtractValue(docText, "Chức danh:", A_LF)
        }
        
        data.bienban := ExtractValue(docText, "Bien ban vi pham hanh chinh ve thue so", "lap")
        if (data.bienban = "")
            data.bienban := ExtractValue(docText, "Biên bản vi phạm hành chính về thuế số", "lập")
        data.violation := ExtractValue(docText, "thuc hien hanh vi vi pham hanh chinh:", "Quy dinh tai:")
        if (data.violation = "")
            data.violation := ExtractValue(docText, "thực hiện hành vi vi phạm hành chính:", "Quy định tại:")
        data.regulation := ExtractValue(docText, "Quy dinh tai:", "Cac tinh tiet")
        if (data.regulation = "")
            data.regulation := ExtractValue(docText, "Quy định tại:", "Các tình tiết")
        
        if (InStr(docText, "Phat tien") || InStr(docText, "Phạt tiền"))
            data.penalty_type := "Phat tien"
        else if (InStr(docText, "Phat canh cao") || InStr(docText, "Phạt cảnh cáo"))
            data.penalty_type := "Phat canh cao"
        
        data.penalty := ExtractValue(docText, "Muc phat:", "dong")
        if (data.penalty = "")
            data.penalty := ExtractValue(docText, "Mức phạt:", "đồng")
        data.account := ExtractValue(docText, "Tai khoan thu NSNN:", ",")
        if (data.account = "")
            data.account := ExtractValue(docText, "Tài khoản thu NSNN:", ",")
        
        objDoc.Close()
        objWord.Quit()
    }
    catch e
    {
        SB_SetText("Loi: " . e.Message)
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
