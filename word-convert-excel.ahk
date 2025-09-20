#Requires AutoHotkey v2.0
#SingleInstance Force

; 热键：按ESC退出
Esc::
{
    MsgBox("正在退出程序...", "退出", "T1")
    ExitApp
}


time := GetBaiduTime()

; 主程序开始
Time_Line_Date := GetBaiduTime()
time := StrReplace(Time_Line_Date,"年","")
time := StrReplace(time,"月","")
time := StrReplace(time,"日","")

EndDate := "2026年12月31日"
Endtime := StrReplace(EndDate,"年","")
Endtime := StrReplace(Endtime,"月","")
Endtime := StrReplace(Endtime,"日","")

if (time > Endtime) {
    MsgBox("今天是" EndDate "以后，程序即将退出。",,"T1")
    ExitApp
}

SetWorkingDir A_ScriptDir
loop files, A_WorkingDir "\*.exe", "R" ; 直接查找pandoc.exe
{
    if (InStr(A_LoopFileFullPath, "pandoc.exe")>0) {
        pandocPath := A_LoopFileFullPath
        break
    } else {
        pandocPath := ""
        continue
    }
}

MsgBox("调用的命令行工具路径：" pandocPath,,"T1")

Transpose_Folder_Path := GetFolder_Path(pandocPath) "\Transpose"

if !DirExist(Transpose_Folder_Path) {
    DirCreate(Transpose_Folder_Path)
}

MsgBox("调用的命令行工具所在文件夹：" Transpose_Folder_Path,,"T1")
SetWorkingDir Transpose_Folder_Path
loop files, A_WorkingDir "\*.txt", "R" 
{
    FileDelete(A_LoopFileFullPath) 
}

SetWorkingDir A_ScriptDir
loop files, A_WorkingDir "\*.xlsx", "R" 
{
    if (InStr(A_LoopFileFullPath, "汇总")>0) {
        FileDelete(A_LoopFileFullPath) 
        continue
    } else {
        continue
    }
}

SetWorkingDir Transpose_Folder_Path
loop files, A_WorkingDir "\*.txt", "R" 
{
    if(InStr(A_LoopFileFullPath, "Transpose")>0) {
        FileDelete(A_LoopFileFullPath) 
    } else {
        continue
    }
}

SetWorkingDir A_ScriptDir
Cmd_Script_List := ""
loop files, A_WorkingDir "\*.docx", "R" 
{
    if (InStr(A_LoopFileFullPath, "pandoc-3.7.0.2")>0) {
        continue
    } else {
        docxFile := A_LoopFileFullPath
        SplitPath(docxFile, &name, &dir, &ext, &name_noext)  ; 修正：获取name_noext
        txtFile := Transpose_Folder_Path "\" name_noext ".txt" ; 输出的txt文件路径
        Cmd_Script := "`"" pandocPath "`"" " " "`"" docxFile "`"" " -o" " " "`"" txtFile "`"" " --wrap=none"

        Cmd_Script_List:=Cmd_Script_List Cmd_Script "`n"
    }
}

; MsgBox(Cmd_Script_List,,"T1")
; MsgBox("保存的txt文件夹路径：" Transpose_Folder_Path,,"T1")
; Sleep 1000

loop
{
    if(FileExist(Transpose_Folder_Path)=true) {
        DirDelete(Transpose_Folder_Path, 1)
    } else
    {
        MsgBox("未检测到Transpose文件夹，程序开始创建","检查","T1")
        DirCreate(Transpose_Folder_Path)
        Sleep 1000
        MsgBox("程序已创建Transpose文件夹","检查","T1")
        break
    }    
}


ids := WinGetList("ahk_exe cmd.exe")
for id in ids
{
    WinClose("ahk_id " id)
    Sleep 500
}

BlockInput("On")
InputBlockGui := Gui(, "输入已禁用")
InputBlockGui.AddText("w220 h40 Center", "当前输入已禁用，请等待自动处理完成……")
InputBlockGui.Show("w240 h60")

Num_Loop_Range := 25
Cmd_Script_List_Array := StrSplit(Cmd_Script_List, "`n", "`r")
if (Mod(Cmd_Script_List_Array.Length, Num_Loop_Range)!=0) {
    Num_loop := Floor(Cmd_Script_List_Array.Length/Num_Loop_Range)+1
}
else {
    Num_loop := Floor(Cmd_Script_List_Array.Length/Num_Loop_Range)
}

loop Num_loop
{
    Start_Index := (A_Index - 1) * Num_Loop_Range + 1
    End_Index := Start_Index + Num_Loop_Range - 1
    if (End_Index > Cmd_Script_List_Array.Length) {
        End_Index := Cmd_Script_List_Array.Length
    }
    Cmd_Script_List_Sub := ""
    loop (End_Index - Start_Index + 1)
    {
        command := Cmd_Script_List_Array[Start_Index + A_Index - 1] "`n"
        RunWait('cmd /c "' command '"', , "Hide")
    }

    Sleep 1000
}

BlockInput("Off")
if (IsSet(InputBlockGui) && InputBlockGui) {
    InputBlockGui.Destroy()
    InputBlockGui := ""
}

Sleep 1000


ids := WinGetList("ahk_exe cmd.exe")
for id in ids
{
    WinClose("ahk_id " id)
    Sleep 500
}

SetWorkingDir Transpose_Folder_Path
loop files, A_WorkingDir "\*.txt", "R" ; 直接查找
{
    Content_One := FileRead(A_LoopFileFullPath) ; 读取旧的txt文件内容
    Content_One := StrReplace(Content_One, "单选题", "`r`n【单选题】") 
    Content_One := StrReplace(Content_One, "多选题", "`r`n【多选题】") 
    Content_One := StrReplace(Content_One, "判断题", "`r`n【判断题】") 
    Content_One := Trim(Content_One, "`r`n")
    Content_Saved_List := "Mpa`r`n"
                      . "kV`r`n"
                      . "m/s`r`n"
                      . "mm`r`n"
                      . "ms`r`n"
                      . "℃`r`n"
                      . "°`r`n"
                      . "`r`n"
                      . "`r`n"
                      . "`r`n"
                      . "`r`n"
                      . ""
    Content_Saved_List := StrReplaceAll(Content_Saved_List, OldString := "`r`n`r`n", NewString := "`r`n")
    Content_Saved_List := Trim(Content_Saved_List, "`r`n")

    Content_One := StrReplaceAll(Content_One, OldString := "`r`n- A ．`r`n", NewString := " A ．")
    Content_One := StrReplaceAll(Content_One, OldString := "`r`n- B ．`r`n", NewString := " B ．")
    Content_One := StrReplaceAll(Content_One, OldString := "`r`n- C ．`r`n", NewString := " C ．")
    Content_One := StrReplaceAll(Content_One, OldString := "`r`n- D ．`r`n", NewString := " D ．")
    Content_One := StrReplaceAll(Content_One, OldString := "`r`n- E ．`r`n", NewString := " E ．")
    Content_One := StrReplaceAll(Content_One, OldString := "`r`n- F ．`r`n", NewString := " F ．")
    Content_One := StrReplaceAll(Content_One, OldString := "`r`n- G ．`r`n", NewString := " G ．")
    Content_One := StrReplaceAll(Content_One, OldString := "． `r`n", NewString := "． ")
	Content_One := StrReplaceAll(Content_One, OldString := "．`r`n", NewString := "．")
	Content_One := StrReplaceAll(Content_One, OldString := "`r`n- `r`n", NewString := "`t")
	Content_One := StrReplaceAll(Content_One, OldString := "`r`n  A  B  C  D  E  F  G`r`n", NewString := "")
	Content_One := StrReplaceAll(Content_One, OldString := "`r`n  A  B  C  D  E  F`r`n", NewString := "")
	Content_One := StrReplaceAll(Content_One, OldString := "`r`n  A  B  C  D  E`r`n", NewString := "")
	Content_One := StrReplaceAll(Content_One, OldString := "`r`n  A  B  C  D`r`n", NewString := "")
	Content_One := StrReplaceAll(Content_One, OldString := "`r`n  A  B  C`r`n", NewString := "")
	Content_One := StrReplaceAll(Content_One, OldString := "`r`n  A  B`r`n", NewString := "")
	Content_One := StrReplaceAll(Content_One, OldString := "`r`nA  B  C  D  E  F  G`r`n", NewString := "")
	Content_One := StrReplaceAll(Content_One, OldString := "`r`nA  B  C  D  E  F`r`n", NewString := "")
	Content_One := StrReplaceAll(Content_One, OldString := "`r`nA  B  C  D  E`r`n", NewString := "")
	Content_One := StrReplaceAll(Content_One, OldString := "`r`nA  B  C  D`r`n", NewString := "")
	Content_One := StrReplaceAll(Content_One, OldString := "`r`nA  B  C`r`n", NewString := "")
	Content_One := StrReplaceAll(Content_One, OldString := "`r`nA  B`r`n", NewString := "")
	Content_One := StrReplaceAll(Content_One, OldString := "．`r`n", NewString := "． ")
	Content_One := StrReplaceAll(Content_One, OldString := "`r`n`r`n", NewString := "`r`n")
	Content_One := StrReplaceAll(Content_One, OldString := "`t`t", NewString := "`t")
	Content_One := StrReplaceAll(Content_One, OldString := "`r`n`t", NewString := "")
	Content_One := StrReplaceAll(Content_One, OldString := "`r`n ", NewString := "`r`n")
	Content_One := StrReplaceAll(Content_One, OldString := " ．`r`n ", NewString := " ．")
	Content_One := StrReplaceAll(Content_One, OldString := "`r`nA B C D E F G", NewString := "")
	Content_One := StrReplaceAll(Content_One, OldString := "`r`nA B C D E F", NewString := "")
	Content_One := StrReplaceAll(Content_One, OldString := "`r`nA B C D E", NewString := "")
	Content_One := StrReplaceAll(Content_One, OldString := "`r`nA B C D", NewString := "")
	Content_One := StrReplaceAll(Content_One, OldString := "`r`nA B C", NewString := "")
	Content_One := StrReplaceAll(Content_One, OldString := "`r`nA B", NewString := "")
	Content_One := StrReplaceAll(Content_One, OldString := "  A  B  C  D  E  F  G`r`n", NewString := "")
	Content_One := StrReplaceAll(Content_One, OldString := "  A  B  C  D  E  F`r`n", NewString := "")
	Content_One := StrReplaceAll(Content_One, OldString := "  A  B  C  D  E`r`n", NewString := "")
	Content_One := StrReplaceAll(Content_One, OldString := "  A  B  C  D`r`n", NewString := "")
	Content_One := StrReplaceAll(Content_One, OldString := "  A  B  C`r`n", NewString := "")
	Content_One := StrReplaceAll(Content_One, OldString := "  A  B`r`n", NewString := "")
	Content_One := StrReplaceAll(Content_One, OldString := "`r`nA", NewString := "`tA")
	Content_One := StrReplaceAll(Content_One, OldString := "`r`nB", NewString := "  B")
	Content_One := StrReplaceAll(Content_One, OldString := "`r`nC", NewString := "  C")
	Content_One := StrReplaceAll(Content_One, OldString := "`r`nD", NewString := "  D")
	Content_One := StrReplaceAll(Content_One, OldString := "`r`nE", NewString := "  E")
	Content_One := StrReplaceAll(Content_One, OldString := "`r`nF", NewString := "  F")
	Content_One := StrReplaceAll(Content_One, OldString := "`r`nG", NewString := "  G")
	Content_One := StrReplaceAll(Content_One, OldString := "`r`n ", NewString := " ")
	Content_One := StrReplaceAll(Content_One, OldString := "  ", NewString := " ")

    
	Content_One := StrReplaceAll(Content_One, OldString := "A ．正确 B ．错误", NewString := "A ．错误 B ．正确")
    Content_One := StrReplace(Content_One, "\`r`n", "")
    Content_One := StrReplace(Content_One, "正确答案：", "`t正确答案：")
    Content_One := StrReplace(Content_One, "题：", "题：`t")
    Content_One := Trim(Content_One, "`r`n")
    Content_One := Trim(Content_One, "  A  B")
    Content_One := Trim(Content_One, "`r`n- ")
    Content_One_Array := StrSplit(Content_One, "`n", "`r")
    Content_One_Length := Content_One_Array.Length
    Content_One_Summary := ""
    loop Content_One_Length
    {
        if (InStr(Content_One_Array[A_Index],"![IMG_")>0 or Content_One_Array[A_Index]="") {
            Content_One_Summary := Content_One_Summary "" 
        } else {
            Content_One_Summary := Content_One_Summary "" Content_One_Array[A_Index] "`n"
        }
    }

    Content_One_Summary := Trim(Content_One_Summary, "`r`n")
    FileDelete(A_LoopFileFullPath) ; 删除旧的txt文件
    FileAppend(Content_One_Summary, A_LoopFileFullPath) ; 写入新的txt文件内容
    Content_One_Summary := ""
}

MsgBox("所有txt文件内容已整理完成",,"T1")
name := "试题汇总.txt"
name_noext := StrReplace(name, ".txt", "")
; MsgBox(Transpose_Folder_Path "\Transpose",,"T1")

NewtxtSummaryFilePath := Transpose_Folder_Path "\" name_noext ".txt"

SetWorkingDir Transpose_Folder_Path
txtFile_List := ""
loop files, A_WorkingDir "\*.txt", "R" ; 直接查找
{
    txtFile_List := txtFile_List A_LoopFileFullPath "`n" ; 删除旧的txt文件
}

txtFile_List := Trim(txtFile_List, "`r`n")
txtFile_List_Array := StrSplit(txtFile_List, "`n", "`r")
txtFile_List_Length := txtFile_List_Array.Length
loop txtFile_List_Length
{
    Old_txtFilePath := txtFile_List_Array[A_Index]
    ; MsgBox(A_LoopFileFullPath,,"T1")
    Content_One := FileRead(Old_txtFilePath) ; 读取旧的txt文件内容
    Content_One := Trim(Content_One, "`r`n")

    if (InStr(Content_One,"单选题")>0 or InStr(Content_One,"多选题")>0 or InStr(Content_One,"判断题")>0) {
        Content_One_Array := StrSplit(Content_One, "`n", "`r")
        Count_Zone_Summary := ""
        String_One_Zone_Summary := ""
        loop Content_One_Array.Length
        {
            Count_Self := (InStr(Content_One_Array[A_Index], "单选题")>0) "" (InStr(Content_One_Array[A_Index], "多选题")>0) "" (InStr(Content_One_Array[A_Index], "判断题")>0)
            Count_Zone_Summary := Count_Zone_Summary "" Count_Self "`n"
        }

        Count_Zone_Summary := Trim(Count_Zone_Summary, "`r`n")
        Count_Zone_Summary_Array := StrSplit(Count_Zone_Summary, "`n", "`r")
        Count_Zone_Summary_Length := Count_Zone_Summary_Array.Length
        Content_One_Index_End := 0
        Content_One_Index_End_Summary := ""
        loop Count_Zone_Summary_Length
        {
            Content_One_String := Content_One_Array[A_Index]
            Content_One_Index := InStr(Count_Zone_Summary_Array[A_Index], "1")
            Content_One_Index_End := Max(Content_One_Index_End ,Content_One_Index)
            Content_One_Index_End_Summary := Content_One_Index_End_Summary "" Content_One_Index_End "`n"
        }

        ; MsgBox("最大索引：`n" Content_One_Index_End_Summary,,"T1")

        Content_One_Index_End_Summary := Trim(Content_One_Index_End_Summary, "`r`n")
        Content_One_Index_End_Summary_Array := StrSplit(Content_One_Index_End_Summary, "`n", "`r")
        Content_One_Index_End_Summary_Length := Content_One_Index_End_Summary_Array.Length

        loop Content_One_Index_End_Summary_Length
        {
            String_One_Zone := Count_Zone_Summary_Array[A_Index] "`t" Content_One_Index_End_Summary_Array[A_Index] "`t" Content_One_Array[A_Index]
            String_One_Zone_Summary := String_One_Zone_Summary "" String_One_Zone "`n"
        }

        ; MsgBox("最终结果：`n" String_One_Zone_Summary,,"T1")

        String_One_Zone_Summary := Trim(String_One_Zone_Summary, "`r`n")
        FileDelete(Old_txtFilePath) ; 删除旧的txt文件
        FileAppend(String_One_Zone_Summary "`n", Old_txtFilePath) ; 写入新的txt文件内容
        String_One_Zone_Summary := ""
    }
}


loop txtFile_List_Length
{
    Old_txtFilePath := txtFile_List_Array[A_Index]
    ; MsgBox(A_LoopFileFullPath,,"T1")
    String_One_Zone_Summary := FileRead(Old_txtFilePath) ; 读取旧的txt文件内容
    String_One_Zone_Summary := Trim(String_One_Zone_Summary, "`r`n")

    String_One_Zone_Summary_Array := StrSplit(String_One_Zone_Summary, "`n", "`r")
    String_One_Zone_Summary_Length := String_One_Zone_Summary_Array.Length
    String_One_Zone_Summary_New := ""
    Content_Index_String := "单选题`n多选题`n判断题`n其他"
    Content_Index_String_Array := StrSplit(Content_Index_String, "`n", "`r")
    ; MsgBox(String_One_Zone_Summary_Length,,"T1")
    Num_Loop_Range := 1000
    FileDelete(Old_txtFilePath) ; 删除旧的txt文件
    loop String_One_Zone_Summary_Length
    {
        String_One_Start := String_One_Zone_Summary_Array[A_Index]
        String_One_Array := StrSplit(String_One_Start, "`t")
        Content_Index_Second := String_One_Array[2]
        Content_Index_First := String_One_Array[1]
        if (InStr(Content_Index_First, "1")>0) {
            Content_Index_Second := 0
        } else if (Content_Index_Second ="0" or Content_Index_Second ="1") {
            Content_Index_Second := 1
        }
        else if (Content_Index_Second ="2") {
            Content_Index_Second := 2
        }
        else if (Content_Index_Second ="3") {
            Content_Index_Second := 3
        }
        else {
            Content_Index_Second := 4
        }

        if (Content_Index_Second>4 or Content_Index_Second<=0) {
            String_One_Pre := ""
        } else {
            String_One_Pre := Content_Index_Second "`t" Content_Index_String_Array[Content_Index_Second]
        }

        if (String_One_Pre="") {
            String_One := ""
        } else {
            String_One_old := String_One_Pre "`t" StrReplace(StrReplace(String_One_Start, String_One_Array[1] "`t", "", 1), String_One_Array[2] "`t", "", 1) 
            String_One_Array := StrSplit(String_One_old, "`t")
            try {
                String_One := String_One_Array[1] "`t" String_One_Array[2] "`t" String_One_Array[3] "`t" String_One_Array[4] "`t" String_One_Array[6]
            } catch {
                Content_Error_Log := "错误的文件路径：`n" Old_txtFilePath "`n错误的行内容：`n" String_One_Start "`n请检查文件内容是否符合规范，此条记录将被跳过。`n`n"
                FileAppend(Content_Error_Log, A_ScriptDir "\_错误日志.txt")
                continue
            }
            ; 这里可以在最后做一下细微调整
            String_One := Trim(String_One, "`t")
            String_One := Trim(String_One, " ")
            End_Part := String_One_Array[5]
            End_Part := Trim(End_Part, "`t")
            End_Part := Trim(End_Part, " ")
            ; -----------------------------------------
            ; 处理判断题答案反转的问题和其他题型的答案提取
            if (InStr(String_One,"判断题")>0 && (End_Part="正确答案：A")) {
                String_One := String_One "`t" "B" "`n"
            } else if (InStr(String_One,"判断题")>0 && (End_Part="正确答案：B")) {
                String_One := String_One "`t" "A" "`n"
            } else if (InStr(String_One,"多选题")>0) {
                End_Part_New := End_Part
                End_Part_New := StrReplace(End_Part_New, "正确答案：A", "正确答案：,A")
                End_Part_New := StrReplace(End_Part_New, "正确答案：B", "正确答案：,B")
                End_Part_New := StrReplace(End_Part_New, "正确答案：C", "正确答案：,C")
                End_Part_New := StrReplace(End_Part_New, "正确答案：D", "正确答案：,D")
                End_Part_New := StrReplace(End_Part_New, "正确答案：E", "正确答案：,E")
                End_Part_New := StrReplace(End_Part_New, "正确答案：F", "正确答案：,F")
                End_Part_New := StrReplace(End_Part_New, "正确答案：G", "正确答案：,G")
                End_Part_New := StrReplace(End_Part_New, "G", ",G")
                End_Part_New := StrReplace(End_Part_New, "F", ",F")
                End_Part_New := StrReplace(End_Part_New, "E", ",E")
                End_Part_New := StrReplace(End_Part_New, "D", ",D")
                End_Part_New := StrReplace(End_Part_New, "C", ",C")
                End_Part_New := StrReplace(End_Part_New, "B", ",B")
                End_Part_New := StrReplace(End_Part_New, "A", ",A")
                loop 
                {
                    End_Part_New := StrReplace(End_Part_New, ",,", ",",, &Count)
                    if (Count=0)
                        break
                }
                End_Part_New := StrReplace(End_Part_New, "正确答案：,", "")
                String_One := String_One "`t" End_Part_New "`n"
            } else {
                String_One := String_One "`t" StrReplace(End_Part,"正确答案：","") "`n"
            }

        }

        String_One_Zone_Summary_New := String_One_Zone_Summary_New "" String_One
        ; 如果是最后一个元素，修剪内容
        if (A_Index = String_One_Zone_Summary_Length) {
            String_One_Zone_Summary_New := Trim(String_One_Zone_Summary_New, "`r`n")
        }

        ; 如果达到循环范围或是最后一个元素，则写入文件
        if (Mod(A_Index, Num_Loop_Range) = 0 || A_Index = String_One_Zone_Summary_Length) {
            FileAppend(String_One_Zone_Summary_New, Old_txtFilePath)
            String_One_Zone_Summary_New := ""
        }
    }
}

MsgBox("所有txt文件内容已整理完成",,"T1")
SetWorkingDir Transpose_Folder_Path
Count_txtFile := 0
File_Path_list := ""
loop files, A_WorkingDir "\*.txt", "R" ; 直接查找
{
    File_Path := A_LoopFileFullPath
    Count_txtFile++
    File_Path_list := File_Path_list "" File_Path "`n"
}

MsgBox("检测到txt文件数量：" Count_txtFile,,"T1")

File_Path_list := Trim(File_Path_list, "`r`n")
File_Path_list_Array := StrSplit(File_Path_list, "`n", "`r")

if (FileExist(NewtxtSummaryFilePath)) {
    FileDelete(NewtxtSummaryFilePath) ; 删除旧的txt文件
}

loop Count_txtFile
{
    File_Path := File_Path_list_Array[A_Index]
    Content_Index_First := FileRead(File_Path) ; 读取旧的txt文件内容
    Content_Index_First := Trim(Content_Index_First, "`r`n")
    if (A_Index = Count_txtFile) {
        Content_Index_First := Content_Index_First
    } else {
        Content_Index_First := Content_Index_First "`n"
    } 
    FileAppend(Content_Index_First, NewtxtSummaryFilePath)
}

Content_Index_Summary:= FileRead(NewtxtSummaryFilePath) ; 读取旧的txt文件内容
Content_Index_Summary := Trim(Content_Index_Summary, "`r`n")
Content_Index_Summary_Array:= StrSplit(Content_Index_Summary, "`n", "`r")
Content_Index_Summary_Length := Content_Index_Summary_Array.Length
MsgBox("最终汇总的txt文件内容行数：`n" Content_Index_Summary_Length,,"T1")

Num_Loop_Range := 1000
String_One_Zone_Summary_New := ""
Content_Index_Summary := FileRead(NewtxtSummaryFilePath) ; 删除旧的txt文件
FileDelete(NewtxtSummaryFilePath) ; 删除旧的txt文件
Content_Index_Summary := Trim(Content_Index_Summary, "`r`n")
Content_Index_Summary_Array := StrSplit(Content_Index_Summary, "`n", "`r")
if(Mod(Content_Index_Summary_Array.Length, Num_Loop_Range)!=0) {
    Num_loop := Floor(Content_Index_Summary_Array.Length/Num_Loop_Range)+1
}
else {
    Num_loop := Floor(Content_Index_Summary_Array.Length/Num_Loop_Range)
}

loop Num_loop
{
    Start_Index := (A_Index - 1) * Num_Loop_Range + 1
    End_Index := Start_Index + Num_Loop_Range - 1
    if (End_Index > Content_Index_Summary_Array.Length) {
        End_Index := Content_Index_Summary_Array.Length
    }
    Content_Index_Summary_Sub := ""
    loop (End_Index - Start_Index + 1)
    {
        Content_Index_Summary_Sub := Content_Index_Summary_Sub "" Content_Index_Summary_Array[Start_Index + A_Index - 1] "`n"
    }
    Content_Index_Summary_Sub := Trim(Content_Index_Summary_Sub, "`r`n")
    if (A_Index = Num_loop) {
        Content_Index_Summary_Sub := Content_Index_Summary_Sub
    } else {
        Content_Index_Summary_Sub := Content_Index_Summary_Sub "`n"
    }

    FileAppend(Content_Index_Summary_Sub, NewtxtSummaryFilePath)
    Sleep 500
}

; 调用优化后的COM函数
excelFilePath := CreateExcelFileFromTxt_File(NewtxtSummaryFilePath)
; MsgBox("生成的Excel文件路径：" excelFilePath,"提示","T1")
; MsgBox("转换完成，文件保存在：" A_ScriptDir,"提示","T1")
Run("explorer.exe " A_ScriptDir)
Sleep 2000

MsgBox("正在退出程序...", "退出", "T3")
ExitApp

; 结束Excel进程的函数
KillExcelProcesses() {
    excelFound := false
    excelCount := 0
    
    try {
        ; 方法1: 使用WMI查询Excel进程
        for process in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process Where Name = 'EXCEL.EXE'") {
            excelFound := true
            excelCount++
            try {
                process.Terminate()
                Sleep(100) ; 稍微等待一下
            }
        }
        
        ; 方法2: 使用taskkill命令确保所有Excel进程都被结束
        if (excelFound) {
            RunWait("taskkill /f /im excel.exe", , "Hide")
        }
        
        ; 等待一段时间确保进程完全结束
        Sleep(500)
        
        ; 显示查找结果提示
        if (excelCount > 0) {
            MsgBox("找到 " excelCount " 个Excel进程，已全部结束。", "Excel进程检查", "T2")
        } else {
            MsgBox("未找到正在运行的Excel进程。", "Excel进程检查", "T2")
        }
        
    } catch as e {
        ; 如果结束进程失败，显示错误信息
        if (excelCount > 0) {
            MsgBox("找到 " excelCount " 个Excel进程，但结束进程时出错: " e.Message, "Excel进程检查", "T2")
        } else {
            MsgBox("未找到Excel进程，检查过程中出错: " e.Message, "Excel进程检查", "T2")
        }
    }
    
    return excelCount
}

; 改进的设置函数（返回对象而不是字符串）
SetExcelCalculation(excel, preferredMode := -4135) {
    result := {}
    result.success := false
    result.mode := ""
    result.message := ""
    result.type := ""  ; "preferred", "alternative", "guessed", "failed"
    
    knownModes := [-4135, -4105, -4104]
    
    ; 尝试已知模式
    for mode in knownModes {
        try {
            excel.Calculation := mode
            result.success := true
            result.mode := mode
            result.message := "计算模式设置为: " mode
            
            if (mode = preferredMode) {
                result.type := "preferred"
                result.message .= " (首选模式)"
            } else {
                result.type := "alternative"
                result.message .= " (备选模式)"
            }
            return result
        } catch {
            Continue
        }
    }
    
    ; 尝试智能猜测
    smartGuesses := [preferredMode, -4135, -4105, -4104, 0, 1, 2, 3, -1, -2, -3]
    
    for mode in smartGuesses {
        try {
            excel.Calculation := mode
            result.success := true
            result.mode := mode
            result.type := "guessed"
            result.message := "计算模式设置为: " mode " (猜测模式)"
            return result
        } catch {
            Continue
        }
    }
    
    ; 所有尝试都失败
    result.type := "failed"
    result.message := "无法设置计算模式"
    return result
}

; 使用方式
excel := ComObject("Excel.Application")
calcResult := SetExcelCalculation(excel, -4135)

; 根据结果类型处理
switch calcResult.type {
    case "preferred":
        ; 完美，无需处理
    case "alternative":
        MsgBox("警告: " calcResult.message, "Excel计算模式检查", "T2")
    case "guessed":
        MsgBox("信息: " calcResult.message, "Excel计算模式检查", "T2")
    case "failed":
        MsgBox("错误: " calcResult.message "，性能可能受影响", "Excel计算模式检查", "T2")
}

; 优化的CreateExcelFileFromTxt函数
CreateExcelFileFromTxt_Line(txtFilePath) {
    startTime := A_TickCount
    excel := ""
    workbook := ""
    
    try {
        ; 1. 先检查并结束所有Excel进程
        excelProcessCount := KillExcelProcesses()
        if (excelProcessCount > 0) {
            Sleep(500)
        }
        
        ; 2. 创建Excel对象（禁用所有不必要的功能）
        excel := ComObject("Excel.Application")
        excel.Visible := true ; 修改为 true，方便观察导入效果
        excel.DisplayAlerts := false
        excel.ScreenUpdating := false
        excel.EnableEvents := false

        calcResult := SetExcelCalculation(excel, -4135)

        ; 新增：直接导入指定 txt 文件
        if FileExist(txtFilePath) {
            excel.Workbooks.OpenText(txtFilePath)
        } else {
            MsgBox("未找到指定的TXT文件: " txtFilePath, "错误")
        }
        
        ; 3. 批量读取txt文件内容
        content := FileRead(txtFilePath)
        if (content = "") {
            throw Error("文件内容为空")
        }
        
        ; 4. 快速解析内容
        lines := StrSplit(Trim(content, "`r`n"), "`n", "`r")
        totalLines := lines.Length
        if (totalLines = 0) {
            throw Error("没有有效数据行")
        }
        
        ; 5. 创建进度提示
        progressGui := Gui("+ToolWindow +AlwaysOnTop", "处理进度")
        progressGui.Add("Text", "w300", "正在处理数据...")
        progressBar := progressGui.Add("Progress", "w300 h20")
        progressText := progressGui.Add("Text", "w300", "0/" totalLines " 行")
        progressGui.Show()
        
        ; 6. 预分配数组大小
        rowCount := totalLines + 1
        colCount := 6
        excelArray := ComObjArray(12, rowCount, colCount)
        
        ; 7. 批量填充数据
        excelArray[0, 0] := "题型"
        excelArray[0, 1] := "题目"
        excelArray[0, 2] := "题号"
        excelArray[0, 3] := "题目内容"
        excelArray[0, 4] := "选项"
        excelArray[0, 5] := "答案"
        
        loop totalLines {
            if (Mod(A_Index, 100) = 0) {
                progressBar.Value := (A_Index / totalLines) * 100
                progressText.Text := A_Index "/" totalLines " 行"
            }
            
            line := lines[A_Index]
            if (line = "") {
                continue
            }
            
            parts := StrSplit(line, "`t")
            partsLength := parts.Length
            
            loop Min(partsLength, colCount) {
                excelArray[A_Index, A_Index - 1] := parts[A_Index]
            }
            
            if (partsLength < colCount) {
                loop (colCount - partsLength) {
                    excelArray[A_Index, partsLength + A_Index - 1] := ""
                }
            }
        }
        
        ; 8. 创建工作簿和工作表
        workbook := excel.Workbooks.Add()
        worksheet := workbook.Worksheets(1)
        worksheet.Name := "数据"
        
        ; 9. 一次性写入所有数据
        progressText.Text := "正在写入Excel..."
        worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(rowCount, colCount)).Value := excelArray
        
        ; 10. 批量设置格式
        progressText.Text := "正在设置格式..."
        headerRange := worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(1, colCount))
        headerRange.Interior.Color := 0xCCE5FF
        headerRange.Font.Bold := true
        headerRange.HorizontalAlignment := -4108  ; xlCenter
        
        worksheet.Columns.AutoFit()
        
        ; 11. 保存文件
        progressText.Text := "正在保存文件..."
        SplitPath(txtFilePath, &name, , , &name_noext)
        currentTime := FormatTime(, "yyyyMMdd_HHmmss")
        excelFilePath := A_ScriptDir "\试题汇总_" currentTime ".xlsx"
        
        workbook.SaveAs(excelFilePath)
        workbook.Close(false)
        
        ; 12. 清理COM对象
        excel.Quit()
        excel := ""
        workbook := ""
        
        ; 计算耗时
        endTime := A_TickCount
        timeTaken := (endTime - startTime) / 1000
        
        progressGui.Destroy()
        
        MsgBox("Excel文件创建完成！`n耗时: " timeTaken " 秒`n文件: " excelFilePath, "完成", "T3")
        return excelFilePath
        
    } catch as e {
        try progressGui.Destroy()
        try workbook.Close(false)
        try excel.Quit()
        
        MsgBox("Excel创建失败: " e.Message "`n文件: " e.File "`n行号: " e.Line, "错误", "T30")
        return ""
    }
}

; 使用Excel导入文本的版本
CreateExcelFileFromTxt_File(txtFilePath) {
    startTime := A_TickCount
    excel := ""
    workbook := ""
    
    try {
        ; 1. 先检查并结束所有Excel进程
        excelProcessCount := KillExcelProcesses()
        if (excelProcessCount > 0) {
            Sleep(500)
        }
        
        ; 2. 创建Excel对象
        excel := ComObject("Excel.Application")
        excel.Visible := true
        excel.DisplayAlerts := false

        ; 3. 创建新工作簿
        workbook := excel.Workbooks.Add()
        worksheet := workbook.Worksheets(1)

        ; 4. 使用QueryTables导入到指定位置
        progressGui := Gui("+ToolWindow +AlwaysOnTop", "处理进度")
        progressGui.Add("Text", "w300", "正在导入TXT文件...")
        progressGui.Show()

        ; 创建查询表
        queryTable := worksheet.QueryTables.Add(
            "TEXT;" txtFilePath,  ; Connection
            worksheet.Range("A1") ; Destination
        )

        ; 设置导入参数
        queryTable.TextFilePlatform := 65001      ; UTF-8编码
        queryTable.TextFileStartRow := 1          ; 起始行
        queryTable.TextFileParseType := 1         ; 分隔符方式
        queryTable.TextFileTextQualifier := 1     ; 文本限定符
        queryTable.TextFileConsecutiveDelimiter := true
        queryTable.TextFileTabDelimiter := true   ; Tab分隔符
        queryTable.TextFileSemicolonDelimiter := false
        queryTable.TextFileCommaDelimiter := false
        queryTable.TextFileSpaceDelimiter := false

        ; 执行导入
        queryTable.Refresh(false)  ; 后台刷新

        ; 删除查询表（保留数据）
        queryTable.Delete()
        
        ; 获取活动工作簿和工作表
        workbook := excel.ActiveWorkbook
        worksheet := workbook.Worksheets(1)


        ; 4. 设置格式
        progressGui.Destroy()
        progressGui := Gui("+ToolWindow +AlwaysOnTop", "处理进度")
        progressGui.Add("Text", "w300", "正在设置格式...")
        progressGui.Show()
        
        ; 添加表头（如果TXT文件没有表头）
        lastRow := worksheet.Cells(worksheet.Rows.Count, 1).End(-4162).Row
        lastCol := worksheet.Cells(1, worksheet.Columns.Count).End(-4159).Column
        
        ; 插入表头行
        worksheet.Rows(1).Insert()
        headers := ["题型", "题目", "题号", "题目内容", "选项", "答案"]
        loop Min(headers.Length, lastCol) {
            worksheet.Cells(1, A_Index).Value := headers[A_Index]
        }
        
        ; 设置表头格式
        headerRange := worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(1, lastCol))
        headerRange.Interior.Color := 0xCCE5FF
        headerRange.Font.Bold := true
        headerRange.HorizontalAlignment := -4108
        
        worksheet.Columns.AutoFit()
        
        ; 5. 保存文件
        SplitPath(txtFilePath, , , , &name_noext)
        currentTime := FormatTime(, "yyyyMMdd_HHmmss")
        excelFilePath := A_ScriptDir "\试题汇总_" currentTime ".xlsx"
        
        workbook.SaveAs(excelFilePath, 51) ; 51 表示 xlsx 格式
        workbook.Close(false)
        excel.Quit()
        
        ; 计算耗时
        endTime := A_TickCount
        timeTaken := (endTime - startTime) / 1000
        
        progressGui.Destroy()
        
        MsgBox("Excel文件创建完成！`n耗时: " timeTaken " 秒`n文件: " excelFilePath, "完成", "T3")
        return excelFilePath
        
    } catch as e {
        try progressGui.Destroy()
        try workbook.Close(false)
        try excel.Quit()
        
        MsgBox("Excel创建失败: " e.Message "`n文件: " e.File "`n行号: " e.Line, "错误", "T30")
        return ""
    }
}

; 如果需要，可以保留ShowComList函数用于调试，但不要在主流程中调用
ShowComList() {
    MyGui := Gui(, "Process List")
    LV := MyGui.Add("ListView", "x2 y0 w1400 h1000", ["Process Name", "Command Line"])
    process_List := ""
    for process in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process") {
        if (process.CommandLine = "") {
            continue
        }
        process_List := process_List "" process.Name ": " process.CommandLine "`n"
    }
    process_List := Trim(process_List, "`r`n")
    FileAppend(process_List, A_ScriptDir "\Process_List.txt")
    return process_List
}

; 最简单的百度时间获取
GetBaiduTime() {
    curlPath := ""
    ids:=WinGetList("ahk_exe cmd.exe")
    for this_id in ids
    {
        WinClose(this_id)
        Sleep(100)
    }

    SetWorkingDir A_ScriptDir
    loop files, A_WorkingDir "\*.exe", "R" ; 直接查找curl.exe
    {
        if (InStr(A_LoopFileFullPath, "curl.exe")>0) {
            curlPath := A_LoopFileFullPath
            break
        } else {
            curlPath := ""
            continue
        }
    }

    if (curlPath = "")
    {
        MsgBox("未找到 curl.exe，请确保 curl.exe 在脚本目录或子目录中。程序几秒钟以后将自动退出。", "错误", "T10")
        ExitApp
    }
    global curlPath
    baiduTime := ""
    day_Line:=" 01 `n 02 `n 03 `n 04 `n 05 `n 06 `n 07 `n 08 `n 09 `n 10 `n 11 `n 12 `n 13 `n 14 `n 15`n 16 `n 17`n 18`n 19`n 20`n 21`n 22`n 23`n 24`n 25`n 26`n 27`n 28`n 29`n 30`n 31 `n 32 `n"
    month_line_1:=" January `n February `n March `n April `n May `n June `n July `n August `n September `n October `n November `n December `n"
    month_line_2:=" Jan `n Feb `n Mar `n Apr `n May `n June `n July `n Aug `n Sep `n Oct `n Nov `n Dec `n"
    year_Line:=" 2021 `n 2022 `n 2023 `n 2024 `n 2025 `n 2026 `n 2027 `n 2028 `n 2029 `n 2030 `n"
    try {
        ; 直接获取百度HTTP响应头中的Date字段
        cmd := "`"" curlPath "`"" " -sI " "-m 5 " "`"https://www.baidu.com`"" " > " A_ScriptDir "\baidu_time.txt"
        RunWait('cmd /c "' cmd '"', , "Hide")
        if FileExist(A_ScriptDir "\baidu_time.txt") {
            response := FileRead(A_ScriptDir "\baidu_time.txt")
            response_Array := StrSplit(response, "`n", "`r")
            for line in response_Array {
                if (InStr(line, "Date: ") = 1) {
                    line:=StrReplace(line,"Date: ","")
                    Start:=InStr(line,",")
                    end:=InStr(line,":")-2
                    baiduTime := SubStr(line, Start+1, end-Start-1)
                    curllog_Path := StrReplace(curlPath, "curl.exe", "curl_log.txt")
                    FileAppend("[" A_Now "] 百度时间: " baiduTime "`n", curllog_Path)
                    break
                }
            }
            FileDelete(A_ScriptDir "\baidu_time.txt")  ; 删除临时文件
        }

    }
    catch {
        baiduTime := ""
    }    
    time := baiduTime
    Time_Line := year_Line month_line_1 month_line_2 day_Line
    loop
    {
        Time_Line := StrReplace(Time_Line, "`r`n`r`n", "`r`n", , &count)
        if (count = 0)
        {
            Time_Line := Trim(Time_Line,"`r`n")
            break
        }
    }

    ; MsgBox("获取到的时间字符串: " time, "获取时间", "T1")
    time_Array := StrSplit(Trim(time), " ")

    ; MsgBox("获取到的时间字符串: " Time_Line, "获取时间", "T1")
    Time_Line_Array := StrSplit(Time_Line, "`n","`r")

    for line in Time_Line_Array
    {
        if (InStr(day_Line, line)>0 && StrLen(Trim(line))=2)
        {
            line := Trim(line)
            for time in time_Array
            {
                if (InStr(line, time)>0)
                {
                    Number_Day := " " line " "
                    break
                }
            }
        }
    }

    for line in Time_Line_Array
    {
        if (InStr(month_Line_1, line)>0 or InStr(month_Line_2, line)>0)
        {
            line := Trim(line)
            for time in time_Array
            {
                if (InStr(line, time)>0)
                {
                    Number_Month := " " line " "
                    break
                }
            }
        }
    }

    for line in Time_Line_Array
    {
        if (InStr(year_Line, line)>0 && StrLen(Trim(line))=4)
        {
            line := Trim(line)
            for time in time_Array
            {
                if (InStr(line, time)>0)
                {
                    Number_Year := " " line " "
                    break
                }
            }
        }
    }

    Time_Line:= Number_Year "年" Number_Month "月" Number_Day "日"
    month_Index_line:=" 01 `n 02 `n 03 `n 04 `n 05 `n 06 `n 07 `n 08 `n 09 `n 10 `n 11 `n 12 "
    month_Index_Array := StrSplit(month_Index_line, "`n", "`r")
    Month_Summary := month_line_1 month_line_2
    loop
    {
        Month_Summary := StrReplace(Month_Summary, "`r`n`r`n", "`r`n", , &count)
        if (count = 0)
        {
            Month_Summary := Trim(Month_Summary,"`r`n")
            break
        }
    }
    Month_Summary_Array := StrSplit(Month_Summary, "`n", "`r")
    ; MsgBox("解析后的时间为: " Time_Line, "获取时间", "T1")
    loop Month_Summary_Array.Length
    {
        this_month := Trim(Month_Summary_Array[A_Index])
        if (this_month = "")
        {
            continue
        }
        if (InStr(Time_Line, Trim(this_month))>0)
        {
            ; MsgBox("找到了月份: " this_month, "获取时间", "T1")
            if(Mod(A_Index,month_Index_Array.Length)=0)
            {
                ; MsgBox("找到了最后一个月", "获取时间", "T1")
                Index_Month := month_Index_Array[month_Index_Array.Length]
            }
            else
            {
                Index_Month := month_Index_Array[Mod(A_Index,month_Index_Array.Length)]
            }
        }
    }

    ; Time_Line:= Number_Year "年" Index_Month "月" Number_Day "日"
    Time_Line:= Number_Year  Index_Month  Number_Day
    Time_Line:= StrReplace(Time_Line," ","")
    ids:=WinGetList("ahk_exe cmd.exe")
    for this_id in ids
    {
        WinClose(this_id)
        Sleep(100)
    }
    return Time_Line
}

GetFolder_Path(File_Path) {
    SplitPath(File_Path, , &dir)
    return dir
}

StrReplaceAll(Content, OldString, NewString) {
    Loop
    {
        Content := StrReplace(Content, OldString, NewString, , &Count)
        if (Count = 0)  ; 不再需要更多的替换.
            break
    }
    return Content
}

/*
    Save_String_To_StrReplaceAll
    用途：将 Content_One 中所有出现在 Content_Saved_List 的字符串替换为加前后缀的新字符串。
    参数：
        Content_One (String)         - 原始内容字符串
        Content_Saved_List (String)  - 以换行分隔的待替换字符串列表
        OldString (String)           - 未使用，仅为兼容参数
        NewString (String)           - 未使用，仅为兼容参数
        Save_Prefix (String)         - 替换时添加的前缀，默认 "【"
        Save_Suffix (String)         - 替换时添加的后缀，默认 "】"
    返回值：
        (String) 替换后的内容字符串

Save_String_To_StrReplaceAll( Content_One, Content_Saved_List, OldString, NewString, Save_Prefix := "{[", Save_Suffix := "]}")
{
    Content_One := StrReplaceAll(Content_One, Save_Prefix, "`t")
    Content_One := StrReplaceAll(Content_One, Save_Suffix, "`t")
    Content_Saved_List := StrSplit(Content_Saved_List, "`n", "`r")
    loop Content_Saved_List.Length
    {
        if (Content_Saved_List[A_Index] = "") {
            continue
        } else if (InStr(Content_One, Content_Saved_List[A_Index])>0 && InStr(Content_One, Save_Prefix Content_Saved_List[A_Index] Save_Suffix)=0) {
            Content_One := StrReplace(Content_One, Content_Saved_List[A_Index], Save_Prefix Content_Saved_List[A_Index] Save_Suffix)
        } else {
            continue
        }
    }
    return Content_One
}

Data_Save_Special_Char(Content_One) {
    ; 保存特殊字符
    Content_One := Save_String_To_StrReplaceAll( Content_One, Content_Saved_List, OldString, NewString, Save_Prefix := "{[", Save_Suffix := "]}")
    Content_One := StrReplaceAll(Content_One, OldString := Save_Prefix, NewString := "")
    Content_One := StrReplaceAll(Content_One, OldString := Save_Suffix, NewString := "")

    Content_One := StrReplace(Content_One, "【单选题】", "`r`n【单选题】") 
    Content_One := StrReplace(Content_One, "【多选题】", "`r`n【多选题】") 
    Content_One := StrReplace(Content_One, "【判断题】", "`r`n【判断题】") 
    Content_One := StrReplace(Content_One, "\", "`r`n") 
    Content_One := Trim(Content_One, "`r`n")
}

*/

 