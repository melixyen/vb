Module ModLangu
    Public NVTest(19) As String
    Public DotCM(40) As String '標點符號
    'Public RecogLeft() As String = {"「", "『", "〝", "【", "《", "（", "‵", Chr(-24153)}
    'Public RecogRight() As String = {"」", "』", "〞", "】", "》", "）", "′", Chr(-24152)}
    Public RecogLeft() As String = {"「", "『", "【", "《", "（", Chr(-24153)}
    Public RecogRight() As String = {"」", "』", "】", "》", "）", Chr(-24152)}

    '單字區
    Public Lang_Total = "總共"
    Public Lang_At = "在"
    Public Lang_Area1 = "地方"
    Public Lang_Ga = "個"
    Public Lang_Finded = "找到"
    Public Lang_And1 = "和"
    Public Lang_And2 = "並且"
    Public Lang_And3 = "及"
    Public Lang_Replace = "取代"

    'Message 區
    Public Msg_DoYouWantSave As String = "這篇小說尚未存檔，是否要儲存呢？"
    Public Msg_YouHadOpenThisNovel As String = "您已經開啟這篇小說了"
    Public Msg_SearchAtLast As String = "找不到下一個了，要不要從頭找起呢？"
    Public Msg_InputJeTell As String = "請輸入您要設定給本分節的標題或名稱，人物設定請輸入您想顯示在列表的名字，如果你想中止請不要輸入任何字" & vbCrLf & vbCrLf & "如果輸入法變成灰色框框無法打中文，請按一次Ｓｈｉｆｔ＋Ａ"
    Public Msg_InputJeEnd1 As String = "我沒看到你輸入的分節名稱，中止新增囉"
    Public Msg_DeleteLsJe1 As String = "你指定要刪除的標記無法在文章中找到喔，請重整後再刪除"
    Public Msg_DeleteLsJe2 As String = "選取區的分節標記是不是你想要刪除的呢？"
    Public Msg_JumpLsJe1 As String = "嗯，我找不到您所選擇的這個分節標記呢"
    Public Msg_OpenTxtBak1 As String = "嗯，我偵測到這個檔案有一個備份檔(.nbk)存在，可能是前次未正常關閉程式造成的，現在您必需做一個選擇" & vbCrLf & vbCrLf & "選擇是(YES)，我們將載入備份檔的內容，它可能是比較接近你原先打到一半的文件" & vbCrLf & vbCrLf & "選擇否(NO)，我們將載入您的原始檔案，它可能停留在您最後一次存檔的內容"
    Public Msg_InputShuQian1 As String = "請問您新增的書籤要用什麼名字？"
    Public Msg_BSExpertTitleDo As String = "是否要加上完整 HTML 檔頭<html><body>等部分？" & vbCrLf & vbCrLf & "不加的話僅會輸出人物設定表格語法"
    Public Msg_DeleteArtJe As String = "你確定要刪除掉這個分節所有內容及標籤嗎？" & vbCrLf & vbCrLf & "如果要快速刪除可以在點選分節列表的項目後按下鍵盤的Delete鍵直接刪掉"

    'ToolTip區
    Public FormEdit_BtnJeNewTP As String = "會在您的游標所在位置插入一個分節段落記號"
    Public FormEdit_BtnJeDelTP As String = "將刪除你在分節列表上所選擇的該分節段落記號"
    Public FormEdit_BtnJeRefreshTP As String = "如果你在文章內做了更動，請按此鈕更新分節列表"
    Public FormEdit_BtnJeSelTP As String = "按下我可以把你選擇的分節段落區域文章選取起來"
    Public FormEdit_LsJeJumpTP As String = "點兩下可以幫助你跳躍到你所選擇的分節段落去" & vbCrLf & vbCrLf & "點選後再按中鍵(滾輪鍵)可選取整段分節標籤與內文" & vbCrLf & vbCrLf & "點選後按鍵盤的 Delete 鍵可以將整段分節內容(含分節標籤)刪掉"
    Public FormEdit_HSTabPGTP As String = "調整文件編輯區和輔助板所分配的畫面比例"
    Public FormEdit_LongSelectToolStripButtonTP As String = "於目前游標位置按下後點到文件編輯框任何一位置並再點一次此按鈕就會將這一段圈選起來"
    Public FormEdit_FormatAllToolStripButtonTP As String = "沒有選取的話會將文章的所有字型和色彩統一為目前系統設定的樣式" & vbCrLf & "有選取的話只會將選取的部分字型和色彩重定格式為系統設定的樣式"
    Public FormEdit_RecognitionToolStripButtonTP As String = "按此幫助您把目前文章使用引號部分標上顏色醒目提示"
    Public FormEdit_BtnPanelCountWordHide As String = "打開或關閉字數計算，需注意當打開時使用傳統中文輸入法會跳過相關字詞選單"
    Public FormEdit_LbTWCTP As String = "去除空白和換行字元後的總字數"
    Public FormEdit_LbNWCTP As String = "去除空白和換行字元以及分節標籤後的總字數"
    Public FormEdit_LbSWCTP As String = "選取範圍去除空白和換行字元後的總字數"

    Public FormSet_CkAutoBakTP As String = "勾選我的話，有存檔位置的檔案每" & AutoBakSec.ToString & "秒在適當時機會備份一次到該檔案資料夾下為該檔名.nbk檔"
    Public FormSet_CkSTimeTP As String = "勾選我的話會在左側字數計顯示您啟動本軟體後已經打文打了多久的時間"
    Public FormSet_BtnBSBK1LTP As String = "將備用項目1的人物設定項目格式取代掉目前使用的格式"
    Public FormSet_BtnBSBK1CTP As String = "交換，讓備用項目1的格式成為使用的格式，然後把目前使用的格式存進備用項目1"
    Public FormSet_BtnBSBK1RTP As String = "將目前使用的人物設定項目格式存進備用項目1"
    Public FormSet_BtnBSBK2LTP As String = "將備用項目2的人物設定項目格式取代掉目前使用的格式"
    Public FormSet_BtnBSBK2CTP As String = "交換，讓備用項目2的格式成為使用的格式，然後把目前使用的格式存進備用項目2"
    Public FormSet_BtnBSBK2RTP As String = "將目前使用的人物設定項目格式存進備用項目2"
    Public FormSet_BtnBSBK3LTP As String = "將備用項目3的人物設定項目格式取代掉目前使用的格式"
    Public FormSet_BtnBSBK3CTP As String = "交換，讓備用項目3的格式成為使用的格式，然後把目前使用的格式存進備用項目3"
    Public FormSet_BtnBSBK3RTP As String = "將目前使用的人物設定項目格式存進備用項目3"




    Public Sub ADDTxTe()
        NVTest(0) = "，等於的Label名稱是labEqu，清除的Label名稱是labClear其他單純顯示訊息用的Label我將名稱改為Not開始的名稱，例如Notlab；，在Form_Load中利用addhandler將相關想要處理的事件加到處理常式"
        Dim i As Integer
        For i = 0 To 18
            'NVTest(i).Text = i.ToString & NVTest(0).Text & vbCrLf & vbCrLf & NVTest(0).Text & vbCrLf & vbCrLf & NVTest(0).Text & vbCrLf & vbCrLf & NVTest(0).Text & vbCrLf & vbCrLf & NVTest(0).Text
            NVTest(i) = CStr(i)
        Next
    End Sub

    Public Sub LoadLangu()
        '針對語言要載入的項目在此設定後統一載入
        LoadLangu_DotCM()
    End Sub

    Public Sub LoadLangu_DotCM()
        DotCM(0) = "，" : DotCM(1) = "。" : DotCM(2) = "、" : DotCM(3) = "．"
        DotCM(4) = "；" : DotCM(5) = "㊣" : DotCM(6) = "：" : DotCM(7) = "？"
        DotCM(8) = "！" : DotCM(9) = "…" : DotCM(10) = "／" : DotCM(11) = "‥"
        DotCM(12) = "∞" : DotCM(13) = "–" : DotCM(14) = "＋" : DotCM(15) = "－"
        DotCM(16) = "×" : DotCM(17) = "÷" : DotCM(18) = "±" : DotCM(19) = "√"
        DotCM(20) = "≦" : DotCM(21) = "≧" : DotCM(22) = "≠" : DotCM(23) = "※"
        DotCM(24) = "∫" : DotCM(25) = "﹏" : DotCM(26) = "※" : DotCM(27) = "◎"
        DotCM(28) = "●" : DotCM(29) = "★" : DotCM(30) = "◆" : DotCM(31) = "■"
        DotCM(32) = "【】" : DotCM(33) = "《》" : DotCM(34) = "「」" : DotCM(35) = "『』"
        DotCM(36) = "〝〞" : DotCM(37) = "‵′" : DotCM(38) = "（）" : DotCM(39) = Chr(-24153) & Chr(-24152)
    End Sub

    Public Function Make_arBtn_TP(ByVal j As Integer) As String
        Make_arBtn_TP = ""
        If j = 0 Then Make_arBtn_TP = "Ctrl + ,"
        If j = 1 Then Make_arBtn_TP = "Ctrl + ."
        If j = 4 Then Make_arBtn_TP = "Ctrl + ;"
        If j = 39 Then Make_arBtn_TP = "Ctrl + """
        If j = 34 Then Make_arBtn_TP = "Ctrl + Q"
        If j = 35 Then Make_arBtn_TP = "Ctrl + W"
        If j = 2 Then Make_arBtn_TP = "Ctrl + 3"
        If j = 9 Then Make_arBtn_TP = "Ctrl + 4"
    End Function
End Module
