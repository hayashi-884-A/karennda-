# === 設定 ===
$xlsxPath = "C:\Users\kosuk\Desktop\線若干うまくいってくる メモ 親うまくいいって、色ついてる２.xlsm"   # ← フルパスに変更
$macro    = "'{0}'!Automation_ShowForm" -f (Split-Path $xlsxPath -Leaf)

# === Excel を非表示で起動 ===
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.ScreenUpdating = $false

# ここから例外が出ても Excel を確実に終了させる
$wb = $null
try {
    $wb = $xl.Workbooks.Open($xlsxPath)

    # UserForm を出すマクロを実行（フォームを閉じるまでここで待機）
    $xl.Run($macro)

    # 入力結果をそのまま上書き保存
    $wb.Save()
}
finally {
    if ($wb) { $wb.Close($true) }
    $xl.Quit()

    # COM 解放（Excel が残らないように）
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
