if (! (Test-Path ./custom-reference.docx)) {
    # https://pandoc.org/MANUAL.html#options-affecting-specific-writers
    pandoc -o custom-reference.docx --print-default-data-file reference.docx
}

# Set-Location "${HOME}\source\repos\zoyo"

try {
    $word = [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
    # wordドキュメントを閉じる
    $word.Documents.Close($false)
    # wordを終了
    $word.Quit()
} catch {
    # throw "Wordが起動していません"
}

# メモリ処理
[gc]::Collect()
[gc]::WaitForPendingFinalizers()
try {
    [Runtime.Interopservices.Marshal]::ReleaseComObject($word) >$null
} catch {}
# 変数除去
if (Test-Path Variable:\word) {
    Remove-Variable word
}

# wordファイルのパス
$wordName = 'zoyo.docx'
$wordFile = Get-Item $wordName
$wordFilePath = $wordFile.FullName

pandoc --reference-doc custom-reference.docx -o $wordFilePath zoyo.md

# pandoc -s zoyo.md -o zoyo.html
# pandoc --reference-doc custom-reference.docx -o $wordFilePath.Replace('.docx', '-html.docx') zoyo.html

# wordオブジェクト生成
$word = New-Object -ComObject word.Application
while ($null -eq $word) {
    Start-Sleep -Seconds 1.0
}

# 表示
$word.Visible = $true
# wordドキュメントを開く
$word.Documents.Open($wordFilePath) >$null
# $word.Documents.Open($wordFilePath.Replace('.docx', '-html.docx')) >$null
