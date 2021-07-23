#
# スキャナ読取A4画像の処理スクリプト
#
# パス変数定義
$MAGICK = D:\ImageMagick-7.0.10-22-portable-Q16-x64
# 画面クリア
Clear-Host
# magick.exe が Dドライブ直下にあるか確認する
if (!(Test-Path -path D:\ImageMagick-7.0.10-22-portable-Q16-x64)) {
    # magick.exe が Dドライブ直下にない
    # 鍵の共有にある ImageMagick の zip 圧縮ファイルを Dドライブ直下に解凍する
    Expand-Archive -Path \\Our\Serveris\here\共有\Tools\ImageMagick-7.0.10-22-portable-Q16-x64.zip -DestinationPath d:\ImageMagick-7.0.10-22-portable-Q16-x64
}
# 初期処理終了
#
# 画像処理開始
# アセンブリの参照設定を追加する
[void][reflection.assembly]::LoadWithPartialName("Microsoft.VisualBasic")

# スクリプトが登録されているフォルダのパスを取得する
$path = Split-Path $MyInvocation.MyCommand.Path -Parent
# write-host $path
#######################################
# ログインユーザーのダウンロード・フォルダ・パスを取得する
[String] $USR_DEKSTOP_PATH   = $env:USERPROFILE + "\Downloads"
# 画像JPGファイル選択ダイアログ作成
#Param([Parameter()] [String] $FilePath )
# $FilePath が設定されていない、又はファイルが存在しないとき
if([string]::IsNullOrEmpty($FilePath) -Or (Test-Path -LiteralPath $FilePath -PathType Leaf) -eq $false) {
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "JPG ファイル(*.jpg)|*.jpg"
    $dialog.InitialDirectory = $USR_DEKSTOP_PATH
    $dialog.Title = "スキャナで読み込んだ画像ファイルを選択してください"
    # キャンセルを押された時は処理を止める
    if($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::NG){
        exit 1
    }
    # 選択したファイルパスを $FilePath に設定
    $FilePath = $dialog.FileName
}
#
# $FilePath処理する
Write-host ">> " $FilePath
# D ドライブ直下の ImageMagick で画像をトリミングする　画像周囲の余白をカット
d:\ImageMagick-7.0.10-22-portable-Q16-x64\magick.exe $FilePath -crop 4614x6670+170+170 $FilePath".trimed.jpg"
# D ドライブ直下の ImageMagick で画像をリサイズする　94%縮小
d:\ImageMagick-7.0.10-22-portable-Q16-x64\magick.exe $FilePath".trimed.jpg" -resize 94%^ -quality 100 $FilePath".shrinked.jpg"
#
# 完了処理　不要なjpgを削除
Remove-Item $FilePath".trimed.jpg"
#
# Why 94% shlinking at line 48 using ImageMagick?
# Because for fit its actual size.
# For example, draw a line on powerpoint document and print it out.
# Measure the length of the line with a ruler.        --> line length (A)
# Scan printed out paper.
# Paste scaned image file on powerpoint document and print it out again.
# Again Measure the length of the line with a ruler.  --> line length (B)
# Calcurate (A)/(B)*100. The answer is scale ratio/factor.  --> shlink/expand ratio(C)
# The length of (B) varies depending on the image processing of the scanner machine you used.
# With imagick.exe, Manipurate scanned image file at that scaling ratio/factor and paste it on powerpoint document again.
# Print it out again and measure the length of the line with a ruler and repeat till both lines matches as same length.
# Determine the scaling ratio so that it is approximately the same length.
#
