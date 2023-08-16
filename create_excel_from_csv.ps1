param(
    [string]$csv_dir = (Split-Path $MyInvocation.MyCommand.Path),
    [string]$path = (Join-Path (Get-Location).Path "new_excel"),
    [switch]$skip_empty_file,
    [string]$encoding = "shift_jis"
)

function OK($msg){
    "`rOK " | Write-Host -ForegroundColor Green -NoNewLine
    $msg | Write-Host
}

function Skip($msg){
    "`rSkip " | Write-Host -ForegroundColor Yellow -NoNewLine
    $msg | Write-Host
}

if(! (Get-Module -ListAvailable ImportExcel)){
    "ImportExcel module not found, install ImportExcel (Install-Module ImportExcel)" | Write-Host -ForegroundColor Red
    return
}
if (! (Test-Path $csv_dir)) {
    "CSV path {0} not found" -F $csv_dir
    return
}

try {
    if((Get-Item $path -ErrorAction "Stop").Attributes -eq "Directory"){
        "Output path {0} is directory, file is expected" -F $path
        return
    }
} catch {
    $tmp = (Join-Path (Get-Location).Path $path) -Split "/"
    $base = $tmp[0..($tmp.Length-2)] -Join "/"
    if(! (Test-Path $base)){
        "Output path base {0} not found" -F $base
        return
    }
}
if(! $path.EndsWith(".xlsx")){
    $path += ".xlsx"
}

$style = New-ExcelStyle -FontName "メイリオ" -FontSize 12
Get-ChildItem -Path $csv_dir -File | where Extension -eq ".csv" | %{
    $csv = $_

    $msg = "{0} -> Sheet[{1}]" -F $csv.Name, $csv.BaseName
    $msg | Write-Host -NoNewLine

    $tmp = Import-Csv $csv.FullName -Encoding shift_jis
    if($tmp.Length -gt 0) {
        $tmp | Export-Excel -Path $path -WorksheetName $csv.BaseName -Style $style
        OK($msg)
    }else{
        if($skip_empty_file){
            Skip($msg)
        }else{
            $tmp | Export-Excel -Path $path -WorksheetName $csv.BaseName
            OK($msg)
        }
    }
}

"{0} に出力しました" -F $path | Write-Host -ForegroundColor Cyan
