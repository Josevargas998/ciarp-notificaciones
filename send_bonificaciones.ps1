# Script to send CIARP bonificaciones (acta 4- 02/10/2026) notifications via Gmail REST API
param (
    [switch]$Test,
    [switch]$TestAll,
    [switch]$Real,
    [switch]$Confirmar
)

$ciarpPath    = Join-Path $PSScriptRoot "bonificaciones2026.xlsx"
$correosPath  = Join-Path $PSScriptRoot "correos.xlsx"
$tokensPath   = Join-Path $PSScriptRoot "tokens.json"
$credsPath    = Join-Path $PSScriptRoot "credentials.json"
$actaRegex    = '^\s*4\s*-\s*02/10/2026'

$char_i_accent = [char]0xed
$char_a_accent = [char]0xe1
$char_o_accent = [char]0xf3
$char_e_accent = [char]0xe9

$notaComisionDefault = "Los puntos ser" + $char_a_accent + "n pagados al t" + $char_e_accent + "rmino de la comisi" + $char_o_accent + "n acad" + $char_e_accent + "mica-administrativa."
$subjectText = "Respuesta a su solicitud de reconocimiento de bonificaciones"

if (-not $Test -and -not $TestAll -and -not $Real) {
    Write-Host "Debes especificar -Test, -TestAll o -Real." -ForegroundColor Yellow
    Write-Host "  -Test    : Envia 1 correo de prueba (primer docente) a jhvargas"
    Write-Host "  -TestAll : Envia TODOS los correos a jhvargas (para revision)"
    Write-Host "  -Real    : Envia correos reales a cada docente"
    exit
}

function Get-AccessToken {
    if (-not (Test-Path $credsPath)) { Write-Error "No se encuentra credentials.json."; return $null }
    $creds = Get-Content $credsPath -Raw | ConvertFrom-Json
    $clientId = $creds.installed.client_id
    $clientSecret = $creds.installed.client_secret
    $tokens = $null
    if (Test-Path $tokensPath) { $tokens = Get-Content $tokensPath -Raw | ConvertFrom-Json }

    if (-not $tokens -or -not $tokens.refresh_token) {
        $port = 8080; $redirectUri = "http://localhost:$port/"
        $scope = "https://www.googleapis.com/auth/gmail.send"
        $authUrl = "https://accounts.google.com/o/oauth2/auth?client_id=$([uri]::EscapeDataString($clientId))&redirect_uri=$([uri]::EscapeDataString($redirectUri))&response_type=code&scope=$([uri]::EscapeDataString($scope))&access_type=offline&prompt=consent"
        $listener = New-Object System.Net.HttpListener
        $listener.Prefixes.Add($redirectUri)
        try { $listener.Start() } catch { Write-Error "No se pudo iniciar servidor local."; return $null }
        Start-Process $authUrl
        $context = $listener.GetContext()
        $code = $context.Request.QueryString["code"]
        $buf = [System.Text.Encoding]::UTF8.GetBytes("<html><body>OK. Puedes cerrar esta pestana.</body></html>")
        $context.Response.ContentLength64 = $buf.Length
        $context.Response.OutputStream.Write($buf, 0, $buf.Length)
        $context.Response.OutputStream.Close()
        $listener.Stop()
        if (-not $code) { Write-Error "Sin codigo de autorizacion."; return $null }
        $tokenRes = Invoke-RestMethod -Uri "https://oauth2.googleapis.com/token" -Method Post -Body @{
            code=$code; client_id=$clientId; client_secret=$clientSecret; redirect_uri=$redirectUri; grant_type="authorization_code"
        }
        $tokens = @{ access_token=$tokenRes.access_token; refresh_token=$tokenRes.refresh_token; expiry_time=(Get-Date).AddSeconds($tokenRes.expires_in).ToString("o") }
        $tokens | ConvertTo-Json | Out-File $tokensPath -Encoding utf8
    } else {
        if ((Get-Date).AddMinutes(5) -ge [DateTime]$tokens.expiry_time) {
            try {
                $tokenRes = Invoke-RestMethod -Uri "https://oauth2.googleapis.com/token" -Method Post -Body @{
                    client_id=$clientId; client_secret=$clientSecret; refresh_token=$tokens.refresh_token; grant_type="refresh_token"
                }
                $tokens.access_token = $tokenRes.access_token
                $tokens.expiry_time = (Get-Date).AddSeconds($tokenRes.expires_in).ToString("o")
                $tokens | ConvertTo-Json | Out-File $tokensPath -Encoding utf8
            } catch { Remove-Item $tokensPath -ErrorAction SilentlyContinue; return Get-AccessToken }
        }
    }
    return $tokens.access_token
}

function Send-GmailMessage {
    param([string]$accessToken,[string]$to,[string]$subject,[string]$htmlBody,[string]$bcc="")
    $fromName = "CIARP Uniquindio"
    $from = "$fromName <jhvargas@uniquindio.edu.co>"
    $bccH = if ($bcc) { "`r`nBcc: $bcc" } else { "" }
    $mime = "From: $from`r`nTo: $to$bccH`r`nSubject: $subject`r`nMIME-Version: 1.0`r`nContent-Type: text/html; charset=utf-8`r`n`r`n$htmlBody"
    $b64url = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($mime)).Replace('+','-').Replace('/','_').Replace('=','')
    try {
        Invoke-RestMethod -Uri "https://gmail.googleapis.com/gmail/v1/users/me/messages/send" -Method Post `
            -Headers @{Authorization="Bearer $accessToken";"Content-Type"="application/json"} `
            -Body (@{raw=$b64url} | ConvertTo-Json) | Out-Null
        return $true
    } catch {
        Write-Error "Error al enviar a $to : $_"
        if ($_.Exception.Response) { $reader=New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream()); Write-Error $reader.ReadToEnd() }
        return $false
    }
}

function Build-EmailHtml ($fullName,$programa,$facultad,$firstActa,$rowsHtml,$sumStr,$notaComision="",$productCount=1) {
    $notaHtml = ""
    $parrafoActo = "<p style='margin-top:20px;color:#2d3748;font-size:15px;text-align:justify'>Por lo anteriormente mencionado, se le estar" + $char_a_accent + " notificando posteriormente del correspondiente acto administrativo.</p>"
    if ($notaComision) {
        $notaHtml = "<div style='background-color:#e3f2fd;border-left:4px solid #1565c0;padding:14px 16px;margin-top:20px;border-radius:0 4px 4px 0;font-size:14px;color:#0d47a1'><strong>Nota sobre comisi&oacute;n acad&eacute;mica-administrativa:</strong><br>$notaComision</div>"
        $parrafoActo = ""
    }
    $actaSessionRef = "<strong>$firstActa</strong>"
    $actaFechaText = ""
    $monthNames = @("","enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre")
    if ($firstActa -match "No\.\s*(\d+)\s+del\s+(\d+\s+de\s+\w+)") {
        $actaSessionRef = "sesi&oacute;n <strong>No. $($matches[1])</strong>"
        $actaFechaText = ", llevada a cabo durante el d&iacute;a $($matches[2]) de la presente anualidad"
    } elseif ($firstActa -match "^(\d+)\s*-\s*(\d{2})/(\d{2})/(\d{4})") {
        $actaSessionRef = "sesi&oacute;n <strong>No. $($matches[1])</strong>"
        $actaFechaText = ", llevada a cabo durante el d&iacute;a $($matches[2]) de $($monthNames[[int]$matches[3]]) de la presente anualidad"
    }
    $introSolicitud = if ($productCount -eq 1) {
        "la solicitud de bonificaci&oacute;n, presentada por usted, fue considerada en el mismo y aprobada como se relaciona a continuaci&oacute;n:"
    } else {
        "las solicitudes de bonificaci&oacute;n, presentadas por usted, fueron consideradas en el mismo y aprobadas como se relaciona a continuaci&oacute;n:"
    }
    $introParrafo = "Dando cumplimiento a la directriz emitida por la jefe de la Oficina de Asuntos Profesorales, me permito informarle que, en $actaSessionRef del Comit&eacute; Interno de Asignaci&oacute;n y Reconocimiento de Puntaje (CIARP)$actaFechaText, $introSolicitud"

    return "<html><head><meta charset='utf-8'><style>body{font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;background-color:#f4f6f8;margin:0;padding:20px}.card{max-width:650px;margin:0 auto;background:#fff;border:1px solid #e0e0e0;border-radius:8px;overflow:hidden;box-shadow:0 4px 10px rgba(0,0,0,.05)}.header{background-color:#1b5e20;color:#fff;padding:25px 20px;text-align:center}.header h1{margin:0;font-size:26px;font-weight:700;letter-spacing:1px}.header p{margin:6px 0 0;font-size:13px;opacity:.9}.content{padding:30px;line-height:1.7}.docente-info{background:#f8fafc;border-left:4px solid #1b5e20;padding:15px;margin-bottom:25px;border-radius:0 4px 4px 0}.docente-info p{margin:0;font-size:14px;color:#4a5568}.docente-info .name{font-size:16px;font-weight:700;color:#1b5e20;text-transform:uppercase;margin-bottom:4px}.table-container{margin:20px 0;overflow-x:auto}table{width:100%;border-collapse:collapse;border:1px solid #e0e0e0}th{background-color:#1b5e20;color:#fff;padding:12px;font-size:14px;text-align:left}.total-box{background-color:#e8f5e9;border:1px solid #c8e6c9;padding:12px 15px;border-radius:4px;font-size:15px;font-weight:700;color:#2e7d32;margin-top:15px}.signature{margin-top:30px;font-size:14px;color:#4a5568;line-height:1.4;border-top:1px solid #e2e8f0;padding-top:20px}</style></head><body><div class='card'><div class='header'><h1>CIARP</h1><p>Comit&eacute; Interno de Asignaci&oacute;n y Reconocimiento de Puntaje</p></div><div class='content'><h2 style='color:#1b5e20;text-align:center;margin-top:0;margin-bottom:25px;font-size:20px'>Reconocimiento de Bonificaciones</h2><div class='docente-info'><p class='name'>$fullName</p><p><strong>Programa:</strong> $programa</p><p><strong>Facultad:</strong> $facultad</p></div><p style='font-size:15px;font-weight:700;margin-bottom:15px;color:#2d3748'>Cordial saludo,</p><p style='color:#2d3748;font-size:15px;margin-bottom:20px;text-align:justify;line-height:1.7'>$introParrafo</p><div class='table-container'><table><thead><tr><th style='width:20%'>Tipo</th><th style='width:55%'>T&iacute;tulo de Producto Acad&eacute;mico</th><th style='width:25%;text-align:right'>Puntos</th></tr></thead><tbody>$rowsHtml</tbody></table></div><div class='total-box'>Total de bonificaciones aprobadas en esta sesi&oacute;n: $sumStr puntos</div>$notaHtml$parrafoActo<p style='color:#2d3748;font-size:15px;text-align:justify'>Cualquier inquietud al respecto, estaremos atentos.</p><div class='signature'>Atentamente,<br><br><div style='background-color:#f1f8e9;border:1px solid #dcedc8;padding:15px;border-radius:4px;text-align:center;color:#1b5e20;font-weight:700;font-size:16px'>Luz Amparo Celis Buritic&aacute;<br><span style='font-size:14px;font-weight:400;color:#2e7d32'>Jefe Oficina de Asuntos Profesorales</span></div><p style='margin-top:12px;font-size:13px;color:#718096'>Proyect&oacute;: Jos&eacute; Heriberto Vargas Espinosa</p></div></div></div></body></html>"
}

function Build-RowsHtml ($products) {
    $rows = ""
    foreach ($p in $products) {
        $rows += "<tr style='border-bottom:1px solid #e0e0e0'><td style='padding:12px;font-size:14px;color:#333;font-weight:700'>$($p.Concepto)</td><td style='padding:12px;font-size:14px;color:#555'>$($p.Detalle)</td><td style='padding:12px;font-size:14px;color:#333;text-align:right;font-weight:700'>$($p.Puntaje) Puntos</td></tr>"
    }
    return $rows
}

function Get-ProductSum ($products) {
    $sum = 0.0
    foreach ($p in $products) {
        $v=0.0
        if ([double]::TryParse(($p.Puntaje -replace ',','.'), [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$v)) { $sum+=$v }
    }
    return $sum
}

function Generate-Report ($reportRows, $reportPath) {
    Write-Output "Generando reporte Excel..."
    $xl = New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
    $wb = $xl.Workbooks.Add()
    $sh1 = $wb.Sheets.Item(1); $sh1.Name = "Notificados"
    $h1 = @("DNI","Nombre Docente","Correo","Programa","Facultad","Tipo","Detalle","Puntaje","Acta")
    for ($c=0;$c -lt $h1.Count;$c++) {
        $cell=$sh1.Cells.Item(1,$c+1); $cell.Value2=$h1[$c]; $cell.Font.Bold=$true
        $cell.Interior.Color=0x1e7a1e; $cell.Font.Color=0xFFFFFF
    }
    $row1=2
    foreach ($r in ($reportRows | Where-Object { $_.Estado -like "Notificado*" })) {
        $sh1.Cells.Item($row1,1).Value2=$r.Dni; $sh1.Cells.Item($row1,2).Value2=$r.Nombre
        $sh1.Cells.Item($row1,3).Value2=$r.Correo; $sh1.Cells.Item($row1,4).Value2=$r.Programa
        $sh1.Cells.Item($row1,5).Value2=$r.Facultad; $sh1.Cells.Item($row1,6).Value2=$r.Concepto
        $sh1.Cells.Item($row1,7).Value2=$r.Detalle; $sh1.Cells.Item($row1,8).Value2=$r.Puntaje
        $sh1.Cells.Item($row1,9).Value2=$r.Acta
        if ($row1 % 2 -eq 0) { $sh1.Rows.Item($row1).Interior.Color=0xE8F5E9 }
        $row1++
    }
    $sh1.Columns.AutoFit() | Out-Null
    $sh2=$wb.Sheets.Add(); $sh2.Name="No Notificados"
    $h2=@("DNI","Nombre Docente","Correo","Programa","Facultad","Tipo","Detalle","Puntaje","Motivo")
    for ($c=0;$c -lt $h2.Count;$c++) {
        $cell=$sh2.Cells.Item(1,$c+1); $cell.Value2=$h2[$c]; $cell.Font.Bold=$true
        $cell.Interior.Color=0x7a1e1e; $cell.Font.Color=0xFFFFFF
    }
    $row2=2
    foreach ($r in ($reportRows | Where-Object { $_.Estado -notlike "Notificado*" })) {
        $sh2.Cells.Item($row2,1).Value2=$r.Dni; $sh2.Cells.Item($row2,2).Value2=$r.Nombre
        $sh2.Cells.Item($row2,3).Value2=$r.Correo; $sh2.Cells.Item($row2,4).Value2=$r.Programa
        $sh2.Cells.Item($row2,5).Value2=$r.Facultad; $sh2.Cells.Item($row2,6).Value2=$r.Concepto
        $sh2.Cells.Item($row2,7).Value2=$r.Detalle; $sh2.Cells.Item($row2,8).Value2=$r.Puntaje
        $sh2.Cells.Item($row2,9).Value2=$r.Estado
        if ($row2 % 2 -eq 0) { $sh2.Rows.Item($row2).Interior.Color=0xFCE4E4 }
        $row2++
    }
    $sh2.Columns.AutoFit() | Out-Null
    if (Test-Path $reportPath) { Remove-Item $reportPath -Force }
    $wb.SaveAs($reportPath,51); $wb.Close($false); $xl.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
    Write-Host "Reporte generado: $reportPath" -ForegroundColor Cyan
}

$accessToken = Get-AccessToken
if (-not $accessToken) { Write-Error "No se pudo obtener el token."; exit }

Write-Output "Cargando datos desde Excel..."
$excel = New-Object -ComObject Excel.Application; $excel.Visible=$false; $excel.DisplayAlerts=$false

$wbCorreos = $excel.Workbooks.Open($correosPath)
$shPlanta  = $wbCorreos.Sheets.Item("Docentes Planta")

function Get-CellValue($sheet,$row,$col) {
    $cell=$sheet.Cells.Item($row,$col)
    if ($cell.MergeCells) { return $cell.MergeArea.Item(1,1).Text.Trim() }
    return $cell.Text.Trim()
}

$docenteMap = @{}
$r=2
while ($true) {
    $dni=Get-CellValue $shPlanta $r 1; if (-not $dni) { break }
    $cleanDni=$dni -replace '[^\d]',''
    $docenteMap[$cleanDni]=@{ Dni=$cleanDni; Nombres=Get-CellValue $shPlanta $r 2; Apellidos=Get-CellValue $shPlanta $r 3; Correo=Get-CellValue $shPlanta $r 9; Programa=""; Facultad="" }
    $r++
}
$wbCorreos.Close($false)

function Get-Teacher($dni,$nombre,$prog,$fac) {
    $d=$dni -replace '[^\d]',''; if (-not $d) { return $null }
    if (-not $docenteMap.ContainsKey($d)) { $docenteMap[$d]=@{Dni=$d;Nombres=$nombre;Apellidos="";Correo="";Programa=$prog;Facultad=$fac} }
    $doc=$docenteMap[$d]
    if (-not $doc.Programa -and $prog) { $doc.Programa=$prog }
    if (-not $doc.Facultad -and $fac) { $doc.Facultad=$fac }
    return $doc
}

function Get-NotaComision($obs) {
    if ($obs -match "comisi") {
        if ($obs -match "(Los puntos[^\n]+comisi[^\n]+\.)") { return $matches[1] }
        return $notaComisionDefault
    }
    return ""
}

$wbBonif = $excel.Workbooks.Open($ciarpPath)
$allProducts = @()

function Read-BonifSheet($sheetName,$m,$buildDetail) {
    $sh=$wbBonif.Sheets.Item($sheetName)
    $nr=$sh.UsedRange.Rows.Count; $cont=0
    for ($r=2;$r -le $nr;$r++) {
        $acta=Get-CellValue $sh $r $m.Acta
        if ($acta -notmatch $actaRegex) { continue }
        $dni=Get-CellValue $sh $r $m.Dni; if (-not $dni) { continue }
        Get-Teacher $dni (Get-CellValue $sh $r $m.Nombre) (Get-CellValue $sh $r $m.Prog) (Get-CellValue $sh $r $m.Fac) | Out-Null
        $info=& $buildDetail $sh $r
        $obs=if ($m.Obs) { Get-CellValue $sh $r $m.Obs } else { "" }
        $script:allProducts += [PSCustomObject]@{
            Dni=($dni -replace '[^\d]',''); Concepto=$info.Concepto; Detalle=$info.Detalle
            Puntaje=(Get-CellValue $sh $r $m.Pts); Acta=$acta; Nota=(Get-NotaComision $obs)
        }
        $cont++
    }
    Write-Output ("  {0}: {1} filas del acta" -f $sheetName,$cont)
}

Read-BonifSheet "Ponencias" @{Dni=4;Nombre=5;Prog=7;Fac=8;Pts=14;Acta=15;Obs=17} {
    param($sh,$r)
    $titulo=Get-CellValue $sh $r 10; $evento=Get-CellValue $sh $r 11; $tipo=Get-CellValue $sh $r 13
    @{ Concepto="PONENCIA"; Detalle="`"$titulo`" - $evento ($tipo)" }
}

Read-BonifSheet "Rev. No Index" @{Dni=4;Nombre=5;Prog=6;Fac=7;Pts=13;Acta=14;Obs=15} {
    param($sh,$r)
    $titulo=Get-CellValue $sh $r 9; $revista=Get-CellValue $sh $r 11
    @{ Concepto="REVISTA"; Detalle="`"$titulo`" - Revista $revista" }
}

Read-BonifSheet "Tesis en Postgrados" @{Dni=4;Nombre=5;Prog=7;Fac=8;Pts=14;Acta=15;Obs=16} {
    param($sh,$r)
    $tesis=Get-CellValue $sh $r 9; $est=Get-CellValue $sh $r 10; $titOpt=Get-CellValue $sh $r 11
    @{ Concepto="TESIS POSGRADO"; Detalle="`"$tesis`" - Dir.: $est ($titOpt)" }
}

Read-BonifSheet "Pos-Doctorado" @{Dni=4;Nombre=5;Prog=9;Fac=10;Pts=16;Acta=17;Obs=18} {
    param($sh,$r)
    $proyecto=Get-CellValue $sh $r 13; $periodo=Get-CellValue $sh $r 14; $fechas=Get-CellValue $sh $r 15
    @{ Concepto="POS-DOCTORADO"; Detalle="`"$proyecto`" ($periodo, $fechas)" }
}

$wbBonif.Close($false); $excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Output "Datos cargados: $($allProducts.Count) productos en total."

$validProducts = @(); $excludedProducts = @()
foreach ($p in $allProducts) {
    $ptsStr=($p.Puntaje -replace ',','.').Trim(); $v=0.0
    $ok=[double]::TryParse($ptsStr,[System.Globalization.NumberStyles]::Any,[System.Globalization.CultureInfo]::InvariantCulture,[ref]$v)
    if (-not $ok -or $v -eq 0.0 -or $ptsStr -eq "") { $excludedProducts+=$p } else { $validProducts+=$p }
}
Write-Output "Productos validos (puntaje > 0): $($validProducts.Count)"
Write-Output "Productos excluidos: $($excludedProducts.Count)"
$groupProducts = $validProducts | Group-Object Dni

$reportRows = @()
foreach ($p in $excludedProducts) {
    $doc=$docenteMap[$p.Dni]
    $reportRows += [PSCustomObject]@{
        Dni=$p.Dni; Nombre=if($doc){($doc.Nombres+" "+$doc.Apellidos).Trim()}else{""}
        Correo=if($doc){$doc.Correo}else{""}; Programa=if($doc){$doc.Programa}else{""}
        Facultad=if($doc){$doc.Facultad}else{""}; Concepto=$p.Concepto; Detalle=$p.Detalle
        Puntaje=$p.Puntaje; Acta=$p.Acta; Estado="Puntaje cero o vacio"
    }
}

if ($Test) {
    Write-Output "MODO PRUEBA: enviando primer docente a jhvargas"
    if ($groupProducts.Count -eq 0) { Write-Host "Sin productos validos." -ForegroundColor Red; exit }
    $g=$groupProducts[0]; $dni=$g.Name; $doc=$docenteMap[$dni]; $products=$g.Group
    $fullName=($doc.Nombres+" "+$doc.Apellidos).Trim()
    $firstActa="4- 02/10/2026"; foreach ($p in $products) { if ($p.Acta) { $firstActa=$p.Acta } }
    $rowsHtml=Build-RowsHtml $products
    $sumStr=(Get-ProductSum $products).ToString("F1") -replace '\.',','
    $nota=($products|Where-Object{$_.Nota}|Select-Object -ExpandProperty Nota -Unique) -join " "
    $htmlBody=Build-EmailHtml $fullName $doc.Programa $doc.Facultad $firstActa $rowsHtml $sumStr $nota $products.Count
    $success=Send-GmailMessage -accessToken $accessToken -to "jhvargas@uniquindio.edu.co" -subject "[PRUEBA] $subjectText" -htmlBody $htmlBody
    if ($success) { Write-Host "Correo de prueba enviado." -ForegroundColor Green } else { Write-Host "Error." -ForegroundColor Red }
    foreach ($group in $groupProducts) {
        $d=$docenteMap[$group.Name]
        $n=if($d){($d.Nombres+" "+$d.Apellidos).Trim()}else{"Docente $($group.Name)"}
        $co=if($d){$d.Correo}else{""}; $pr=if($d){$d.Programa}else{""}; $fa=if($d){$d.Facultad}else{""}
        $es=if($co){"Notificado (prueba)"}else{"Sin correo registrado"}
        foreach ($p in $group.Group) {
            $reportRows+=[PSCustomObject]@{Dni=$p.Dni;Nombre=$n;Correo=$co;Programa=$pr;Facultad=$fa;Concepto=$p.Concepto;Detalle=$p.Detalle;Puntaje=$p.Puntaje;Acta=$p.Acta;Estado=$es}
        }
    }
    $ts=Get-Date -Format "yyyyMMdd_HHmmss"
    Generate-Report $reportRows (Join-Path $PSScriptRoot "reporte_bonif_prueba_$ts.xlsx")
}

if ($TestAll) {
    Write-Output "MODO PRUEBA TOTAL: enviando todos a jhvargas"
    if ($groupProducts.Count -eq 0) { Write-Host "Sin productos validos." -ForegroundColor Red; exit }
    Write-Host "Total docentes: $($groupProducts.Count)" -ForegroundColor Yellow
    $sent=0; $fail=0
    foreach ($group in $groupProducts) {
        $dni=$group.Name; $doc=$docenteMap[$dni]; $products=$group.Group
        $fullName=if($doc){($doc.Nombres+" "+$doc.Apellidos).Trim()}else{"Docente $dni"}
        $prog=if($doc){$doc.Programa}else{""}; $fac=if($doc){$doc.Facultad}else{""}
        $correoReal=if($doc){$doc.Correo}else{"sin-correo"}
        $firstActa="4- 02/10/2026"; foreach ($p in $products) { if ($p.Acta) { $firstActa=$p.Acta } }
        $rowsHtml=Build-RowsHtml $products
        $sumStr=(Get-ProductSum $products).ToString("F1") -replace '\.',','
        $nota=($products|Where-Object{$_.Nota}|Select-Object -ExpandProperty Nota -Unique) -join " "
        $htmlBody=Build-EmailHtml $fullName $prog $fac $firstActa $rowsHtml $sumStr $nota $products.Count
        Write-Output "Enviando prueba: $fullName..."
        $ok=Send-GmailMessage -accessToken $accessToken -to "jhvargas@uniquindio.edu.co" -subject "[PRUEBA - Para: $correoReal] $subjectText" -htmlBody $htmlBody
        if ($ok) { $sent++; Write-Output "OK."; $es="Notificado (prueba)" } else { $fail++; $es="Error al enviar (prueba)" }
        foreach ($p in $products) {
            $reportRows+=[PSCustomObject]@{Dni=$p.Dni;Nombre=$fullName;Correo=$correoReal;Programa=$prog;Facultad=$fac;Concepto=$p.Concepto;Detalle=$p.Detalle;Puntaje=$p.Puntaje;Acta=$p.Acta;Estado=$es}
        }
        Start-Sleep -Milliseconds 500
    }
    Write-Host "Enviados: $sent | Fallidos: $fail" -ForegroundColor Cyan
    $ts=Get-Date -Format "yyyyMMdd_HHmmss"
    Generate-Report $reportRows (Join-Path $PSScriptRoot "reporte_bonif_prueba_$ts.xlsx")
}

if ($Real) {
    Write-Output "MODO REAL: Enviando notificaciones de bonificaciones"
    Write-Host "Docentes con productos validos: $($groupProducts.Count)" -ForegroundColor Yellow
    if (-not $Confirmar) {
        $conf=Read-Host "Deseas enviar los correos reales? (Escribe 'S' para confirmar)"
        if ($conf -ne "S" -and $conf -ne "s") { Write-Output "Cancelado."; exit }
    } else {
        Write-Host "Confirmacion automatica (-Confirmar). Iniciando envio..." -ForegroundColor Cyan
    }
    $bcc="asuntosprofesorales@uniquindio.edu.co"; $sent=0; $fail=0
    Write-Host "BCC: $bcc" -ForegroundColor Cyan
    foreach ($group in $groupProducts) {
        $dni=$group.Name; $doc=$docenteMap[$dni]; $products=$group.Group
        $fullName=if($doc){($doc.Nombres+" "+$doc.Apellidos).Trim()}else{"Docente $dni"}
        $prog=if($doc){$doc.Programa}else{""}; $fac=if($doc){$doc.Facultad}else{""}
        if (-not $doc) {
            Write-Warning "Sin datos para DNI $dni"
            foreach ($p in $products) {
                $reportRows+=[PSCustomObject]@{Dni=$p.Dni;Nombre=$fullName;Correo="";Programa="";Facultad="";Concepto=$p.Concepto;Detalle=$p.Detalle;Puntaje=$p.Puntaje;Acta=$p.Acta;Estado="Sin datos de docente"}
            }
            $fail++; continue
        }
        $email=$doc.Correo
        if (-not $email) {
            Write-Warning "Sin correo: $fullName"
            foreach ($p in $products) {
                $reportRows+=[PSCustomObject]@{Dni=$p.Dni;Nombre=$fullName;Correo="";Programa=$prog;Facultad=$fac;Concepto=$p.Concepto;Detalle=$p.Detalle;Puntaje=$p.Puntaje;Acta=$p.Acta;Estado="Sin correo registrado"}
            }
            $fail++; continue
        }
        $firstActa="4- 02/10/2026"; foreach ($p in $products) { if ($p.Acta) { $firstActa=$p.Acta } }
        $rowsHtml=Build-RowsHtml $products
        $sumStr=(Get-ProductSum $products).ToString("F1") -replace '\.',','
        $nota=($products|Where-Object{$_.Nota}|Select-Object -ExpandProperty Nota -Unique) -join " "
        $htmlBody=Build-EmailHtml $fullName $prog $fac $firstActa $rowsHtml $sumStr $nota $products.Count
        Write-Output "Enviando a: $fullName ($email)..."
        $ok=Send-GmailMessage -accessToken $accessToken -to $email -subject $subjectText -htmlBody $htmlBody -bcc $bcc
        if ($ok) {
            $sent++; Write-Output "OK."
            $es="Notificado"
        } else {
            $fail++; $es="Error al enviar"
        }
        foreach ($p in $products) {
            $reportRows+=[PSCustomObject]@{Dni=$p.Dni;Nombre=$fullName;Correo=$email;Programa=$prog;Facultad=$fac;Concepto=$p.Concepto;Detalle=$p.Detalle;Puntaje=$p.Puntaje;Acta=$p.Acta;Estado=$es}
        }
        Start-Sleep -Seconds 1
    }
    Write-Host "Proceso terminado." -ForegroundColor Green
    Write-Host "Enviados: $sent | Fallidos: $fail" -ForegroundColor $(if ($fail -gt 0){"Red"}else{"Green"})
    $ts=Get-Date -Format "yyyyMMdd_HHmmss"
    Generate-Report $reportRows (Join-Path $PSScriptRoot "reporte_bonif_$ts.xlsx")
}
