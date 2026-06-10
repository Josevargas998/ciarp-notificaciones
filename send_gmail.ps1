# Script to send CIARP docente notifications via Gmail REST API
param (
    [switch]$Test,
    [switch]$TestAll,
    [switch]$Real
)

$ciarpPath    = Join-Path $PSScriptRoot "ciarp 2.xlsx"
$correosPath  = Join-Path $PSScriptRoot "correos.xlsx"
$tokensPath   = Join-Path $PSScriptRoot "tokens.json"
$credsPath    = Join-Path $PSScriptRoot "credentials.json"

# Unicode character codes for Spanish accents
$char_i_accent = [char]0xed
$char_a_accent = [char]0xe1
$char_o_accent = [char]0xf3
$char_e_accent = [char]0xe9

# Nota de comision (sin tildes hardcodeadas para evitar problemas de encoding)
$notaComisionDefault = "Los puntos ser" + $char_a_accent + "n pagados al t" + $char_e_accent + "rmino de la comisi" + $char_o_accent + "n acad" + $char_e_accent + "mica-administrativa."

$tituloConcepto  = "T" + $char_i_accent + "tulo universitario de posgrado"
$articuloConcepto = "Art" + $char_i_accent + "culo en revista indexada"
$luzName         = "Luz Amparo Celis Buritic" + $char_a_accent
$subjectText     = "Respuesta a solicitud de asignaci" + $char_o_accent + "n de puntos salariales"

if (-not $Test -and -not $TestAll -and -not $Real) {
    Write-Host "Debes especificar -Test, -TestAll o -Real." -ForegroundColor Yellow
    Write-Host "  -Test    : Envia 1 correo de prueba (primer docente) a jhvargas"
    Write-Host "  -TestAll : Envia TODOS los correos a jhvargas (para revision)"
    Write-Host "  -Real    : Envia correos reales a cada docente"
    exit
}

# ---------------------------------------------------------------------------
# OAUTH AUTHENTICATION
# ---------------------------------------------------------------------------
function Get-AccessToken {
    if (-not (Test-Path $credsPath)) {
        Write-Error "No se encuentra credentials.json en $credsPath."
        return $null
    }

    $creds = Get-Content $credsPath -Raw | ConvertFrom-Json
    $clientId     = $creds.installed.client_id
    $clientSecret = $creds.installed.client_secret

    $tokens = $null
    if (Test-Path $tokensPath) {
        $tokens = Get-Content $tokensPath -Raw | ConvertFrom-Json
    }

    if (-not $tokens -or -not $tokens.refresh_token) {
        Write-Output "--------------------------------------------------------"
        Write-Output "INICIANDO AUTENTICACION CON GOOGLE OAUTH"
        Write-Output "--------------------------------------------------------"
        Write-Output "Se abrira el navegador. Inicia sesion con jhvargas@uniquindio.edu.co"
        Write-Output "y autoriza la aplicacion CIARP."
        Write-Output "Si aparece advertencia, haz clic en 'Configuracion avanzada' -> 'Ir a CIARP...'"
        Write-Output "--------------------------------------------------------"

        $port        = 8080
        $redirectUri = "http://localhost:$port/"
        $scope       = "https://www.googleapis.com/auth/gmail.send"

        $escClientId = [uri]::EscapeDataString($clientId)
        $escRedirect = [uri]::EscapeDataString($redirectUri)
        $escScope    = [uri]::EscapeDataString($scope)

        $authUrl = "https://accounts.google.com/o/oauth2/auth?client_id=$escClientId&redirect_uri=$escRedirect&response_type=code&scope=$escScope&access_type=offline&prompt=consent"

        $listener = New-Object System.Net.HttpListener
        $listener.Prefixes.Add($redirectUri)
        try { $listener.Start() } catch {
            Write-Error "No se pudo iniciar el servidor local en puerto $port."
            return $null
        }

        Start-Process $authUrl

        $context = $listener.GetContext()
        $code    = $context.Request.QueryString["code"]

        $responseHtml = "<html><body style='font-family:Arial;text-align:center;padding:50px'><h1 style='color:#2e7d32'>Autenticacion Exitosa!</h1><p>Puedes cerrar esta pestana.</p></body></html>"
        $buffer = [System.Text.Encoding]::UTF8.GetBytes($responseHtml)
        $context.Response.ContentLength64 = $buffer.Length
        $context.Response.OutputStream.Write($buffer, 0, $buffer.Length)
        $context.Response.OutputStream.Close()
        $listener.Stop()

        if (-not $code) { Write-Error "No se obtuvo el codigo de autorizacion."; return $null }

        Write-Output "Intercambiando codigo por tokens..."
        $tokenBody = @{
            code          = $code
            client_id     = $clientId
            client_secret = $clientSecret
            redirect_uri  = $redirectUri
            grant_type    = "authorization_code"
        }
        try {
            $tokenRes = Invoke-RestMethod -Uri "https://oauth2.googleapis.com/token" -Method Post -Body $tokenBody
            $expiry   = (Get-Date).AddSeconds($tokenRes.expires_in).ToString("o")
            $tokens   = @{ access_token = $tokenRes.access_token; refresh_token = $tokenRes.refresh_token; expiry_time = $expiry }
            $tokens | ConvertTo-Json | Out-File $tokensPath -Encoding utf8
            Write-Output "Autenticacion guardada."
        } catch {
            Write-Error "Error al obtener token: $_"; return $null
        }
    } else {
        $expiryTime = [DateTime]$tokens.expiry_time
        if ((Get-Date).AddMinutes(5) -ge $expiryTime) {
            Write-Output "Refrescando token..."
            $refreshBody = @{
                client_id     = $clientId
                client_secret = $clientSecret
                refresh_token = $tokens.refresh_token
                grant_type    = "refresh_token"
            }
            try {
                $tokenRes = Invoke-RestMethod -Uri "https://oauth2.googleapis.com/token" -Method Post -Body $refreshBody
                $expiry   = (Get-Date).AddSeconds($tokenRes.expires_in).ToString("o")
                $tokens.access_token = $tokenRes.access_token
                $tokens.expiry_time  = $expiry
                $tokens | ConvertTo-Json | Out-File $tokensPath -Encoding utf8
                Write-Output "Token refrescado."
            } catch {
                Write-Warning "No se pudo refrescar. Reautenticando..."
                Remove-Item $tokensPath -ErrorAction SilentlyContinue
                return Get-AccessToken
            }
        }
    }
    return $tokens.access_token
}

# ---------------------------------------------------------------------------
# SEND EMAIL VIA GMAIL API
# ---------------------------------------------------------------------------
function Send-GmailMessage {
    param ([string]$accessToken, [string]$to, [string]$subject, [string]$htmlBody, [string]$bcc="")

    $subjectBase64  = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($subject))
    $fromName       = "CIARP Uniquindi" + [char]0xf3
    $fromNameBase64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($fromName))
    $fromHeader     = "=?utf-8?B?" + $fromNameBase64 + "?= <jhvargas@uniquindio.edu.co>"

    $bccHeader = ""
    if ($bcc) { $bccHeader = "`r`nBcc: " + $bcc }

    $mimeMessage = "From: " + $fromHeader + "`r`nTo: " + $to + $bccHeader + "`r`nSubject: =?utf-8?B?" + $subjectBase64 + "?=`r`nMIME-Version: 1.0`r`nContent-Type: text/html; charset=utf-8`r`n`r`n" + $htmlBody

    $bytes     = [System.Text.Encoding]::UTF8.GetBytes($mimeMessage)
    $base64    = [Convert]::ToBase64String($bytes)
    $base64Url = $base64.Replace('+', '-').Replace('/', '_').Replace('=', '')

    $bodyJson = @{ raw = $base64Url } | ConvertTo-Json
    $headers  = @{ Authorization = "Bearer $accessToken"; "Content-Type" = "application/json" }

    try {
        Invoke-RestMethod -Uri "https://gmail.googleapis.com/gmail/v1/users/me/messages/send" -Method Post -Headers $headers -Body $bodyJson | Out-Null
        return $true
    } catch {
        Write-Error "Error al enviar correo a $to : $_"
        if ($_.Exception.Response) {
            $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
            Write-Error "Respuesta API: $($reader.ReadToEnd())"
        }
        return $false
    }
}

# ---------------------------------------------------------------------------
# BUILD HTML EMAIL BODY
# ---------------------------------------------------------------------------
function Build-EmailHtml ($fullName, $programa, $facultad, $firstActa, $rowsHtml, $sumStr, $notaComision="") {
    $notaHtml = ""
    $parrafoActo = ""
    if ($notaComision) {
        # Cuando hay comision: fusionar la nota + el parrafo del acto administrativo en la caja azul
        $notaHtml = "<div style='background-color:#e3f2fd;border-left:4px solid #1565c0;padding:14px 16px;margin-top:20px;border-radius:0 4px 4px 0;font-size:14px;color:#0d47a1'>" +
                    "<strong>Nota sobre comisi&oacute;n acad&eacute;mica-administrativa:</strong><br>" + $notaComision +
                    "<br><br>Una vez concluida la comisi&oacute;n, con posterioridad se estar&aacute; notificando el correspondiente acto administrativo.</div>"
    } else {
        # Sin comision: parrafo estandar
        $parrafoActo = "<p style='margin-top:20px'>Por lo anteriormente mencionado, con posterioridad se estara notificando el correspondiente acto administrativo.</p>"
    }
    return "<html><head><meta charset='utf-8'><style>body{font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;background-color:#f4f6f8;margin:0;padding:20px}.card{max-width:650px;margin:0 auto;background:#fff;border:1px solid #e0e0e0;border-radius:8px;overflow:hidden;box-shadow:0 4px 10px rgba(0,0,0,.05)}.header{background-color:#1b5e20;color:#fff;padding:25px 20px;text-align:center}.header h1{margin:0;font-size:26px;font-weight:700;letter-spacing:1px}.header p{margin:6px 0 0;font-size:13px;opacity:.9}.content{padding:30px;color:#2d3748;line-height:1.6}.docente-info{background:#f8fafc;border-left:4px solid #1b5e20;padding:15px;margin-bottom:25px;border-radius:0 4px 4px 0}.docente-info p{margin:0;font-size:14px;color:#4a5568}.docente-info .name{font-size:16px;font-weight:700;color:#1b5e20;text-transform:uppercase;margin-bottom:4px}.greeting{font-size:15px;font-weight:700;margin-bottom:15px}.table-container{margin:20px 0;overflow-x:auto}table{width:100%;border-collapse:collapse;border:1px solid #e0e0e0}th{background-color:#1b5e20;color:#fff;padding:12px;font-size:14px;text-align:left}.total-box{background-color:#e8f5e9;border:1px solid #c8e6c9;padding:12px 15px;border-radius:4px;font-size:15px;font-weight:700;color:#2e7d32;margin-top:15px}.signature{margin-top:30px;font-size:14px;color:#4a5568;line-height:1.4;border-top:1px solid #e2e8f0;padding-top:20px}</style></head><body><div class='card'><div class='header'><h1>CIARP</h1><p>Comite Interno de Asignacion y Reconocimiento de Puntaje</p></div><div class='content'><h2 style='color:#1b5e20;text-align:center;margin-top:0;margin-bottom:25px;font-size:20px'>Reconocimiento de Puntos Salariales</h2><div class='docente-info'><p class='name'>" + $fullName + "</p><p><strong>Programa:</strong> " + $programa + "</p><p><strong>Facultad:</strong> " + $facultad + "</p></div><p class='greeting'>Cordial saludo,</p><p>Dando cumplimiento a la directriz emitida por la jefe de la Oficina de Asuntos Profesionales, me permito informarle que, en sesion del Comite Interno de Asignacion y Reconocimiento de Puntaje (CIARP), llevada a cabo segun consta en <strong>" + $firstActa + "</strong>, la solicitud de asignacion de puntos salariales presentada por usted fue considerada y aprobada como se relaciona a continuacion:</p><div class='table-container'><table><thead><tr><th style='width:30%'>Concepto</th><th style='width:50%'>Detalle del Producto</th><th style='width:20%;text-align:right'>Puntaje</th></tr></thead><tbody>" + $rowsHtml + "</tbody></table></div><div class='total-box'>Puntaje total aprobado en esta sesion: " + $sumStr + " puntos</div>" + $notaHtml + $parrafoActo + "<p>Cualquier inquietud al respecto, estaremos atentos.</p><div class='signature'>Atentamente,<br><br><div style='background-color:#f1f8e9;border:1px solid #dcedc8;padding:15px;border-radius:4px;text-align:center;color:#1b5e20;font-weight:700;font-size:16px'>Luz Amparo Celis Buritic&aacute;<br><span style='font-size:14px;font-weight:400;color:#2e7d32'>Jefe Oficina de Asuntos Profesorales</span></div><p style='margin-top:12px;font-size:13px;color:#718096'>Proyect&oacute;: Jos&eacute; Heriberto Vargas Espinosa</p></div></div></div></body></html>"
}

function Build-RowsHtml ($products) {
    $rows = ""
    foreach ($p in $products) {
        $rows += "<tr style='border-bottom:1px solid #e0e0e0'><td style='padding:12px;font-size:14px;color:#333;font-weight:700'>" + $p.Concepto + "</td><td style='padding:12px;font-size:14px;color:#555'>" + $p.Detalle + "</td><td style='padding:12px;font-size:14px;color:#333;text-align:right;font-weight:700'>" + $p.Puntaje + " puntos</td></tr>"
    }
    return $rows
}

function Get-ProductSum ($products) {
    $sum = 0.0
    foreach ($p in $products) {
        $ptsStr = $p.Puntaje -replace ',', '.'
        $ptsVal = 0.0
        if ([double]::TryParse($ptsStr, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$ptsVal)) {
            $sum += $ptsVal
        }
    }
    return $sum
}

# ---------------------------------------------------------------------------
# GENERATE EXCEL REPORT
# ---------------------------------------------------------------------------
function Generate-Report ($reportRows, $reportPath) {
    Write-Output "Generando reporte Excel..."
    $xlApp = New-Object -ComObject Excel.Application
    $xlApp.Visible = $false
    $xlApp.DisplayAlerts = $false
    $wb = $xlApp.Workbooks.Add()

    # Sheet 1: Notificados
    $sh1 = $wb.Sheets.Item(1)
    $sh1.Name = "Notificados"
    $h1 = @("DNI","Nombre Docente","Correo","Programa","Facultad","Concepto","Detalle del Producto","Puntaje","Acta")
    for ($c = 0; $c -lt $h1.Count; $c++) {
        $cell = $sh1.Cells.Item(1, $c+1)
        $cell.Value2 = $h1[$c]; $cell.Font.Bold = $true
        $cell.Interior.Color = 0x1e7a1e; $cell.Font.Color = 0xFFFFFF
    }
    $row1 = 2
    foreach ($r in ($reportRows | Where-Object { $_.Estado -like "Notificado*" })) {
        $sh1.Cells.Item($row1,1).Value2 = $r.Dni
        $sh1.Cells.Item($row1,2).Value2 = $r.Nombre
        $sh1.Cells.Item($row1,3).Value2 = $r.Correo
        $sh1.Cells.Item($row1,4).Value2 = $r.Programa
        $sh1.Cells.Item($row1,5).Value2 = $r.Facultad
        $sh1.Cells.Item($row1,6).Value2 = $r.Concepto
        $sh1.Cells.Item($row1,7).Value2 = $r.Detalle
        $sh1.Cells.Item($row1,8).Value2 = $r.Puntaje
        $sh1.Cells.Item($row1,9).Value2 = $r.Acta
        if ($row1 % 2 -eq 0) { $sh1.Rows.Item($row1).Interior.Color = 0xE8F5E9 }
        $row1++
    }
    $sh1.Columns.AutoFit() | Out-Null

    # Sheet 2: No Notificados
    $sh2 = $wb.Sheets.Add()
    $sh2.Name = "No Notificados"
    $h2 = @("DNI","Nombre Docente","Correo","Programa","Facultad","Concepto","Detalle del Producto","Puntaje","Motivo")
    for ($c = 0; $c -lt $h2.Count; $c++) {
        $cell = $sh2.Cells.Item(1, $c+1)
        $cell.Value2 = $h2[$c]; $cell.Font.Bold = $true
        $cell.Interior.Color = 0x7a1e1e; $cell.Font.Color = 0xFFFFFF
    }
    $row2 = 2
    foreach ($r in ($reportRows | Where-Object { $_.Estado -notlike "Notificado*" })) {
        $sh2.Cells.Item($row2,1).Value2 = $r.Dni
        $sh2.Cells.Item($row2,2).Value2 = $r.Nombre
        $sh2.Cells.Item($row2,3).Value2 = $r.Correo
        $sh2.Cells.Item($row2,4).Value2 = $r.Programa
        $sh2.Cells.Item($row2,5).Value2 = $r.Facultad
        $sh2.Cells.Item($row2,6).Value2 = $r.Concepto
        $sh2.Cells.Item($row2,7).Value2 = $r.Detalle
        $sh2.Cells.Item($row2,8).Value2 = $r.Puntaje
        $sh2.Cells.Item($row2,9).Value2 = $r.Estado
        if ($row2 % 2 -eq 0) { $sh2.Rows.Item($row2).Interior.Color = 0xFCE4E4 }
        $row2++
    }
    $sh2.Columns.AutoFit() | Out-Null

    if (Test-Path $reportPath) { Remove-Item $reportPath -Force }
    $wb.SaveAs($reportPath, 51)
    $wb.Close($false)
    $xlApp.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xlApp) | Out-Null
    Write-Host "Reporte generado: $reportPath" -ForegroundColor Cyan
}

# ---------------------------------------------------------------------------
# AUTHENTICATE
# ---------------------------------------------------------------------------
$accessToken = Get-AccessToken
if (-not $accessToken) { Write-Error "No se pudo obtener el token. Abortando."; exit }

# ---------------------------------------------------------------------------
# LOAD EXCEL DATA
# ---------------------------------------------------------------------------
Write-Output "Cargando datos desde Excel..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# Load correos.xlsx -> Docentes Planta
$wbCorreos = $excel.Workbooks.Open($correosPath)
$shPlanta  = $wbCorreos.Sheets.Item("Docentes Planta")

function Get-CellValue($sheet, $row, $col) {
    $cell = $sheet.Cells.Item($row, $col)
    if ($cell.MergeCells) {
        return $cell.MergeArea.Item(1,1).Text.Trim()
    }
    return $cell.Text.Trim()
}

$docenteMap = @{}
$r = 2
while ($true) {
    $dni = Get-CellValue $shPlanta $r 1
    if (-not $dni) { break }
    $cleanDni = $dni -replace '[^\d]', ''
    $docenteMap[$cleanDni] = @{
        Dni       = $cleanDni
        Nombres   = Get-CellValue $shPlanta $r 2
        Apellidos = Get-CellValue $shPlanta $r 3
        Correo    = Get-CellValue $shPlanta $r 9
        Programa  = ""
        Facultad  = ""
    }
    $r++
}
$wbCorreos.Close($false)

function Get-Teacher($dni, $nombreDocente, $programa, $facultad) {
    $cleanDni = $dni -replace '[^\d]', ''
    if (-not $cleanDni) { return $null }
    if (-not $docenteMap.ContainsKey($cleanDni)) {
        $docenteMap[$cleanDni] = @{ Dni=$cleanDni; Nombres=$nombreDocente; Apellidos=""; Correo=""; Programa=$programa; Facultad=$facultad }
    }
    $doc = $docenteMap[$cleanDni]
    if (-not $doc.Programa -and $programa) { $doc.Programa = $programa }
    if (-not $doc.Facultad -and $facultad) { $doc.Facultad = $facultad }
    return $doc
}

# Load ciarp 2.xlsx
$wbCiarp    = $excel.Workbooks.Open($ciarpPath)
$allProducts = @()

# Titulo
$shTitulo = $wbCiarp.Sheets.Item("Titulo")
$r = 3
$maxEmpty = 0
while ($maxEmpty -lt 8) {
    $dni = Get-CellValue $shTitulo $r 4
    if (-not $dni) { $maxEmpty++; $r++; continue }
    $maxEmpty = 0
    $nombre   = Get-CellValue $shTitulo $r 5
    $programa = Get-CellValue $shTitulo $r 9
    $facultad = Get-CellValue $shTitulo $r 10
    Get-Teacher $dni $nombre $programa $facultad | Out-Null
    $univ     = Get-CellValue $shTitulo $r 11
    $titulo   = Get-CellValue $shTitulo $r 12
    $fechaGrad = Get-CellValue $shTitulo $r 14
    $pts      = Get-CellValue $shTitulo $r 15
    $acta     = Get-CellValue $shTitulo $r 16
    $obs      = Get-CellValue $shTitulo $r 17
    $detail   = $titulo + " - " + $univ + " (Fecha de grado: " + $fechaGrad + ")"
    $nota = ""
    if ($obs -match "comisi") {
        if ($obs -match "(Los puntos[^\n]+comisi[^\n]+\.)") { $nota = $matches[1] }
        else { $nota = $notaComisionDefault }
    }
    $allProducts += [PSCustomObject]@{ Dni=($dni -replace '[^\d]',''); Concepto=$tituloConcepto; Detalle=$detail; Puntaje=$pts; Acta=$acta; Nota=$nota }
    $r++
}

# Pub_Rev_Index
$shPub = $wbCiarp.Sheets.Item("Pub_Rev_Index")
$r = 3
$maxEmpty = 0
while ($maxEmpty -lt 8) {
    $dni       = Get-CellValue $shPub $r 17
    $tituloArt = Get-CellValue $shPub $r 4
    if (-not $dni -and -not $tituloArt) { $maxEmpty++; $r++; continue }
    $maxEmpty = 0
    if ($dni) {
        $nombre   = Get-CellValue $shPub $r 18
        $programa = Get-CellValue $shPub $r 22
        $facultad = Get-CellValue $shPub $r 23
        Get-Teacher $dni $nombre $programa $facultad | Out-Null
        $revista  = Get-CellValue $shPub $r 9
        $cat      = Get-CellValue $shPub $r 11
        $pts      = Get-CellValue $shPub $r 24
        $tipo     = Get-CellValue $shPub $r 6
        $obs      = Get-CellValue $shPub $r 27
        $concepto = if ($tipo -like "*Editorial*") { "Editorial en revista indexada" } else { $articuloConcepto }
        $detail   = '"' + $tituloArt + '" - Revista ' + $revista + ' (Categor' + $char_i_accent + 'a ' + $cat + ')'
        $nota = ""
        if ($obs -match "comisi") {
            if ($obs -match "(Los puntos[^\n]+comisi[^\n]+\.)") { $nota = $matches[1] }
            else { $nota = $notaComisionDefault }
        }
        $allProducts += [PSCustomObject]@{ Dni=($dni -replace '[^\d]',''); Concepto=$concepto; Detalle=$detail; Puntaje=$pts; Acta="Acta No. 2 del 04 de junio de 2026"; Nota=$nota }
    }
    $r++
}

# Libro_Ensayo
$shEnsayo = $wbCiarp.Sheets.Item("Libro_Ensayo")
$r = 2
$maxEmpty = 0
while ($maxEmpty -lt 8) {
    $dni        = Get-CellValue $shEnsayo $r 8
    $nombreLibro = Get-CellValue $shEnsayo $r 4
    if (-not $dni -and -not $nombreLibro) { $maxEmpty++; $r++; continue }
    $maxEmpty = 0
    if ($dni) {
        $nombre   = Get-CellValue $shEnsayo $r 9
        $programa = Get-CellValue $shEnsayo $r 13
        $facultad = Get-CellValue $shEnsayo $r 14
        Get-Teacher $dni $nombre $programa $facultad | Out-Null
        $isbn  = Get-CellValue $shEnsayo $r 5
        $pts   = Get-CellValue $shEnsayo $r 19
        $acta  = Get-CellValue $shEnsayo $r 20
        $obs   = Get-CellValue $shEnsayo $r 21
        $detail = '"' + $nombreLibro + '" (ISBN: ' + $isbn + ')'
        $nota = ""
        if ($obs -match "comisi") {
            if ($obs -match "(Los puntos[^\n]+comisi[^\n]+\.)") { $nota = $matches[1] }
            else { $nota = $notaComisionDefault }
        }
        $allProducts += [PSCustomObject]@{ Dni=($dni -replace '[^\d]',''); Concepto="Libro de ensayo"; Detalle=$detail; Puntaje=$pts; Acta=$acta; Nota=$nota }
    }
    $r++
}

# Libro_Texto
$shTexto = $wbCiarp.Sheets.Item("Libro_Texto")
$r = 3
$maxEmpty = 0
while ($maxEmpty -lt 8) {
    $dni        = Get-CellValue $shTexto $r 9
    $nombreLibro = Get-CellValue $shTexto $r 4
    if (-not $dni -and -not $nombreLibro) { $maxEmpty++; $r++; continue }
    $maxEmpty = 0
    if ($dni) {
        $nombre   = Get-CellValue $shTexto $r 10
        $programa = Get-CellValue $shTexto $r 14
        $facultad = Get-CellValue $shTexto $r 15
        Get-Teacher $dni $nombre $programa $facultad | Out-Null
        $isbn  = Get-CellValue $shTexto $r 5
        $pts   = Get-CellValue $shTexto $r 20
        $acta  = Get-CellValue $shTexto $r 21
        $obs   = Get-CellValue $shTexto $r 22
        $detail = '"' + $nombreLibro + '" (ISBN: ' + $isbn + ')'
        $nota = ""
        if ($obs -match "comisi") {
            if ($obs -match "(Los puntos[^\n]+comisi[^\n]+\.)") { $nota = $matches[1] }
            else { $nota = $notaComisionDefault }
        }
        $allProducts += [PSCustomObject]@{ Dni=($dni -replace '[^\d]',''); Concepto="Libro de texto"; Detalle=$detail; Puntaje=$pts; Acta=$acta; Nota=$nota }
    }
    $r++
}

# Premios
$shPremios = $wbCiarp.Sheets.Item("Premios")
$r = 3
$maxEmpty = 0
while ($maxEmpty -lt 8) {
    $dni     = Get-CellValue $shPremios $r 4
    $trabajo = Get-CellValue $shPremios $r 11
    if (-not $dni -and -not $trabajo) { $maxEmpty++; $r++; continue }
    $maxEmpty = 0
    if ($dni) {
        $nombre   = Get-CellValue $shPremios $r 5
        $programa = Get-CellValue $shPremios $r 9
        $facultad = Get-CellValue $shPremios $r 10
        Get-Teacher $dni $nombre $programa $facultad | Out-Null
        $premio  = Get-CellValue $shPremios $r 12
        $entidad = Get-CellValue $shPremios $r 13
        $pts     = Get-CellValue $shPremios $r 15
        $acta    = Get-CellValue $shPremios $r 16
        $obs     = Get-CellValue $shPremios $r 17
        $detail  = '"' + $trabajo + '" - ' + $premio + ', ' + $entidad
        # Detectar nota de comision academica-administrativa
        $nota = ""
        if ($obs -match "comisi") {
            # Extraer solo la oracion relevante sobre el pago al termino de la comision
            if ($obs -match "(Los puntos[^\n]+comisi[^\n]+\.)") {
                $nota = $matches[1]
            } else {
                $nota = $notaComisionDefault
            }
        }
        $allProducts += [PSCustomObject]@{ Dni=($dni -replace '[^\d]',''); Concepto="Premio nacional"; Detalle=$detail; Puntaje=$pts; Acta=$acta; Nota=$nota }
    }
    $r++
}

$wbCiarp.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Output "Datos cargados: $($allProducts.Count) productos en total."

# ---------------------------------------------------------------------------
# FILTER: remove products with zero or empty score
# ---------------------------------------------------------------------------
$validProducts   = @()
$excludedProducts = @()
foreach ($p in $allProducts) {
    $ptsStr = ($p.Puntaje -replace ',', '.').Trim()
    $ptsVal = 0.0
    $isNum  = [double]::TryParse($ptsStr, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$ptsVal)
    if (-not $isNum -or $ptsVal -eq 0.0 -or $ptsStr -eq "") {
        $excludedProducts += $p
    } else {
        $validProducts += $p
    }
}

Write-Output "Productos validos (puntaje > 0): $($validProducts.Count)"
Write-Output "Productos excluidos (puntaje 0 o vacio): $($excludedProducts.Count)"

$groupProducts = $validProducts | Group-Object Dni

# ---------------------------------------------------------------------------
# BUILD REPORT BASE (excluded products always go to No Notificados)
# ---------------------------------------------------------------------------
$reportRows = @()
foreach ($p in $excludedProducts) {
    $doc      = $docenteMap[$p.Dni]
    $nombre   = if ($doc) { ($doc.Nombres + " " + $doc.Apellidos).Trim() } else { "" }
    $correo   = if ($doc) { $doc.Correo } else { "" }
    $programa = if ($doc) { $doc.Programa } else { "" }
    $facultad = if ($doc) { $doc.Facultad } else { "" }
    $reportRows += [PSCustomObject]@{
        Dni=$p.Dni; Nombre=$nombre; Correo=$correo; Programa=$programa; Facultad=$facultad
        Concepto=$p.Concepto; Detalle=$p.Detalle; Puntaje=$p.Puntaje; Acta=$p.Acta
        Estado="Puntaje cero o vacio"
    }
}

# ---------------------------------------------------------------------------
# TEST MODE
# ---------------------------------------------------------------------------
if ($Test) {
    Write-Output "========================================================"
    Write-Output "MODO DE PRUEBA: Enviando un correo de prueba a ti mismo"
    Write-Output "========================================================"

    if ($groupProducts.Count -eq 0) {
        Write-Host "No hay docentes con productos validos." -ForegroundColor Red
        exit
    }

    $testGroup = $groupProducts[0]
    $dni       = $testGroup.Name
    $doc       = $docenteMap[$dni]
    $products  = $testGroup.Group
    $fullName  = ($doc.Nombres + " " + $doc.Apellidos).Trim()
    Write-Output "Preparando correo simulado para: $fullName"

    $firstActa = "Acta No. 2 del 04 de junio de 2026"
    [string[]]$conceptos = @($products | Select-Object -ExpandProperty Concepto -Unique)
    $subjectConcepto = if ($conceptos.Count -eq 1) { " por " + $conceptos[0].ToLower() } else { " por productividad academica" }
    foreach ($p in $products) { if ($p.Acta) { $firstActa = $p.Acta } }

    $rowsHtml = Build-RowsHtml $products
    $sumVal   = Get-ProductSum $products
    $sumStr   = $sumVal.ToString("F1") -replace '\.', ','
    $notaComision = ($products | Where-Object { $_.Nota } | Select-Object -ExpandProperty Nota -Unique) -join " "
    $subject  = "[PRUEBA] " + $subjectText + $subjectConcepto
    $htmlBody = Build-EmailHtml $fullName $doc.Programa $doc.Facultad $firstActa $rowsHtml $sumStr $notaComision

    $testRecipient = "jhvargas@uniquindio.edu.co"
    Write-Output "Enviando correo de prueba a: $testRecipient"
    $success = Send-GmailMessage -accessToken $accessToken -to $testRecipient -subject $subject -htmlBody $htmlBody

    if ($success) {
        Write-Host "Correo de prueba enviado con exito! Revisa tu bandeja en $testRecipient." -ForegroundColor Green
    } else {
        Write-Host "Error al enviar el correo de prueba." -ForegroundColor Red
    }

    # Build preview report: all valid docentes marked as "Notificado (prueba)"
    foreach ($group in $groupProducts) {
        $dniR     = $group.Name
        $docR     = $docenteMap[$dniR]
        $nombreR  = if ($docR) { ($docR.Nombres + " " + $docR.Apellidos).Trim() } else { "Docente $dniR" }
        $correoR  = if ($docR) { $docR.Correo } else { "" }
        $programaR = if ($docR) { $docR.Programa } else { "" }
        $facultadR = if ($docR) { $docR.Facultad } else { "" }
        $estadoR  = if ($correoR) { "Notificado (prueba)" } else { "Sin correo registrado" }
        foreach ($p in $group.Group) {
            $reportRows += [PSCustomObject]@{
                Dni=$p.Dni; Nombre=$nombreR; Correo=$correoR; Programa=$programaR; Facultad=$facultadR
                Concepto=$p.Concepto; Detalle=$p.Detalle; Puntaje=$p.Puntaje; Acta=$p.Acta
                Estado=$estadoR
            }
        }
    }

    $timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
    $reportPath = Join-Path $PSScriptRoot "reporte_ciarp_prueba_$timestamp.xlsx"
    Generate-Report $reportRows $reportPath
}

# ---------------------------------------------------------------------------
# TEST-ALL MODE: send every docente's email to jhvargas for review
# ---------------------------------------------------------------------------
if ($TestAll) {
    Write-Output "========================================================"
    Write-Output "MODO PRUEBA TOTAL: Enviando TODOS los correos a jhvargas"
    Write-Output "========================================================"

    if ($groupProducts.Count -eq 0) {
        Write-Host "No hay docentes con productos validos." -ForegroundColor Red
        exit
    }

    $proxyRecipient = "jhvargas@uniquindio.edu.co"
    Write-Host "Todos los correos se enviaran a: $proxyRecipient" -ForegroundColor Cyan
    Write-Host "Total de docentes a procesar: $($groupProducts.Count)" -ForegroundColor Yellow

    $sentCount = 0
    $failCount = 0

    foreach ($group in $groupProducts) {
        $dni      = $group.Name
        $doc      = $docenteMap[$dni]
        $products = $group.Group
        $fullName = if ($doc) { ($doc.Nombres + " " + $doc.Apellidos).Trim() } else { "Docente $dni" }
        $programa = if ($doc) { $doc.Programa } else { "" }
        $facultad = if ($doc) { $doc.Facultad } else { "" }
        $correoReal = if ($doc) { $doc.Correo } else { "sin-correo" }

        $firstActa = "Acta No. 2 del 04 de junio de 2026"
        [string[]]$conceptos = @($products | Select-Object -ExpandProperty Concepto -Unique)
        $subjectConcepto = if ($conceptos.Count -eq 1) { " por " + $conceptos[0].ToLower() } else { " por productividad academica" }
        foreach ($p in $products) { if ($p.Acta) { $firstActa = $p.Acta } }

        $rowsHtml = Build-RowsHtml $products
        $sumVal   = Get-ProductSum $products
        $sumStr   = $sumVal.ToString("F1") -replace '\.', ','
        $notaComision = ($products | Where-Object { $_.Nota } | Select-Object -ExpandProperty Nota -Unique) -join " "
        # Subject includes real recipient so you can identify each email
        $subject  = "[PRUEBA - Para: $correoReal] " + $subjectText + $subjectConcepto
        $htmlBody = Build-EmailHtml $fullName $programa $facultad $firstActa $rowsHtml $sumStr $notaComision

        Write-Output "Enviando prueba de: $fullName -> $proxyRecipient..."
        $success = Send-GmailMessage -accessToken $accessToken -to $proxyRecipient -subject $subject -htmlBody $htmlBody

        if ($success) {
            $sentCount++
            Write-Output "OK."
            foreach ($p in $products) {
                $reportRows += [PSCustomObject]@{
                    Dni=$p.Dni; Nombre=$fullName; Correo=$correoReal; Programa=$programa; Facultad=$facultad
                    Concepto=$p.Concepto; Detalle=$p.Detalle; Puntaje=$p.Puntaje; Acta=$p.Acta
                    Estado="Notificado (prueba)"
                }
            }
        } else {
            $failCount++
            foreach ($p in $products) {
                $reportRows += [PSCustomObject]@{
                    Dni=$p.Dni; Nombre=$fullName; Correo=$correoReal; Programa=$programa; Facultad=$facultad
                    Concepto=$p.Concepto; Detalle=$p.Detalle; Puntaje=$p.Puntaje; Acta=$p.Acta
                    Estado="Error al enviar (prueba)"
                }
            }
        }

        Start-Sleep -Milliseconds 500
    }

    Write-Host "Prueba total terminada." -ForegroundColor Green
    Write-Host "Enviados: $sentCount | Fallidos: $failCount" -ForegroundColor Cyan

    $timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
    $reportPath = Join-Path $PSScriptRoot "reporte_ciarp_prueba_$timestamp.xlsx"
    Generate-Report $reportRows $reportPath
}

# ---------------------------------------------------------------------------
# REAL MODE

# ---------------------------------------------------------------------------
if ($Real) {
    Write-Output "========================================================"
    Write-Output "MODO REAL: Enviando notificaciones a todos los docentes"
    Write-Output "========================================================"

    $totalCount = $groupProducts.Count
    Write-Host "Docentes con productos validos: $totalCount" -ForegroundColor Yellow
    Write-Host "Productos omitidos (puntaje 0 o vacio): $($excludedProducts.Count)" -ForegroundColor DarkYellow

    $confirmation = Read-Host "Deseas enviar los correos reales? (Escribe 'S' para confirmar)"
    if ($confirmation -ne "S" -and $confirmation -ne "s") {
        Write-Output "Envio cancelado."; exit
    }

    $bccEmail  = "asuntosprofesorales@uniquindio.edu.co"
    $sentCount = 0
    $failCount = 0
    Write-Host "Copia oculta (BCC) en cada correo a: $bccEmail" -ForegroundColor Cyan

    foreach ($group in $groupProducts) {
        $dni      = $group.Name
        $doc      = $docenteMap[$dni]
        $products = $group.Group
        $fullName = if ($doc) { ($doc.Nombres + " " + $doc.Apellidos).Trim() } else { "Docente $dni" }
        $programa = if ($doc) { $doc.Programa } else { "" }
        $facultad = if ($doc) { $doc.Facultad } else { "" }

        if (-not $doc) {
            Write-Warning "Sin datos para DNI $dni"
            foreach ($p in $products) {
                $reportRows += [PSCustomObject]@{ Dni=$p.Dni; Nombre=$fullName; Correo=""; Programa=""; Facultad=""
                    Concepto=$p.Concepto; Detalle=$p.Detalle; Puntaje=$p.Puntaje; Acta=$p.Acta; Estado="Sin datos de docente" }
            }
            $failCount++; continue
        }

        $recipientEmail = $doc.Correo
        if (-not $recipientEmail) {
            Write-Warning "Sin correo: $fullName (DNI $dni)"
            foreach ($p in $products) {
                $reportRows += [PSCustomObject]@{ Dni=$p.Dni; Nombre=$fullName; Correo=""; Programa=$programa; Facultad=$facultad
                    Concepto=$p.Concepto; Detalle=$p.Detalle; Puntaje=$p.Puntaje; Acta=$p.Acta; Estado="Sin correo registrado" }
            }
            $failCount++; continue
        }

        $firstActa = "Acta No. 2 del 04 de junio de 2026"
        [string[]]$conceptos = @($products | Select-Object -ExpandProperty Concepto -Unique)
        $subjectConcepto = if ($conceptos.Count -eq 1) { " por " + $conceptos[0].ToLower() } else { " por productividad academica" }
        foreach ($p in $products) { if ($p.Acta) { $firstActa = $p.Acta } }

        $rowsHtml = Build-RowsHtml $products
        $sumVal   = Get-ProductSum $products
        $sumStr   = $sumVal.ToString("F1") -replace '\.', ','
        $notaComision = ($products | Where-Object { $_.Nota } | Select-Object -ExpandProperty Nota -Unique) -join " "
        $subject  = $subjectText + $subjectConcepto
        $htmlBody = Build-EmailHtml $fullName $programa $facultad $firstActa $rowsHtml $sumStr $notaComision

        Write-Output "Enviando a: $fullName ($recipientEmail)..."
        $success = Send-GmailMessage -accessToken $accessToken -to $recipientEmail -subject $subject -htmlBody $htmlBody -bcc $bccEmail

        if ($success) {
            $sentCount++
            Write-Output "Enviado con exito."
            foreach ($p in $products) {
                $reportRows += [PSCustomObject]@{ Dni=$p.Dni; Nombre=$fullName; Correo=$recipientEmail; Programa=$programa; Facultad=$facultad
                    Concepto=$p.Concepto; Detalle=$p.Detalle; Puntaje=$p.Puntaje; Acta=$p.Acta; Estado="Notificado" }
            }
        } else {
            $failCount++
            foreach ($p in $products) {
                $reportRows += [PSCustomObject]@{ Dni=$p.Dni; Nombre=$fullName; Correo=$recipientEmail; Programa=$programa; Facultad=$facultad
                    Concepto=$p.Concepto; Detalle=$p.Detalle; Puntaje=$p.Puntaje; Acta=$p.Acta; Estado="Error al enviar" }
            }
        }

        Start-Sleep -Seconds 1
    }

    Write-Host "Proceso terminado." -ForegroundColor Green
    Write-Host "Enviados con exito: $sentCount" -ForegroundColor Green
    Write-Host "Fallidos o sin correo: $failCount" -ForegroundColor Red

    $timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
    $reportPath = Join-Path $PSScriptRoot "reporte_ciarp_$timestamp.xlsx"
    Generate-Report $reportRows $reportPath
}
