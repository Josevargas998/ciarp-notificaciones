# Script to send CIARP docente notifications via Gmail REST API
param (
    [switch]$Test,
    [switch]$TestAll,
    [switch]$Real,
    [int]$Sesion
)

$ciarpPath    = Join-Path $PSScriptRoot "novedades2.xlsx"
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
function Build-EmailHtml ($fullName, $programa, $facultad, $firstActa, $productsHtml, $productCount=1) {
    # Parsear acta en dos formatos posibles:
    $actaSessionRef = "<strong>" + $firstActa + "</strong>"
    $actaFechaText  = ""
    $monthNames = @("","enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre")
    if ($firstActa -match "No\.\s*(\d+)\s+del\s+(\d+\s+de\s+\w+)") {
        $actaSessionRef = "sesi&oacute;n <strong>No. " + $matches[1] + "</strong>"
        $actaFechaText  = ", llevada a cabo el d&iacute;a " + $matches[2] + " de la presente anualidad"
    } elseif ($firstActa -match "^(\d+)\s*-\s*(\d{2})/(\d{2})/(\d{4})") {
        $sesNum    = $matches[1]
        $day       = $matches[2]
        $monthNum  = [int]$matches[3]
        $monthName = $monthNames[$monthNum]
        $actaSessionRef = "sesi&oacute;n <strong>No. $sesNum</strong>"
        $actaFechaText  = ", llevada a cabo el d&iacute;a $day de $monthName de la presente anualidad"
    }

    if ($productCount -eq 1) {
        $introSolicitud = "su solicitud de asignaci&oacute;n de puntos salariales por el siguiente art&iacute;culo, fue analizada y puesta a consideraci&oacute;n del Comit&eacute;:"
    } else {
        $introSolicitud = "sus solicitudes de asignaci&oacute;n de puntos salariales por los siguientes art&iacute;culos, fueron analizadas y puestas a consideraci&oacute;n del Comit&eacute;:"
    }

    $introParrafo = "Dando cumplimiento a la directriz emitida por la Jefe de la Oficina de Asuntos Profesorales, me permito informarle que, en " + $actaSessionRef + " del Comit&eacute; Interno de Asignaci&oacute;n y Reconocimiento de Puntaje (CIARP)" + $actaFechaText + ", " + $introSolicitud

    return "<html><head><meta charset='utf-8'><style>body{font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;background-color:#f4f6f8;margin:0;padding:20px}.card{max-width:650px;margin:0 auto;background:#fff;border:1px solid #e0e0e0;border-radius:8px;overflow:hidden;box-shadow:0 4px 10px rgba(0,0,0,.05)}.header{background-color:#1b5e20;color:#fff;padding:25px 20px;text-align:center}.header h1{margin:0;font-size:26px;font-weight:700;letter-spacing:1px}.header p{margin:6px 0 0;font-size:13px;opacity:.9}.content{padding:30px;line-height:1.7}.docente-info{background:#f8fafc;border-left:4px solid #1b5e20;padding:15px;margin-bottom:25px;border-radius:0 4px 4px 0}.docente-info p{margin:0;font-size:14px;color:#4a5568}.docente-info .name{font-size:16px;font-weight:700;color:#1b5e20;text-transform:uppercase;margin-bottom:4px}.signature{margin-top:30px;font-size:14px;color:#4a5568;line-height:1.4;border-top:1px solid #e2e8f0;padding-top:20px}</style></head><body><div class='card'><div class='header'><h1>CIARP</h1><p>Comit&eacute; Interno de Asignaci&oacute;n y Reconocimiento de Puntaje</p></div><div class='content'><div class='docente-info'><p class='name'>" + $fullName + "</p><p><strong>Programa:</strong> " + $programa + "</p><p><strong>Facultad:</strong> " + $facultad + "</p></div><p style='font-size:15px;font-weight:700;margin-bottom:15px;color:#2d3748'>Cordial saludo,</p><p style='color:#2d3748;font-size:15px;margin-bottom:20px;text-align:justify;line-height:1.7'>" + $introParrafo + "</p>" + $productsHtml + "<div class='signature'>Atentamente,<br><br><div style='background-color:#f1f8e9;border:1px solid #dcedc8;padding:15px;border-radius:4px;text-align:center;color:#1b5e20;font-weight:700;font-size:16px'>Luz Amparo Celis Buritic&aacute;<br><span style='font-size:14px;font-weight:400;color:#2e7d32'>Jefe Oficina de Asuntos Profesorales</span></div><p style='margin-top:12px;font-size:13px;color:#718096'>Proyect&oacute;: Jos&eacute; Heriberto Vargas Espinosa</p></div></div></div></body></html>"
}

function Build-RowsHtml ($products) {
    $html = ""
    foreach ($p in $products) {
        $html += "<div style='background-color:#e8f5e9;border-left:4px solid #1b5e20;padding:15px;margin-bottom:20px;border-radius:0 4px 4px 0;font-size:15px;color:#2d3748'>" +
                 "<p style='margin-top:0;margin-bottom:15px;line-height:1.5'>" + $p.Detalle + "</p>" +
                 "<p style='margin:0;color:#c62828;text-align:justify;line-height:1.6'>" + $p.Nota + "</p></div>"
    }
    return $html
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

# Load novedades2.xlsx
$wbCiarp    = $excel.Workbooks.Open($ciarpPath)
$allProducts = @()

# Pub_Rev_Pendientes_Homologacion
$shPub = $wbCiarp.Sheets.Item("Pub_Rev_Pendientes_Homologacion")
$r = 4
$maxEmpty = 0
while ($maxEmpty -lt 8) {
    $dni       = Get-CellValue $shPub $r 17
    $tituloArt = Get-CellValue $shPub $r 4
    if (-not $dni -and -not $tituloArt) { $maxEmpty++; $r++; continue }
    $maxEmpty = 0
    
    # Ignorar notas al pie o filas donde el DNI no contenga numeros
    if ($dni -match "[a-zA-Z]" -and $dni -notmatch "\d") {
        $r++; continue
    }
    
    if ($dni) {
        $nombre   = Get-CellValue $shPub $r 18
        $programa = Get-CellValue $shPub $r 22
        $facultad = Get-CellValue $shPub $r 23
        Get-Teacher $dni $nombre $programa $facultad | Out-Null
        $issn     = Get-CellValue $shPub $r 8
        $revista  = Get-CellValue $shPub $r 9
        $fechaPub = Get-CellValue $shPub $r 16
        $acta     = Get-CellValue $shPub $r 26
        
        if ($Sesion -gt 0) {
            if ($acta -notmatch "^$Sesion\s*-" -and $acta -notmatch "No\.\s*$Sesion\b") {
                $r++; continue
            }
        }

        $obs      = Get-CellValue $shPub $r 27
        $concepto = "art&iacute;culo"
        
        # Para Novedades, Detalle incluirá todo el bloque del artículo tal como en la imagen
        $detail   = "&quot;" + $tituloArt + "&quot;<br>Revista: " + $revista + " (ISSN " + $issn + ") &ndash; Publicado: " + $fechaPub
        $allProducts += [PSCustomObject]@{ Dni=($dni -replace '[^\d]',''); Concepto=$concepto; Detalle=$detail; Puntaje="0"; Acta=$acta; Nota=$obs }
    }
    $r++
}

$wbCiarp.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Output "Datos cargados: $($allProducts.Count) novedades en total."

# ---------------------------------------------------------------------------
# FILTER: remove products with zero or empty score
# ---------------------------------------------------------------------------
# Filter valid Novedades (we don't filter by score, we just process all of them)
$validProducts = $allProducts
$excludedProducts = @()

Write-Output "Novedades validas: $($validProducts.Count)"

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

    $productsHtml = Build-RowsHtml $products
    $subject  = "[PRUEBA] Solicitud de asignaci" + $char_o_accent + "n de puntos salariales por art" + $char_i_accent + "culo"
    $htmlBody = Build-EmailHtml $fullName $doc.Programa $doc.Facultad $firstActa $productsHtml $products.Count

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

        $productsHtml = Build-RowsHtml $products
        $subject  = "[PRUEBA - Para: $correoReal] Solicitud de asignaci" + $char_o_accent + "n de puntos salariales por art" + $char_i_accent + "culo"
        $htmlBody = Build-EmailHtml $fullName $programa $facultad $firstActa $productsHtml $products.Count

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

        $productsHtml = Build-RowsHtml $products
        $subject  = "Solicitud de asignaci" + $char_o_accent + "n de puntos salariales por art" + $char_i_accent + "culo"
        $htmlBody = Build-EmailHtml $fullName $programa $facultad $firstActa $productsHtml $products.Count

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
