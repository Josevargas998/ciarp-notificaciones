# Script to generate Outlook drafts for CIARP docente notifications
$ciarpPath = "C:\Users\Admin\Downloads\docentes\ciarp 2.xlsx"
$correosPath = "C:\Users\Admin\Downloads\docentes\correos.xlsx"
$previewDir = "C:\Users\Admin\Downloads\docentes\previews"

if (-not (Test-Path $previewDir)) {
    New-Item -ItemType Directory -Path $previewDir | Out-Null
}

Write-Output "Starting process..."

# 1. Connect to Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# 1.1 Load Docentes Planta (emails)
Write-Output "Loading email directories from correos.xlsx..."
$wbCorreos = $excel.Workbooks.Open($correosPath)
$shPlanta = $wbCorreos.Sheets.Item("Docentes Planta")
$docenteMap = @{} # Key: DNI (clean), Value: @{ Dni, Nombres, Apellidos, Correo, Programa, Facultad }
$r = 2
while ($true) {
    $dni = $shPlanta.Cells.Item($r, 1).Text.Trim()
    if (-not $dni) { break }
    $cleanDni = $dni -replace '[^\d]', ''
    
    $nombres = $shPlanta.Cells.Item($r, 2).Text.Trim()
    $apellidos = $shPlanta.Cells.Item($r, 3).Text.Trim()
    $correo = $shPlanta.Cells.Item($r, 9).Text.Trim()
    
    # We will populate Faculty and Program from product sheets later if not found
    $docenteMap[$cleanDni] = @{
        Dni = $cleanDni
        Nombres = $nombres
        Apellidos = $apellidos
        Correo = $correo
        Programa = ""
        Facultad = ""
    }
    $r++
}
$wbCorreos.Close($false)
Write-Output "Loaded $($docenteMap.Count) teacher records from correos.xlsx."

# Helper to find/create teacher info
function Get-Teacher($dni, $nombreDocente, $programa, $facultad) {
    $cleanDni = $dni -replace '[^\d]', ''
    if (-not $cleanDni) { return $null }
    
    if (-not $docenteMap.ContainsKey($cleanDni)) {
        # Create temp teacher record if not in correos.xlsx
        $docenteMap[$cleanDni] = @{
            Dni = $cleanDni
            Nombres = $nombreDocente
            Apellidos = ""
            Correo = ""
            Programa = $programa
            Facultad = $facultad
        }
    }
    
    $doc = $docenteMap[$cleanDni]
    if (-not $doc.Programa -and $programa) { $doc.Programa = $programa }
    if (-not $doc.Facultad -and $facultad) { $doc.Facultad = $facultad }
    
    return $doc
}

# 2. Parse Products from ciarp 2.xlsx
$wbCiarp = $excel.Workbooks.Open($ciarpPath)
$allProducts = @()

# --- 2.1 Pestaña Titulo ---
Write-Output "Reading sheet: Titulo..."
$shTitulo = $wbCiarp.Sheets.Item("Titulo")
$r = 3
while ($true) {
    $dni = $shTitulo.Cells.Item($r, 4).Text.Trim()
    if (-not $dni) { break }
    
    $nombre = $shTitulo.Cells.Item($r, 5).Text.Trim()
    $programa = $shTitulo.Cells.Item($r, 9).Text.Trim()
    $facultad = $shTitulo.Cells.Item($r, 10).Text.Trim()
    
    $doc = Get-Teacher $dni $nombre $programa $facultad
    
    $univ = $shTitulo.Cells.Item($r, 11).Text.Trim()
    $titulo = $shTitulo.Cells.Item($r, 12).Text.Trim()
    $fechaGrad = $shTitulo.Cells.Item($r, 14).Text.Trim()
    $pts = $shTitulo.Cells.Item($r, 15).Text.Trim()
    $obs = $shTitulo.Cells.Item($r, 17).Text.Trim()
    
    $detail = $titulo + " - " + $univ + " (Fecha de grado: " + $fechaGrad + ")"
    
    $allProducts += [PSCustomObject]@{
        Dni = ($dni -replace '[^\d]', '')
        Concepto = "Título universitario de posgrado"
        Detalle = $detail
        Puntaje = $pts
        Observaciones = $obs
        Acta = $shTitulo.Cells.Item($r, 16).Text.Trim()
    }
    $r++
}

# --- 2.2 Pestaña Pub_Rev_Index ---
Write-Output "Reading sheet: Pub_Rev_Index..."
$shPub = $wbCiarp.Sheets.Item("Pub_Rev_Index")
$r = 3
while ($true) {
    $dni = $shPub.Cells.Item($r, 17).Text.Trim()
    $tituloArt = $shPub.Cells.Item($r, 4).Text.Trim()
    
    if (-not $dni -and -not $tituloArt) {
        $nextDni = $shPub.Cells.Item($r+1, 17).Text.Trim()
        if (-not $nextDni) { break }
    }
    
    if ($dni) {
        $nombre = $shPub.Cells.Item($r, 18).Text.Trim()
        $programa = $shPub.Cells.Item($r, 22).Text.Trim()
        $facultad = $shPub.Cells.Item($r, 23).Text.Trim()
        
        $doc = Get-Teacher $dni $nombre $programa $facultad
        
        $revista = $shPub.Cells.Item($r, 9).Text.Trim()
        $cat = $shPub.Cells.Item($r, 11).Text.Trim()
        $autoresCount = $shPub.Cells.Item($r, 12).Text.Trim()
        $pts = $shPub.Cells.Item($r, 24).Text.Trim()
        
        $concepto = "Artículo en revista indexada"
        $tipo = $shPub.Cells.Item($r, 6).Text.Trim()
        if ($tipo -like "*Editorial*") {
            $concepto = "Editorial en revista indexada"
        }
        
        $obs = ""
        if ($autoresCount) {
            $obs = $autoresCount + " autores."
        }
        # Special note if points is 0.0 and it's because of 2026/date
        if ($pts -eq "0,0" -or $pts -eq "0.0") {
            $fechaPub = $shPub.Cells.Item($r, 14).Text.Trim()
            $obs += " No cumple porque la publicación es de fecha " + $fechaPub + "."
        }
        
        $detail = '"' + $tituloArt + '" - Revista ' + $revista + ' (Categoría ' + $cat + ')'
        
        $allProducts += [PSCustomObject]@{
            Dni = ($dni -replace '[^\d]', '')
            Concepto = $concepto
            Detalle = $detail
            Puntaje = $pts
            Observaciones = $obs
            Acta = "Acta No. 2 del 04 de junio de 2026"
        }
    }
    $r++
}

# --- 2.3 Pestaña Libro_Ensayo ---
Write-Output "Reading sheet: Libro_Ensayo..."
$shEnsayo = $wbCiarp.Sheets.Item("Libro_Ensayo")
$r = 2
while ($true) {
    $dni = $shEnsayo.Cells.Item($r, 8).Text.Trim()
    $nombreLibro = $shEnsayo.Cells.Item($r, 4).Text.Trim()
    
    if (-not $dni -and -not $nombreLibro) { break }
    
    if ($dni) {
        $nombre = $shEnsayo.Cells.Item($r, 9).Text.Trim()
        $programa = $shEnsayo.Cells.Item($r, 13).Text.Trim()
        $facultad = $shEnsayo.Cells.Item($r, 14).Text.Trim()
        
        $doc = Get-Teacher $dni $nombre $programa $facultad
        
        $isbn = $shEnsayo.Cells.Item($r, 5).Text.Trim()
        $pts = $shEnsayo.Cells.Item($r, 19).Text.Trim()
        $obs = $shEnsayo.Cells.Item($r, 21).Text.Trim()
        
        # Check if points is 0.0 due to catedratico
        if (($pts -eq "0,0" -or $pts -eq "0.0") -and ($obs -like "*catedratico*")) {
            # Let it keep the observations from sheet
        }
        
        $detail = '"' + $nombreLibro + '" (ISBN: ' + $isbn + ')'
        
        $allProducts += [PSCustomObject]@{
            Dni = ($dni -replace '[^\d]', '')
            Concepto = "Libro de ensayo"
            Detalle = $detail
            Puntaje = $pts
            Observaciones = $obs
            Acta = $shEnsayo.Cells.Item($r, 20).Text.Trim()
        }
    }
    $r++
}

# --- 2.4 Pestaña Libro_Texto ---
Write-Output "Reading sheet: Libro_Texto..."
$shTexto = $wbCiarp.Sheets.Item("Libro_Texto")
$r = 3
while ($true) {
    $dni = $shTexto.Cells.Item($r, 9).Text.Trim()
    $nombreLibro = $shTexto.Cells.Item($r, 4).Text.Trim()
    
    if (-not $dni -and -not $nombreLibro) { break }
    
    if ($dni) {
        $nombre = $shTexto.Cells.Item($r, 10).Text.Trim()
        $programa = $shTexto.Cells.Item($r, 14).Text.Trim()
        $facultad = $shTexto.Cells.Item($r, 15).Text.Trim()
        
        $doc = Get-Teacher $dni $nombre $programa $facultad
        
        $isbn = $shTexto.Cells.Item($r, 5).Text.Trim()
        $pts = $shTexto.Cells.Item($r, 20).Text.Trim()
        $obs = $shTexto.Cells.Item($r, 22).Text.Trim()
        
        $detail = '"' + $nombreLibro + '" (ISBN: ' + $isbn + ')'
        
        $allProducts += [PSCustomObject]@{
            Dni = ($dni -replace '[^\d]', '')
            Concepto = "Libro de texto"
            Detalle = $detail
            Puntaje = $pts
            Observaciones = $obs
            Acta = $shTexto.Cells.Item($r, 21).Text.Trim()
        }
    }
    $r++
}

# --- 2.5 Pestaña Premios ---
Write-Output "Reading sheet: Premios..."
$shPremios = $wbCiarp.Sheets.Item("Premios")
$r = 3
while ($true) {
    $dni = $shPremios.Cells.Item($r, 4).Text.Trim()
    $trabajo = $shPremios.Cells.Item($r, 11).Text.Trim()
    
    if (-not $dni -and -not $trabajo) { break }
    
    if ($dni) {
        $nombre = $shPremios.Cells.Item($r, 5).Text.Trim()
        $programa = $shPremios.Cells.Item($r, 9).Text.Trim()
        $facultad = $shPremios.Cells.Item($r, 10).Text.Trim()
        
        $doc = Get-Teacher $dni $nombre $programa $facultad
        
        $premio = $shPremios.Cells.Item($r, 12).Text.Trim()
        $entidad = $shPremios.Cells.Item($r, 13).Text.Trim()
        $pts = $shPremios.Cells.Item($r, 15).Text.Trim()
        $obs = $shPremios.Cells.Item($r, 17).Text.Trim()
        
        $detail = '"' + $trabajo + '" - ' + $premio + ', ' + $entidad
        
        $allProducts += [PSCustomObject]@{
            Dni = ($dni -replace '[^\d]', '')
            Concepto = "Premio nacional"
            Detalle = $detail
            Puntaje = $pts
            Observaciones = $obs
            Acta = $shPremios.Cells.Item($r, 16).Text.Trim()
        }
    }
    $r++
}

$wbCiarp.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Output "Loaded $($allProducts.Count) total products."

# 3. Generate Drafts in Outlook
Write-Output "Connecting to Outlook..."
$outlook = New-Object -ComObject Outlook.Application
$session = $outlook.Session
$session.Logon()

# Find account for jhvargas@uniquindio.edu.co
$targetAccount = $null
foreach ($account in $session.Accounts) {
    if ($account.SmtpAddress -eq "jhvargas@uniquindio.edu.co") {
        $targetAccount = $account
        Write-Output "Found matching Outlook account for jhvargas@uniquindio.edu.co"
        break
    }
}

$groupProducts = $allProducts | Group-Object Dni

$processedCount = 0

foreach ($group in $groupProducts) {
    $dni = $group.Name
    $doc = $docenteMap[$dni]
    if (-not $doc) {
        Write-Warning "No metadata found for DNI $dni"
        continue
    }
    
    $products = $group.Group
    
    # Format teacher name
    $fullName = ($doc.Nombres + " " + $doc.Apellidos).Trim()
    if (-not $fullName) {
        $fullName = $products[0].DocenteNombre
    }
    
    # Calculate sum points and format them
    $sum = 0.0
    $rowsHtml = ""
    $firstActa = "Acta No. 2 del 04 de junio de 2026"
    
    # Find list of unique conceptos
    $conceptos = $products | Select-Object -ExpandProperty Concepto -Unique
    $conceptoText = $conceptos -join " y "
    if ($conceptos.Count -eq 1) {
        $subjectConcepto = " por " + $conceptos[0].ToLower()
    } else {
        $subjectConcepto = " por productividad académica"
    }
    
    foreach ($p in $products) {
        # Parse points to float. Note: handle comma and dot decimals
        $ptsStr = $p.Puntaje -replace ',', '.'
        if ($ptsStr -as [double]) {
            $sum += [double]$ptsStr
        }
        
        # Determine Acta if present
        if ($p.Acta) { $firstActa = $p.Acta }
        
        $rowsHtml += @"
        <tr style="border-bottom: 1px solid #e0e0e0;">
            <td style="padding: 12px; font-size: 14px; color: #333; font-weight: bold;">$($p.Concepto)</td>
            <td style="padding: 12px; font-size: 14px; color: #555;">$($p.Detalle)</td>
            <td style="padding: 12px; font-size: 14px; color: #333; text-align: right; font-weight: bold;">$($p.Puntaje) puntos</td>
            <td style="padding: 12px; font-size: 13px; color: #666; font-style: italic;">$($p.Observaciones)</td>
        </tr>
"@
    }
    
    # Format sum to display nicely
    $sumStr = $sum.ToString("F1") -replace '\.', ','
    
    $subject = "Respuesta a solicitud de asignación de puntos salariales$subjectConcepto"
    
    # Build complete HTML body
    $htmlBody = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f6f8; margin: 0; padding: 20px; }
        .card { max-width: 650px; margin: 0 auto; background: #ffffff; border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 10px rgba(0,0,0,0.05); }
        .header { background-color: #1b5e20; color: #ffffff; padding: 25px 20px; text-align: center; }
        .header h1 { margin: 0; font-size: 26px; font-weight: 700; letter-spacing: 1px; }
        .header p { margin: 6px 0 0 0; font-size: 13px; opacity: 0.9; }
        .content { padding: 30px; color: #2d3748; line-height: 1.6; }
        .docente-info { background: #f8fafc; border-left: 4px solid #1b5e20; padding: 15px; margin-bottom: 25px; border-radius: 0 4px 4px 0; }
        .docente-info p { margin: 0; font-size: 14px; color: #4a5568; }
        .docente-info .name { font-size: 16px; font-weight: bold; color: #1b5e20; text-transform: uppercase; margin-bottom: 4px; }
        .greeting { font-size: 15px; font-weight: bold; margin-bottom: 15px; }
        .table-container { margin: 20px 0; overflow-x: auto; }
        table { width: 100%; border-collapse: collapse; border: 1px solid #e0e0e0; }
        th { background-color: #1b5e20; color: #ffffff; padding: 12px; font-size: 14px; text-align: left; }
        .total-box { background-color: #e8f5e9; border: 1px solid #c8e6c9; padding: 12px 15px; border-radius: 4px; font-size: 15px; font-weight: bold; color: #2e7d32; margin-top: 15px; }
        .signature { margin-top: 30px; font-size: 14px; color: #4a5568; line-height: 1.4; border-top: 1px solid #e2e8f0; padding-top: 20px; }
    </style>
</head>
<body>
    <div class="card">
        <div class="header">
            <h1>CIARP</h1>
            <p>Comité Interno de Asignación y Reconocimiento de Puntaje</p>
        </div>
        <div class="content">
            <h2 style="color: #1b5e20; text-align: center; margin-top: 0; margin-bottom: 25px; font-size: 20px;">Reconocimiento de Puntos Salariales</h2>
            
            <div class="docente-info">
                <p class="name">$fullName</p>
                <p><strong>Programa:</strong> $($doc.Programa)</p>
                <p><strong>Facultad:</strong> $($doc.Facultad)</p>
            </div>
            
            <p class="greeting">Cordial saludo,</p>
            <p>Dando cumplimiento a la directriz emitida por la jefe de la Oficina de Asuntos Profesionales, me permito informarle que, en sesión del Comité Interno de Asignación y Reconocimiento de Puntaje (CIARP), llevada a cabo según consta en <strong>$firstActa</strong>, la solicitud de asignación de puntos salariales presentada por usted fue considerada y aprobada como se relaciona a continuación:</p>
            
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th style="width: 25%;">Concepto</th>
                            <th style="width: 40%;">Detalle del Producto</th>
                            <th style="width: 15%; text-align: right;">Puntaje</th>
                            <th style="width: 20%;">Observaciones</th>
                        </tr>
                    </thead>
                    <tbody>
                        $rowsHtml
                    </tbody>
                </table>
            </div>
            
            <div class="total-box">
                Puntaje total aprobado en esta sesión: $sumStr puntos
            </div>
            
            <p style="margin-top: 20px;">Por lo anteriormente mencionado, con posterioridad se estará notificando el correspondiente acto administrativo.</p>
            <p>Cualquier inquietud al respecto, estaremos atentos.</p>
            
            <div class="signature">
                Atentamente,<br><br>
                <strong>José Heriberto Vargas Espinosa</strong><br>
                Contratista<br>
                Oficina de Asuntos Profesorales<br>
                tel.: (606) 7359300 ext.: 843<br>
                e-mail: jhvargas@uniquindio.edu.co
            </div>
        </div>
    </div>
</body>
</html>
"@

    # Save local preview html
    $safeName = $fullName -replace '[^\w\s-]', ''
    $safeName = $safeName -replace '\s+', '_'
    $previewPath = Join-Path $previewDir "$safeName.html"
    $htmlBody | Out-File $previewPath -Encoding utf8
    
    # Create Outlook draft
    $mail = $outlook.CreateItem(0) # 0 = olMailItem
    $mail.Subject = $subject
    $mail.HTMLBody = $htmlBody
    $mail.To = $doc.Correo
    
    # Apply account sender
    if ($targetAccount) {
        $mail.SendUsingAccount = $targetAccount
    }
    $mail.SentOnBehalfOfName = "jhvargas@uniquindio.edu.co"
    
    $mail.Save()
    
    $processedCount++
    Write-Output "Created draft and preview for: $fullName ($($doc.Correo))"
}

Write-Output "Done! Generated $processedCount drafts in Outlook and HTML previews in: $previewDir"
