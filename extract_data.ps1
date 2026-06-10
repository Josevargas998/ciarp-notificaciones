$ciarpPath = "C:\Users\Admin\Downloads\docentes\ciarp 2.xlsx"
$correosPath = "C:\Users\Admin\Downloads\docentes\correos.xlsx"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# 1. Load Emails
$wbCorreos = $excel.Workbooks.Open($correosPath)
$shPlanta = $wbCorreos.Sheets.Item("Docentes Planta")
$emails = @{}
$docenteData = @{} # Key: Documento, Value: @{ Nombre, Apellido, Correo, Programa, Facultad }
$r = 2
while ($true) {
    $dni = $shPlanta.Cells.Item($r, 1).Text.Trim()
    if (-not $dni) { break }
    
    # Standardize DNI to remove dots and spaces if any
    $cleanDni = $dni -replace '[^\d]', ''
    
    $nombres = $shPlanta.Cells.Item($r, 2).Text.Trim()
    $apellidos = $shPlanta.Cells.Item($r, 3).Text.Trim()
    $cargoCat = $shPlanta.Cells.Item($r, 4).Text.Trim()
    $catDed = $shPlanta.Cells.Item($r, 5).Text.Trim()
    $correo = $shPlanta.Cells.Item($r, 9).Text.Trim()
    
    $emails[$cleanDni] = $correo
    $docenteData[$cleanDni] = @{
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

Write-Output "Loaded $($emails.Count) emails from correos.xlsx."

# Helper to add or update docente
function Get-Or-Create-Docente($dni, $nombreDocente, $programa, $facultad) {
    $cleanDni = $dni -replace '[^\d]', ''
    if (-not $cleanDni) { return $null }
    
    if (-not $docenteData.ContainsKey($cleanDni)) {
        # Split names
        $parts = $nombreDocente -split ' '
        $nombres = ""
        $apellidos = ""
        if ($parts.Count -ge 2) {
            $nombres = $parts[0..(($parts.Count/2)-1)] -join " "
            $apellidos = $parts[($parts.Count/2)..($parts.Count-1)] -join " "
        } else {
            $nombres = $nombreDocente
        }
        
        $docenteData[$cleanDni] = @{
            Dni = $cleanDni
            Nombres = $nombres
            Apellidos = $apellidos
            Correo = ""
            Programa = $programa
            Facultad = $facultad
        }
    }
    
    $doc = $docenteData[$cleanDni]
    if (-not $doc.Programa -and $programa) { $doc.Programa = $programa }
    if (-not $doc.Facultad -and $facultad) { $doc.Facultad = $facultad }
    
    return $doc
}

# 2. Load Products from ciarp 2.xlsx
$wbCiarp = $excel.Workbooks.Open($ciarpPath)

$allProducts = @() # Array of custom objects for each product

# --- 2.1 Pestana Titulo ---
Write-Output "Processing Titulo..."
$shTitulo = $wbCiarp.Sheets.Item("Titulo")
$r = 3
while ($true) {
    $dni = $shTitulo.Cells.Item($r, 4).Text.Trim()
    if (-not $dni) { break }
    
    $nombre = $shTitulo.Cells.Item($r, 5).Text.Trim()
    $programa = $shTitulo.Cells.Item($r, 9).Text.Trim()
    $facultad = $shTitulo.Cells.Item($r, 10).Text.Trim()
    
    $doc = Get-Or-Create-Docente $dni $nombre $programa $facultad
    
    $univ = $shTitulo.Cells.Item($r, 11).Text.Trim()
    $titulo = $shTitulo.Cells.Item($r, 12).Text.Trim()
    $fechaGrad = $shTitulo.Cells.Item($r, 14).Text.Trim()
    $pts = $shTitulo.Cells.Item($r, 15).Text.Trim()
    $obs = $shTitulo.Cells.Item($r, 17).Text.Trim()
    
    $detail = $titulo + " - " + $univ + " (Fecha de grado: " + $fechaGrad + ")"
    
    $allProducts += [PSCustomObject]@{
        Dni = ($dni -replace '[^\d]', '')
        DocenteNombre = $nombre
        Concepto = "Titulo universitario de posgrado"
        Detalle = $detail
        Puntaje = $pts
        Observaciones = $obs
        Acta = $shTitulo.Cells.Item($r, 16).Text.Trim()
        Sheet = "Titulo"
    }
    $r++
}

# --- 2.2 Pestana Pub_Rev_Index ---
Write-Output "Processing Pub_Rev_Index..."
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
        
        $doc = Get-Or-Create-Docente $dni $nombre $programa $facultad
        
        $revista = $shPub.Cells.Item($r, 9).Text.Trim()
        $cat = $shPub.Cells.Item($r, 11).Text.Trim()
        $autoresCount = $shPub.Cells.Item($r, 12).Text.Trim()
        $pts = $shPub.Cells.Item($r, 24).Text.Trim()
        
        $concepto = "Articulo en revista indexada"
        $tipo = $shPub.Cells.Item($r, 6).Text.Trim()
        if ($tipo -like "*Editorial*") {
            $concepto = "Editorial en revista indexada"
        }
        
        $obs = ""
        if ($autoresCount) {
            $obs = $autoresCount + " autores."
        }
        
        $detail = '"' + $tituloArt + '" - Revista ' + $revista + ' (Categoria ' + $cat + ')'
        
        $allProducts += [PSCustomObject]@{
            Dni = ($dni -replace '[^\d]', '')
            DocenteNombre = $nombre
            Concepto = $concepto
            Detalle = $detail
            Puntaje = $pts
            Observaciones = $obs
            Acta = "Acta No. 2 del 04 de junio de 2026"
            Sheet = "Pub_Rev_Index"
        }
    }
    $r++
}

# --- 2.3 Pestana Libro_Ensayo ---
Write-Output "Processing Libro_Ensayo..."
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
        
        $doc = Get-Or-Create-Docente $dni $nombre $programa $facultad
        
        $isbn = $shEnsayo.Cells.Item($r, 5).Text.Trim()
        $pts = $shEnsayo.Cells.Item($r, 19).Text.Trim()
        $obs = $shEnsayo.Cells.Item($r, 21).Text.Trim()
        
        $detail = '"' + $nombreLibro + '" (ISBN: ' + $isbn + ')'
        
        $allProducts += [PSCustomObject]@{
            Dni = ($dni -replace '[^\d]', '')
            DocenteNombre = $nombre
            Concepto = "Libro de ensayo"
            Detalle = $detail
            Puntaje = $pts
            Observaciones = $obs
            Acta = $shEnsayo.Cells.Item($r, 20).Text.Trim()
            Sheet = "Libro_Ensayo"
        }
    }
    $r++
}

# --- 2.4 Pestana Libro_Texto ---
Write-Output "Processing Libro_Texto..."
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
        
        $doc = Get-Or-Create-Docente $dni $nombre $programa $facultad
        
        $isbn = $shTexto.Cells.Item($r, 5).Text.Trim()
        $pts = $shTexto.Cells.Item($r, 20).Text.Trim()
        $obs = $shTexto.Cells.Item($r, 22).Text.Trim()
        
        $detail = '"' + $nombreLibro + '" (ISBN: ' + $isbn + ')'
        
        $allProducts += [PSCustomObject]@{
            Dni = ($dni -replace '[^\d]', '')
            DocenteNombre = $nombre
            Concepto = "Libro de texto"
            Detalle = $detail
            Puntaje = $pts
            Observaciones = $obs
            Acta = $shTexto.Cells.Item($r, 21).Text.Trim()
            Sheet = "Libro_Texto"
        }
    }
    $r++
}

# --- 2.5 Pestana Premios ---
Write-Output "Processing Premios..."
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
        
        $doc = Get-Or-Create-Docente $dni $nombre $programa $facultad
        
        $premio = $shPremios.Cells.Item($r, 12).Text.Trim()
        $entidad = $shPremios.Cells.Item($r, 13).Text.Trim()
        $pts = $shPremios.Cells.Item($r, 15).Text.Trim()
        $obs = $shPremios.Cells.Item($r, 17).Text.Trim()
        
        $detail = '"' + $trabajo + '" - ' + $premio + ', ' + $entidad
        
        $allProducts += [PSCustomObject]@{
            Dni = ($dni -replace '[^\d]', '')
            DocenteNombre = $nombre
            Concepto = "Premio nacional"
            Detalle = $detail
            Puntaje = $pts
            Observaciones = $obs
            Acta = $shPremios.Cells.Item($r, 16).Text.Trim()
            Sheet = "Premios"
        }
    }
    $r++
}

$wbCiarp.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Output "Loaded $($allProducts.Count) product entries."

$docentesList = @()
foreach ($dni in $docenteData.Keys) {
    $doc = $docenteData[$dni]
    $myProducts = $allProducts | Where-Object { $_.Dni -eq $dni }
    if ($myProducts.Count -gt 0) {
        $docentesList += [PSCustomObject]@{
            Dni = $doc.Dni
            Nombres = $doc.Nombres
            Apellidos = $doc.Apellidos
            Correo = $doc.Correo
            Programa = $doc.Programa
            Facultad = $doc.Facultad
            Products = $myProducts
        }
    }
}

$docentesList | ConvertTo-Json -Depth 5 | Out-File "C:\Users\Admin\Downloads\docentes\parsed_docentes.json" -Encoding utf8
Write-Output "Successfully dumped parsed data to parsed_docentes.json."
