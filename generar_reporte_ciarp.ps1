# Genera reporte completo en Excel Y Word del Comite CIARP 2 - Acta 2
$ciarpPath   = "C:\Users\Admin\Downloads\docentes\ciarp 2.xlsx"
$reportDir   = "C:\Users\Admin\Downloads\docentes"
$timestamp   = Get-Date -Format "yyyyMMdd_HHmmss"
$excelOut    = "$reportDir\ResumenCIARP2_$timestamp.xlsx"
$wordOut     = "$reportDir\ResumenCIARP2_$timestamp.docx"

# ---------- Helpers ----------
function Get-CellValue($sheet, $row, $col) {
    $cell = $sheet.Cells.Item($row, $col)
    if ($cell.MergeCells) { return $cell.MergeArea.Item(1,1).Text.Trim() }
    return $cell.Text.Trim()
}
function Parse-Number($txt) {
    $t = $txt -replace ',','.'
    $n = 0.0
    if ([double]::TryParse($t, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$n)) { return $n }
    return 0.0
}

# ---------- Read CIARP Excel ----------
$srcExcel = New-Object -ComObject Excel.Application
$srcExcel.Visible = $false
$srcExcel.DisplayAlerts = $false
$wb = $srcExcel.Workbooks.Open($ciarpPath)

$allRows = @()

# TITULO
$sh = $wb.Sheets.Item("Titulo")
$r = 3; $emp = 0
while ($emp -lt 8) {
    $dni = Get-CellValue $sh $r 4
    if (-not $dni) { $emp++; $r++; continue }; $emp = 0
    $cleanDni = $dni -replace '[^\d]',''
    $allRows += [PSCustomObject]@{
        Cedula=$cleanDni; Nombre=(Get-CellValue $sh $r 5); Tipo="Titulo posgrado"
        Producto=(Get-CellValue $sh $r 12); Puntaje=(Parse-Number (Get-CellValue $sh $r 15))
        Acta=(Get-CellValue $sh $r 16)
    }
    $r++
}

# PUB_REV_INDEX
$sh = $wb.Sheets.Item("Pub_Rev_Index")
$r = 3; $emp = 0
while ($emp -lt 8) {
    $dni = Get-CellValue $sh $r 17; $art = Get-CellValue $sh $r 4
    if (-not $dni -and -not $art) { $emp++; $r++; continue }; $emp = 0
    if ($dni) {
        $cleanDni = $dni -replace '[^\d]',''
        $allRows += [PSCustomObject]@{
            Cedula=$cleanDni; Nombre=(Get-CellValue $sh $r 18); Tipo="Articulo revista indexada"
            Producto=$art; Puntaje=(Parse-Number (Get-CellValue $sh $r 24))
            Acta="2 - 04/06/2026"
        }
    }
    $r++
}

# LIBRO_ENSAYO
$sh = $wb.Sheets.Item("Libro_Ensayo")
$r = 2; $emp = 0
while ($emp -lt 8) {
    $dni = Get-CellValue $sh $r 8; $lib = Get-CellValue $sh $r 4
    if (-not $dni -and -not $lib) { $emp++; $r++; continue }; $emp = 0
    if ($dni) {
        $cleanDni = $dni -replace '[^\d]',''
        $allRows += [PSCustomObject]@{
            Cedula=$cleanDni; Nombre=(Get-CellValue $sh $r 9); Tipo="Libro de ensayo"
            Producto=$lib; Puntaje=(Parse-Number (Get-CellValue $sh $r 19))
            Acta=(Get-CellValue $sh $r 20)
        }
    }
    $r++
}

# LIBRO_TEXTO
$sh = $wb.Sheets.Item("Libro_Texto")
$r = 3; $emp = 0
while ($emp -lt 8) {
    $dni = Get-CellValue $sh $r 9; $lib = Get-CellValue $sh $r 4
    if (-not $dni -and -not $lib) { $emp++; $r++; continue }; $emp = 0
    if ($dni) {
        $cleanDni = $dni -replace '[^\d]',''
        $allRows += [PSCustomObject]@{
            Cedula=$cleanDni; Nombre=(Get-CellValue $sh $r 10); Tipo="Libro de texto"
            Producto=$lib; Puntaje=(Parse-Number (Get-CellValue $sh $r 20))
            Acta=(Get-CellValue $sh $r 21)
        }
    }
    $r++
}

# PREMIOS
$sh = $wb.Sheets.Item("Premios")
$r = 3; $emp = 0
while ($emp -lt 8) {
    $dni = Get-CellValue $sh $r 4; $trab = Get-CellValue $sh $r 11
    if (-not $dni -and -not $trab) { $emp++; $r++; continue }; $emp = 0
    if ($dni) {
        $cleanDni = $dni -replace '[^\d]',''
        $allRows += [PSCustomObject]@{
            Cedula=$cleanDni; Nombre=(Get-CellValue $sh $r 5); Tipo="Premio nacional"
            Producto=$trab; Puntaje=(Parse-Number (Get-CellValue $sh $r 15))
            Acta=(Get-CellValue $sh $r 16)
        }
    }
    $r++
}

$wb.Close($false)
$srcExcel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($srcExcel) | Out-Null

Write-Host "Productos leidos: $($allRows.Count)"

# Separar notificados vs no
$notificados = $allRows | Where-Object { $_.Puntaje -gt 0 }
$noNotificados = $allRows | Where-Object { $_.Puntaje -le 0 }

# Docentes unicos notificados
$docentesNotif = $notificados | Group-Object Cedula | ForEach-Object {
    $g = $_.Group
    $tipos = ($g | Select-Object -ExpandProperty Tipo -Unique) -join ", "
    $totalPts = ($g | Measure-Object Puntaje -Sum).Sum
    [PSCustomObject]@{
        N = 0; Cedula=$g[0].Cedula; Nombre=$g[0].Nombre
        TipoProducto=$tipos; CantidadProductos=$g.Count
        PuntajeTotal=[math]::Round($totalPts,1)
        Estado="Notificado"
    }
} | Sort-Object Nombre
$i = 1; foreach ($d in $docentesNotif) { $d.N = $i++ }

$docentesNoNotif = $noNotificados | Group-Object Cedula | ForEach-Object {
    $g = $_.Group
    $tipos = ($g | Select-Object -ExpandProperty Tipo -Unique) -join ", "
    [PSCustomObject]@{
        N = 0; Cedula=$g[0].Cedula; Nombre=$g[0].Nombre
        TipoProducto=$tipos; CantidadProductos=$g.Count
        PuntajeTotal=0
        Estado="No notificado (puntaje 0)"
    }
} | Sort-Object Nombre
$i = 1; foreach ($d in $docentesNoNotif) { $d.N = $i++ }

# ========== EXCEL ==========
Write-Host "Generando Excel..."
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xwb = $xl.Workbooks.Add()

# Colores
$verde      = 0x00C000   # verde oscuro
$verdeClaro = 0xC6EFCE   # verde claro fondo
$rojo       = 0xC00000
$rojoClaro  = 0xFFC7CE
$azul       = 0x1F497D
$azulClaro  = 0xDEE6EF
$gris       = 0xD9D9D9
$blanco     = 0xFFFFFF
$amarillo   = 0xFFEB9C

function Set-Header($ws, $row, $headers, $bgColor, $fgColor=0xFFFFFF) {
    for ($c=1; $c -le $headers.Count; $c++) {
        $cell = $ws.Cells.Item($row, $c)
        $cell.Value2 = $headers[$c-1]
        $cell.Font.Bold = $true
        $cell.Font.Color = $fgColor
        $cell.Interior.Color = $bgColor
        $cell.HorizontalAlignment = -4108  # xlCenter
        $cell.VerticalAlignment   = -4108
        $cell.WrapText = $true
    }
}

# ---- Hoja 1: Portada / Resumen ----
$ws1 = $xwb.Sheets.Item(1)
$ws1.Name = "Resumen"

# Titulo principal
$ws1.Rows.Item(1).RowHeight = 35
$ws1.Range("A1:F1").Merge()
$ws1.Cells.Item(1,1).Value2 = "COMITE INTERNO DE ASIGNACION Y RECONOCIMIENTO DE PUNTAJE - CIARP"
$ws1.Cells.Item(1,1).Font.Bold = $true
$ws1.Cells.Item(1,1).Font.Size = 14
$ws1.Cells.Item(1,1).Font.Color = $blanco
$ws1.Cells.Item(1,1).Interior.Color = $azul
$ws1.Cells.Item(1,1).HorizontalAlignment = -4108

$ws1.Rows.Item(2).RowHeight = 25
$ws1.Range("A2:F2").Merge()
$ws1.Cells.Item(2,1).Value2 = "Universidad del Quindio  -  Acta No. 2 del 04 de junio de 2026"
$ws1.Cells.Item(2,1).Font.Bold = $true
$ws1.Cells.Item(2,1).Font.Size = 12
$ws1.Cells.Item(2,1).Font.Color = $azul
$ws1.Cells.Item(2,1).HorizontalAlignment = -4108

# Stats
$ws1.Cells.Item(4,1).Value2 = "Fecha de generacion:"
$ws1.Cells.Item(4,2).Value2 = (Get-Date -Format "dd/MM/yyyy HH:mm")
$ws1.Cells.Item(5,1).Value2 = "Total docentes:"
$ws1.Cells.Item(5,2).Value2 = ($docentesNotif.Count + $docentesNoNotif.Count)
$ws1.Cells.Item(6,1).Value2 = "Notificados (puntaje > 0):"
$ws1.Cells.Item(6,2).Value2 = $docentesNotif.Count
$ws1.Cells.Item(6,2).Font.Color = 0x006100
$ws1.Cells.Item(7,1).Value2 = "No notificados (puntaje 0):"
$ws1.Cells.Item(7,2).Value2 = $docentesNoNotif.Count
$ws1.Cells.Item(7,2).Font.Color = 0x9C0006
$ws1.Cells.Item(8,1).Value2 = "Total productos presentados:"
$ws1.Cells.Item(8,2).Value2 = $allRows.Count

foreach ($r in 4..8) { $ws1.Cells.Item($r,1).Font.Bold = $true }

# Tabla resumen por tipo
$ws1.Cells.Item(10,1).Value2 = "Tipo de Producto"
$ws1.Cells.Item(10,2).Value2 = "Docentes"
$ws1.Cells.Item(10,3).Value2 = "Productos"
$ws1.Cells.Item(10,4).Value2 = "Puntaje Promedio"
foreach ($c in 1..4) {
    $ws1.Cells.Item(10,$c).Font.Bold = $true
    $ws1.Cells.Item(10,$c).Interior.Color = $azul
    $ws1.Cells.Item(10,$c).Font.Color = $blanco
}

$tiposStats = $allRows | Group-Object Tipo | ForEach-Object {
    $g = $_.Group
    $withPts = $g | Where-Object { $_.Puntaje -gt 0 }
    $docCount = ($withPts | Select-Object -Property Cedula -Unique | Measure-Object).Count
    $avg = if ($withPts.Count -gt 0) { [math]::Round((($withPts | Measure-Object Puntaje -Sum).Sum / $withPts.Count),1) } else { 0 }
    [PSCustomObject]@{ Tipo=$_.Name; Docentes=$docCount; Productos=$g.Count; Promedio=$avg }
}

$row = 11
foreach ($t in $tiposStats) {
    $ws1.Cells.Item($row,1).Value2 = $t.Tipo
    $ws1.Cells.Item($row,2).Value2 = $t.Docentes
    $ws1.Cells.Item($row,3).Value2 = $t.Productos
    $ws1.Cells.Item($row,4).Value2 = $t.Promedio
    if ($row % 2 -eq 0) {
        foreach ($c in 1..4) { $ws1.Cells.Item($row,$c).Interior.Color = $azulClaro }
    }
    $row++
}
# Total
$ws1.Cells.Item($row,1).Value2 = "TOTAL"
$ws1.Cells.Item($row,2).Value2 = ($docentesNotif.Count + $docentesNoNotif.Count)
$ws1.Cells.Item($row,3).Value2 = $allRows.Count
foreach ($c in 1..4) {
    $ws1.Cells.Item($row,$c).Font.Bold = $true
    $ws1.Cells.Item($row,$c).Interior.Color = $gris
}
$ws1.Columns.AutoFit() | Out-Null

# ---- Hoja 2: Notificados ----
$ws2 = $xwb.Sheets.Add([System.Reflection.Missing]::Value, $xwb.Sheets.Item($xwb.Sheets.Count))
$ws2.Name = "Notificados"

$hdrs2 = @("#","Cedula","Nombre del Docente","Tipo de Producto","No. Productos","Puntaje Total","Estado")
Set-Header $ws2 1 $hdrs2 0x006100

$r = 2
foreach ($d in $docentesNotif) {
    $ws2.Cells.Item($r,1).Value2 = $d.N
    $ws2.Cells.Item($r,2).Value2 = $d.Cedula
    $ws2.Cells.Item($r,3).Value2 = $d.Nombre
    $ws2.Cells.Item($r,4).Value2 = $d.TipoProducto
    $ws2.Cells.Item($r,5).Value2 = $d.CantidadProductos
    $ws2.Cells.Item($r,6).Value2 = $d.PuntajeTotal
    $ws2.Cells.Item($r,7).Value2 = $d.Estado
    $ws2.Cells.Item($r,7).Font.Color = 0x006100
    $ws2.Cells.Item($r,7).Font.Bold = $true
    if ($r % 2 -eq 0) { foreach ($c in 1..7) { $ws2.Cells.Item($r,$c).Interior.Color = $verdeClaro } }
    $r++
}
$ws2.ListObjects.Add(1, $ws2.Range("A1:G$($r-1)"), [System.Reflection.Missing]::Value, 1) | Out-Null
$ws2.Columns.AutoFit() | Out-Null

# ---- Hoja 3: No Notificados ----
$ws3 = $xwb.Sheets.Add([System.Reflection.Missing]::Value, $xwb.Sheets.Item($xwb.Sheets.Count))
$ws3.Name = "No Notificados"

$hdrs3 = @("#","Cedula","Nombre del Docente","Tipo de Producto","No. Productos","Puntaje","Estado")
Set-Header $ws3 1 $hdrs3 0x9C0006

$r = 2
foreach ($d in $docentesNoNotif) {
    $ws3.Cells.Item($r,1).Value2 = $d.N
    $ws3.Cells.Item($r,2).Value2 = $d.Cedula
    $ws3.Cells.Item($r,3).Value2 = $d.Nombre
    $ws3.Cells.Item($r,4).Value2 = $d.TipoProducto
    $ws3.Cells.Item($r,5).Value2 = $d.CantidadProductos
    $ws3.Cells.Item($r,6).Value2 = 0
    $ws3.Cells.Item($r,7).Value2 = $d.Estado
    $ws3.Cells.Item($r,7).Font.Color = 0x9C0006
    $ws3.Cells.Item($r,7).Font.Bold = $true
    if ($r % 2 -eq 0) { foreach ($c in 1..7) { $ws3.Cells.Item($r,$c).Interior.Color = $rojoClaro } }
    $r++
}
$ws3.ListObjects.Add(1, $ws3.Range("A1:G$($r-1)"), [System.Reflection.Missing]::Value, 1) | Out-Null
$ws3.Columns.AutoFit() | Out-Null

# ---- Hoja 4: Todos los Productos ----
$ws4 = $xwb.Sheets.Add([System.Reflection.Missing]::Value, $xwb.Sheets.Item($xwb.Sheets.Count))
$ws4.Name = "Todos los Productos"

$hdrs4 = @("#","Cedula","Nombre del Docente","Tipo","Titulo del Producto","Puntaje","Acta")
Set-Header $ws4 1 $hdrs4 $azul

$r = 2; $n = 1
$sortedAll = $allRows | Sort-Object Nombre
foreach ($p in $sortedAll) {
    $ws4.Cells.Item($r,1).Value2 = $n++
    $ws4.Cells.Item($r,2).Value2 = $p.Cedula
    $ws4.Cells.Item($r,3).Value2 = $p.Nombre
    $ws4.Cells.Item($r,4).Value2 = $p.Tipo
    $ws4.Cells.Item($r,5).Value2 = $p.Producto
    $ws4.Cells.Item($r,6).Value2 = $p.Puntaje
    $ws4.Cells.Item($r,7).Value2 = $p.Acta
    # Color segun puntaje
    $bgc = if ($p.Puntaje -gt 0) { $verdeClaro } else { $rojoClaro }
    foreach ($c in 1..7) { $ws4.Cells.Item($r,$c).Interior.Color = $bgc }
    $r++
}
$ws4.ListObjects.Add(1, $ws4.Range("A1:G$($r-1)"), [System.Reflection.Missing]::Value, 1) | Out-Null
$ws4.Columns.AutoFit() | Out-Null

# Ir a la primera hoja
$xwb.Sheets.Item("Resumen").Activate()
$xwb.SaveAs($excelOut)
$xwb.Close($false)
$xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
Write-Host "Excel generado: $excelOut"

# ========== WORD ==========
Write-Host "Generando Word..."
$wd = New-Object -ComObject Word.Application
$wd.Visible = $false
$wdoc = $wd.Documents.Add()
$wsel = $wd.Selection

# Estilos
$wsel.Style = "Titulo"
$wsel.TypeText("COMITE INTERNO DE ASIGNACION Y RECONOCIMIENTO DE PUNTAJE")
$wsel.TypeParagraph()
$wsel.Style = "Normal"

$wsel.Style = "Titulo 1"
$wsel.TypeText("Universidad del Quindio")
$wsel.TypeParagraph()

$wsel.Style = "Normal"
$wsel.TypeText("Acta No. 2 | Fecha: 04 de junio de 2026")
$wsel.TypeParagraph()
$wsel.TypeText("Reporte generado: " + (Get-Date -Format "dd/MM/yyyy HH:mm"))
$wsel.TypeParagraph()
$wsel.TypeParagraph()

# Seccion 1
$wsel.Style = "Titulo 2"
$wsel.TypeText("1. Resumen General")
$wsel.TypeParagraph()
$wsel.Style = "Normal"
$wsel.TypeText("En el Comite CIARP No. 2 del 04 de junio de 2026 se evaluaron los productos academicos presentados por los docentes de planta de la Universidad del Quindio.")
$wsel.TypeParagraph()
$wsel.TypeParagraph()

# Tabla resumen numeros
$tbl0 = $wdoc.Tables.Add($wsel.Range, 5, 2)
$tbl0.Style = "Tabla con cuadricula"
$items = @(
    @("Total de docentes evaluados", ($docentesNotif.Count + $docentesNoNotif.Count)),
    @("Docentes con puntaje asignado (notificados)", $docentesNotif.Count),
    @("Docentes con puntaje cero (no notificados)", $docentesNoNotif.Count),
    @("Total de productos presentados", $allRows.Count),
    @("Fecha del acta", "04 de junio de 2026")
)
for ($ri=0; $ri -lt $items.Count; $ri++) {
    $tbl0.Cell($ri+1,1).Range.Text = $items[$ri][0]
    $tbl0.Cell($ri+1,2).Range.Text = [string]$items[$ri][1]
    $tbl0.Cell($ri+1,1).Range.Bold = 1
}
$wsel.EndOf(6) | Out-Null  # wdStory
$wsel.TypeParagraph()
$wsel.TypeParagraph()

# Seccion 2: Notificados
$wsel.Style = "Titulo 2"
$wsel.TypeText("2. Docentes Notificados (" + $docentesNotif.Count + ")")
$wsel.TypeParagraph()
$wsel.Style = "Normal"
$wsel.TypeText("Los siguientes docentes recibieron notificacion de los puntos asignados:")
$wsel.TypeParagraph()

$tbl1 = $wdoc.Tables.Add($wsel.Range, ($docentesNotif.Count + 1), 5)
$tbl1.Style = "Tabla con cuadricula"
$hdN = @("#","Cedula","Nombre","Tipo de Producto","Puntaje Total")
for ($c=0; $c -lt 5; $c++) {
    $tbl1.Cell(1,$c+1).Range.Text = $hdN[$c]
    $tbl1.Cell(1,$c+1).Range.Bold = 1
    $tbl1.Cell(1,$c+1).Shading.BackgroundPatternColor = 0x1F497D
    $tbl1.Cell(1,$c+1).Range.Font.ColorIndex = 2  # white
}
for ($ri=0; $ri -lt $docentesNotif.Count; $ri++) {
    $d = $docentesNotif[$ri]
    $tbl1.Cell($ri+2,1).Range.Text = [string]$d.N
    $tbl1.Cell($ri+2,2).Range.Text = $d.Cedula
    $tbl1.Cell($ri+2,3).Range.Text = $d.Nombre
    $tbl1.Cell($ri+2,4).Range.Text = $d.TipoProducto
    $tbl1.Cell($ri+2,5).Range.Text = [string]$d.PuntajeTotal
}
$tbl1.Columns.AutoFit() | Out-Null

$wsel.EndOf(6) | Out-Null
$wsel.TypeParagraph()
$wsel.TypeParagraph()

# Seccion 3: No notificados
$wsel.Style = "Titulo 2"
$wsel.TypeText("3. Docentes NO Notificados - Puntaje Cero (" + $docentesNoNotif.Count + ")")
$wsel.TypeParagraph()
$wsel.Style = "Normal"
$wsel.TypeText("Los siguientes docentes presentaron productos con puntaje 0 o no asignado, por lo que no fueron notificados:")
$wsel.TypeParagraph()

$tbl2 = $wdoc.Tables.Add($wsel.Range, ($docentesNoNotif.Count + 1), 4)
$tbl2.Style = "Tabla con cuadricula"
$hdNn = @("#","Cedula","Nombre","Tipo de Producto")
for ($c=0; $c -lt 4; $c++) {
    $tbl2.Cell(1,$c+1).Range.Text = $hdNn[$c]
    $tbl2.Cell(1,$c+1).Range.Bold = 1
    $tbl2.Cell(1,$c+1).Shading.BackgroundPatternColor = 0xC00000
    $tbl2.Cell(1,$c+1).Range.Font.ColorIndex = 2
}
for ($ri=0; $ri -lt $docentesNoNotif.Count; $ri++) {
    $d = $docentesNoNotif[$ri]
    $tbl2.Cell($ri+2,1).Range.Text = [string]$d.N
    $tbl2.Cell($ri+2,2).Range.Text = $d.Cedula
    $tbl2.Cell($ri+2,3).Range.Text = $d.Nombre
    $tbl2.Cell($ri+2,4).Range.Text = $d.TipoProducto
}
$tbl2.Columns.AutoFit() | Out-Null

$wsel.EndOf(6) | Out-Null
$wsel.TypeParagraph()
$wsel.TypeParagraph()

# Seccion 4: Todos los productos
$wsel.Style = "Titulo 2"
$wsel.TypeText("4. Detalle de Todos los Productos Presentados")
$wsel.TypeParagraph()

$tbl3 = $wdoc.Tables.Add($wsel.Range, ($allRows.Count + 1), 5)
$tbl3.Style = "Tabla con cuadricula"
$hdP = @("#","Cedula","Nombre","Tipo","Puntaje")
for ($c=0; $c -lt 5; $c++) {
    $tbl3.Cell(1,$c+1).Range.Text = $hdP[$c]
    $tbl3.Cell(1,$c+1).Range.Bold = 1
    $tbl3.Cell(1,$c+1).Shading.BackgroundPatternColor = 0x1F497D
    $tbl3.Cell(1,$c+1).Range.Font.ColorIndex = 2
}
$n = 1
foreach ($p in $sortedAll) {
    $tbl3.Cell($n+1,1).Range.Text = [string]$n
    $tbl3.Cell($n+1,2).Range.Text = $p.Cedula
    $tbl3.Cell($n+1,3).Range.Text = $p.Nombre
    $tbl3.Cell($n+1,4).Range.Text = $p.Tipo
    $tbl3.Cell($n+1,5).Range.Text = [string]$p.Puntaje
    $n++
}
$tbl3.Columns.AutoFit() | Out-Null

$wdoc.SaveAs2($wordOut)
$wdoc.Close($false)
$wd.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wd) | Out-Null

Write-Host "Word generado: $wordOut"
Write-Host "Listo!"
