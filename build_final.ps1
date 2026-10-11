param([string]$Job='job.txt')
$ErrorActionPreference = 'Stop'
$jobLines = @(Get-Content $Job -Encoding UTF8)
$mdName = $jobLines[0].Trim()
$outName = $jobLines[1].Trim()
$newTitle = $jobLines[2].Trim()
$src = Join-Path (Get-Location).Path $jobLines[3].Trim()
$tocLabelDef = $jobLines[4].Trim()
$song = 'SimSun'
$yahei = 'Microsoft YaHei'
$lines = [IO.File]::ReadAllLines((Join-Path (Get-Location).Path $mdName), [Text.Encoding]::UTF8)

$outPath = Join-Path (Get-Location).Path $outName
Copy-Item -LiteralPath $src -Destination $outPath -Force
Write-Output 'STAGE=INIT'

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$doc = $word.Documents.Open($outPath, $false, $false)
$sel = $word.Selection
Write-Output 'STAGE=OPENED'

$doc.Content.Delete() | Out-Null
$sel.HomeKey(6) | Out-Null
Write-Output 'STAGE=CLEARED'

$st = $doc.Styles.Item(-1)
$st.Font.Name = 'Calibri'
$st.Font.NameFarEast = $song
$st.Font.Size = 10.5
$st.ParagraphFormat.Alignment = 3
$st.ParagraphFormat.SpaceAfter = 6
$st.ParagraphFormat.LineSpacingRule = 5
$st.ParagraphFormat.LineSpacing = 15
$st.ParagraphFormat.PageBreakBefore = $false

$h1 = $doc.Styles.Item(-2)
$h1.Font.Name = 'Calibri'
$h1.Font.NameFarEast = $yahei
$h1.Font.Size = 16
$h1.Font.Bold = 1
$h1.Font.Color = 6567967
$h1.ParagraphFormat.Alignment = 0
$h1.ParagraphFormat.PageBreakBefore = $true

$h2 = $doc.Styles.Item(-3)
$h2.Font.Name = 'Calibri'
$h2.Font.NameFarEast = $yahei
$h2.Font.Size = 13
$h2.Font.Bold = 1
$h2.ParagraphFormat.Alignment = 0
$h2.ParagraphFormat.PageBreakBefore = $false

$h3 = $doc.Styles.Item(-4)
$h3.Font.Name = 'Calibri'
$h3.Font.NameFarEast = $yahei
$h3.Font.Size = 12
$h3.Font.Bold = 1
$h3.Font.Color = 5855577
$h3.ParagraphFormat.Alignment = 0
$h3.ParagraphFormat.PageBreakBefore = $false
Write-Output 'STAGE=STYLES_DONE'

function Reset-Body {
  $sel.Style = $doc.Styles.Item(-1)
  $pf = $sel.ParagraphFormat
  $pf.Alignment = 3
  $pf.PageBreakBefore = 0
  $pf.LeftIndent = 0
  $pf.RightIndent = 0
  $pf.FirstLineIndent = 0
  $sel.Font.Bold = 0
  $sel.Font.Italic = 0
  $sel.Font.Color = -16777216
  $sel.Font.Size = 10.5
  $sel.Font.Name = 'Calibri'
  $sel.Font.NameFarEast = $song
  $sel.Shading.BackgroundPatternColor = -16777216
}

function Write-Inline([string]$s) {
  $parts = [regex]::Split($s, '(\*\*[^*]+\*\*|`[^`]+`)')
  foreach ($p in $parts) {
    if ($p.Length -eq 0) { continue }
    if ($p.StartsWith('**')) {
      $sel.Font.Bold = 1
      $sel.TypeText($p.Substring(2, $p.Length - 4))
      $sel.Font.Bold = 0
    } elseif ($p.StartsWith('`')) {
      $sel.Font.Name = 'Consolas'
      $sel.Font.NameFarEast = $song
      $sel.TypeText($p.Substring(1, $p.Length - 2))
      $sel.Font.Name = 'Calibri'
      $sel.Font.NameFarEast = $song
    } else {
      $sel.TypeText($p)
    }
  }
}

$tableBuf = New-Object 'System.Collections.Generic.List[string]'

function Emit-Table {
  $rows = @($tableBuf.ToArray() | Where-Object { $_ -notmatch '^\|[-\s:|]+\|$' })
  if ($rows.Count -eq 0) { return }
  $ncols = $rows[0].Trim().Trim('|').Split('|').Count
  $nrows = $rows.Count
  $tbl = $doc.Tables.Add($sel.Range, $nrows, $ncols)
  $tbl.Borders.Enable = $true
  for ($r = 0; $r -lt $nrows; $r++) {
    $cells = $rows[$r].Trim().Trim('|').Split('|')
    for ($c = 0; $c -lt $ncols; $c++) {
      $t = ''
      if ($c -lt $cells.Count) { $t = $cells[$c] }
      $t = (($t -replace '\*\*','') -replace '`','').Trim()
      $cell = $tbl.Cell($r+1, $c+1)
      $cellR = $cell.Range
      $cellR.Text = $t
    }
  }
  $tblR = $tbl.Range
  $tblF = $tblR.Font
  $tblF.Size = 9
  $tblP = $tblR.ParagraphFormat
  $tblP.SpaceAfter = 0
  $tblP.LineSpacingRule = 0
  $hdr = $tbl.Rows.Item(1)
  $hdrR = $hdr.Range
  $hdrF = $hdrR.Font
  $hdrF.Bold = 1
  $hc = $hdr.Cells.Count
  for ($k = 1; $k -le $hc; $k++) {
    $hCell = $hdr.Cells.Item($k)
    $hCell.Shading.BackgroundPatternColor = 16181982
  }
  $tbl.AutoFitBehavior(2)
  $sel.SetRange($tbl.Range.End, $tbl.Range.End)
  $sel.TypeParagraph()
}

$infoRows = @()
$titleText = ''
$subtitleText = ''
$inCode = $false

foreach ($raw in $lines) {
  $line = $raw.TrimEnd()

  if ($inCode) {
    if ($line.Trim() -eq '```') { $inCode = $false; continue }
    $sel.Font.Name = 'Consolas'
    $sel.Font.NameFarEast = $song
    $sel.Font.Size = 9
    $sel.Font.Color = -16777216
    $sel.Shading.BackgroundPatternColor = 15921906
    $sel.TypeText($line)
    $sel.TypeParagraph()
    continue
  }
  if ($line.TrimStart().StartsWith('```')) { $inCode = $true; Reset-Body; continue }

  if ($line.StartsWith('|')) { $tableBuf.Add($line); continue }
  if ($tableBuf.Count -gt 0) { Emit-Table; $tableBuf.Clear() }

  $t = $line.Trim()
  if ($t.Length -eq 0) { continue }
  if ($t -eq '---') { continue }

  if ($t.StartsWith('@@TITLE@@')) { $titleText = $t.Substring(9).Trim(); continue }
  if ($t.StartsWith('@@SUBTITLE@@')) { $subtitleText = $t.Substring(12).Trim(); continue }
  if ($t.StartsWith('@@INFO@@')) { $infoRows += $t.Substring(8).Trim(); continue }

  if ($t.StartsWith('@@TOC@@')) {
    $tocLabel = $t.Substring(7).Trim()
    if ($tocLabel.Length -eq 0) { $tocLabel = $tocLabelDef }
    Reset-Body
    $sel.Style = $doc.Styles.Item(-63)
    $sel.Font.Size = 20; $sel.Font.Bold = 1; $sel.Font.Color = 6567967
    $sel.ParagraphFormat.Alignment = 1
    $sel.TypeText($titleText)
    $sel.TypeParagraph()

    Reset-Body
    $sel.Style = $doc.Styles.Item(-75)
    $sel.Font.Size = 13; $sel.Font.Color = 5855577
    $sel.ParagraphFormat.Alignment = 1
    $sel.TypeText($subtitleText)
    $sel.TypeParagraph()

    1..4 | ForEach-Object { $sel.TypeParagraph() }

    if ($infoRows.Count -gt 0) {
      $it = $doc.Tables.Add($sel.Range, $infoRows.Count, 2)
      $it.Borders.Enable = $true
      for ($r = 0; $r -lt $infoRows.Count; $r++) {
        $body = $infoRows[$r]
        $ix = $body.IndexOf('|')
        $lb = $body.Substring(0, $ix).Trim()
        $vl = $body.Substring($ix + 1).Trim()
        $cl = $it.Cell($r+1, 1)
        $clr = $cl.Range
        $clr.Text = $lb
        $clf = $cl.Range.Font
        $clf.Bold = 1
        $cr = $it.Cell($r+1, 2)
        $crr = $cr.Range
        $crr.Text = $vl
      }
      $c1 = $it.Columns.Item(1)
      $c1.Width = 120
      $c2 = $it.Columns.Item(2)
      $c2.Width = 280
      $it.Rows.Alignment = 1
      $sel.SetRange($it.Range.End, $it.Range.End)
      $sel.TypeParagraph()
    }
    $sel.InsertBreak(7)

    Reset-Body
    $sel.Font.Size = 14; $sel.Font.Bold = 1; $sel.Font.Color = 6567967
    $sel.ParagraphFormat.Alignment = 1
    $sel.TypeText($tocLabel)
    $sel.TypeParagraph()

    $toc = $doc.TablesOfContents.Add($sel.Range, $true, 1, 3)
    $sel.SetRange($toc.Range.End, $toc.Range.End)
    continue
  }

  if ($t.StartsWith('# ')) {
    $sel.Style = $doc.Styles.Item(-2)
    $sel.ParagraphFormat.Alignment = 0
    $sel.ParagraphFormat.LeftIndent = 0
    $sel.ParagraphFormat.FirstLineIndent = 0
    $sel.TypeText($t.Substring(2).Trim())
    $sel.TypeParagraph()
    continue
  }
  if ($t.StartsWith('## ')) {
    $sel.Style = $doc.Styles.Item(-3)
    $sel.ParagraphFormat.Alignment = 0
    $sel.ParagraphFormat.LeftIndent = 0
    $sel.ParagraphFormat.FirstLineIndent = 0
    $sel.TypeText($t.Substring(3).Trim())
    $sel.TypeParagraph()
    continue
  }
  if ($t.StartsWith('### ')) {
    $sel.Style = $doc.Styles.Item(-4)
    $sel.ParagraphFormat.Alignment = 0
    $sel.ParagraphFormat.LeftIndent = 0
    $sel.ParagraphFormat.FirstLineIndent = 0
    $sel.TypeText($t.Substring(4).Trim())
    $sel.TypeParagraph()
    continue
  }
  if ($t.StartsWith('> ')) {
    Reset-Body
    $sel.ParagraphFormat.LeftIndent = 14
    $sel.Font.Italic = 1
    $sel.Font.Color = 5855577
    Write-Inline $t.Substring(2)
    $sel.TypeParagraph()
    continue
  }
  if ($t.StartsWith('- ')) {
    Reset-Body
    $sel.Style = $doc.Styles.Item(-49)
    Write-Inline $t.Substring(2)
    $sel.TypeParagraph()
    continue
  }
  if ($t -match '^(\d+)\. (.*)$') {
    Reset-Body
    $sel.ParagraphFormat.LeftIndent = 14
    $sel.ParagraphFormat.FirstLineIndent = -14
    $sel.TypeText($Matches[1] + '. ')
    Write-Inline $Matches[2]
    $sel.TypeParagraph()
    continue
  }

  Reset-Body
  Write-Inline $t
  $sel.TypeParagraph()
}

if ($tableBuf.Count -gt 0) { Emit-Table }
Write-Output 'STAGE=BODY_DONE'

$sec0 = $doc.Sections.Item(1)
if ($newTitle.Length -gt 0) {
  $h0 = $sec0.Headers.Item(1)
  $h0r = $h0.Range
  $h0r.Text = $newTitle
  $h0r2 = $h0.Range
  $h0f = $h0r2.Font
  $h0f.Size = 8
  $h0f.Color = 8421504
  try {
    $bp = $doc.BuiltInDocumentProperties
    $bp.Item('Title').Value = $newTitle
  } catch { Write-Output 'TITLE_PROP_WARN' }
}
Write-Output 'STAGE=HDR_DONE'

foreach ($t2 in $doc.TablesOfContents) { $t2.Update() }
Write-Output 'STAGE=TOC_DONE'
$lastTxt = $doc.Paragraphs.Last.Range.Text
Write-Output 'STAGE=LAST_OK'
$tailOk = $lastTxt.Contains('t.me/one00190')
$tocCount = $doc.TablesOfContents.Item(1).Range.Paragraphs.Count
Write-Output 'STAGE=TOCCOUNT_OK'

$doc.Save()
Write-Output 'STAGE=SAVED'
$doc.Close(0)
$word.Quit()

Write-Output ("TAIL_OK=" + $tailOk)
Write-Output ("TOC_N=" + $tocCount)
