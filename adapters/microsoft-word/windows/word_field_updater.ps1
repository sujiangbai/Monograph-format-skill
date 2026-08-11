param()

$ErrorActionPreference = "Stop"
$word = $null
$document = $null
$verification = $null

function Get-FieldName([int]$type) {
    switch ($type) {
        3 { "REF" }
        13 { "TOC" }
        33 { "PAGE" }
        37 { "PAGEREF" }
        default { $null }
    }
}

function Update-ApprovedFields($doc, [System.Collections.Generic.HashSet[string]]$allowed) {
    $counts = @{}
    foreach ($storyType in 1..17) {
        try { $range = $doc.StoryRanges.Item($storyType) } catch { $range = $null }
        while ($null -ne $range) {
            foreach ($field in @($range.Fields)) {
                $name = Get-FieldName ([int]$field.Type)
                if ($null -ne $name -and $allowed.Contains($name)) {
                    [void]$field.Update()
                    if (-not $counts.ContainsKey($name)) { $counts[$name] = 0 }
                    $counts[$name]++
                }
            }
            try { $range = $range.NextStoryRange } catch { $range = $null }
        }
    }
    foreach ($toc in @($doc.TablesOfContents)) {
        if ($allowed.Contains("TOC")) { [void]$toc.Update() }
    }
    return $counts
}

function Count-ApprovedFields($doc, [System.Collections.Generic.HashSet[string]]$allowed) {
    $count = 0
    foreach ($storyType in 1..17) {
        try { $range = $doc.StoryRanges.Item($storyType) } catch { $range = $null }
        while ($null -ne $range) {
            foreach ($field in @($range.Fields)) {
                $name = Get-FieldName ([int]$field.Type)
                if ($null -ne $name -and $allowed.Contains($name)) { $count++ }
            }
            try { $range = $range.NextStoryRange } catch { $range = $null }
        }
    }
    return $count
}

function Get-LastContentPage($section) {
    $paragraphs = @($section.Range.Paragraphs)
    for ($index = $paragraphs.Count - 1; $index -ge 0; $index--) {
        $paragraph = $paragraphs[$index]
        $visible = ([string]$paragraph.Range.Text)
        $visible = $visible.Replace(([char]13).ToString(), [string]::Empty)
        $visible = $visible.Replace(([char]7).ToString(), [string]::Empty)
        $visible = $visible.Replace(([char]12).ToString(), [string]::Empty).Trim()
        if ($visible -ne "" -or $paragraph.Range.InlineShapes.Count -gt 0) {
            return [int]$paragraph.Range.Information(3)
        }
    }
    $start = $section.Range.Duplicate
    $start.Collapse(1)
    return [int]$start.Information(3)
}

function Set-DisplayedPageField(
    $footer,
    [int]$offset,
    [int]$alignment
) {
    $footer.LinkToPrevious = $false
    $paragraph = $footer.Range.Paragraphs.First
    $paragraph.Alignment = $alignment
    $target = $paragraph.Range.Duplicate
    if ($target.End -gt $target.Start) { $target.End-- }
    $target.Text = ""
    if ($offset -eq 0) {
        $field = $footer.Range.Document.Fields.Add($target, 33, "", $false)
        [void]$field.Update()
        return
    }
    $outer = $footer.Range.Document.Fields.Add(
        $target,
        -1,
        "=  - $offset",
        $false
    )
    $code = $outer.Code.Duplicate
    $placeholder = ([string]$code.Text).IndexOf("-")
    if ($placeholder -lt 0) { throw "Unable to create a calculated PAGE field." }
    $inner = $footer.Range.Duplicate
    $inner.Start = $code.Start + $placeholder
    $inner.End = $inner.Start
    $page = $footer.Range.Document.Fields.Add($inner, 33, "", $false)
    [void]$page.Update()
    [void]$outer.Update()
}

function Unlock-ApprovedFields($doc, [System.Collections.Generic.HashSet[string]]$allowed) {
    foreach ($storyType in 1..17) {
        try { $range = $doc.StoryRanges.Item($storyType) } catch { $range = $null }
        while ($null -ne $range) {
            foreach ($field in @($range.Fields)) {
                $name = Get-FieldName ([int]$field.Type)
                if ($null -ne $name -and $allowed.Contains($name)) {
                    $field.Locked = $false
                }
            }
            try { $range = $range.NextStoryRange } catch { $range = $null }
        }
    }
}

function Normalize-DisplayedPageReferences($doc, $displayOffsets) {
    $corrected = 0
    foreach ($field in @($doc.StoryRanges.Item(1).Fields)) {
        if ([int]$field.Type -ne 37) { continue }
        $instruction = ([string]$field.Code.Text).Trim()
        if ($instruction -notmatch '^PAGEREF\s+([^\s\\]+)') { continue }
        $bookmarkName = $Matches[1]
        if (-not $doc.Bookmarks.Exists($bookmarkName)) { continue }
        $sectionNumber = [int]$doc.Bookmarks.Item($bookmarkName).Range.Information(2)
        $offset = 0
        if ($displayOffsets.ContainsKey([string]$sectionNumber)) {
            $offset = [int]$displayOffsets[[string]$sectionNumber]
        }
        if ($offset -le 0) { continue }
        [void]$field.Update()
        $visible = ([string]$field.Result.Text).Trim()
        $number = 0
        if (-not [int]::TryParse($visible, [ref]$number)) { continue }
        $field.Result.Text = [string]($number - $offset)
        $field.Locked = $true
        $corrected++
    }
    $tocLocked = 0
    if ($corrected -gt 0) {
        foreach ($toc in @($doc.TablesOfContents)) {
            foreach ($field in @($toc.Range.Fields)) {
                if ([int]$field.Type -eq 13) {
                    $field.Locked = $true
                    $tocLocked++
                }
            }
        }
    }
    return [ordered]@{ corrected = $corrected; toc_locked = $tocLocked }
}

function Remove-PageBoundaryBlockSpacers($doc, [string]$styleName) {
    $removed = 0
    for ($pass = 0; $pass -lt 10; $pass++) {
        [void]$doc.Repaginate()
        $targets = @()
        foreach ($paragraph in @($doc.Paragraphs)) {
            $name = ""
            try { $name = [string]$paragraph.Style.NameLocal } catch {}
            $visible = ([string]$paragraph.Range.Text).Replace("`r", "").Replace("`a", "").Trim()
            if ($name -eq $styleName -and $visible -eq "") { $targets += $paragraph }
        }
        $remove = @()
        foreach ($paragraph in $targets) {
            $range = $paragraph.Range.Duplicate
            if ($range.Start -le 0) { continue }
            $before = $doc.Range($range.Start - 1, $range.Start)
            $spacerPage = [int]$range.Information(3)
            $previousPage = [int]$before.Information(3)
            if ($spacerPage -ne $previousPage) { $remove += $paragraph }
        }
        if ($remove.Count -eq 0) { break }
        for ($index = $remove.Count - 1; $index -ge 0; $index--) {
            [void]$remove[$index].Range.Delete()
            $removed++
        }
    }
    return $removed
}

function Normalize-FrontMatterPagination($doc) {
    if ($doc.Sections.Count -lt 3) {
        throw "Approved title/TOC/body pagination requires at least three sections."
    }
    $doc.Sections.Item(1).PageSetup.DifferentFirstPageHeaderFooter = -1
    $changed = 1
    $displayOffsets = @{}
    foreach ($index in 2..3) {
        [void]$doc.Repaginate()
        $lastContentPage = Get-LastContentPage $doc.Sections.Item($index - 1)
        $desiredPhysicalPage = $lastContentPage + 1
        $startsEven = ($desiredPhysicalPage % 2 -eq 0)
        $doc.Sections.Item($index).PageSetup.SectionStart = if ($startsEven) { 3 } else { 4 }
        $offset = if ($startsEven) { 1 } else { 0 }
        $displayOffsets[[string]$index] = $offset
        try {
            [void]($doc.Sections.Item($index).Range.Paragraphs.First.Format.PageBreakBefore = 0)
        } catch {}
        Set-DisplayedPageField $doc.Sections.Item($index).Footers.Item(1) $offset 2
        Set-DisplayedPageField $doc.Sections.Item($index).Footers.Item(3) $offset 0
        $changed++
    }
    $bodyOffset = [int]$displayOffsets["3"]
    if ($doc.Sections.Count -ge 4) {
        foreach ($index in 4..$doc.Sections.Count) {
            $displayOffsets[[string]$index] = $bodyOffset
        }
    }
    [void]$doc.Repaginate()
    return [ordered]@{ changed = $changed; display_offsets = $displayOffsets }
}

try {
    $requestText = [Console]::In.ReadToEnd()
    $request = $requestText | ConvertFrom-Json
    if ($request.protocol_version -ne "1.0") {
        throw "Unsupported external field protocol version."
    }
    if ([string]$request.target_software -notmatch "Microsoft (Word|365)") {
        throw "This adapter only supports Microsoft Word target software."
    }
    $inputPath = [IO.Path]::GetFullPath([string]$request.input_path)
    $outputPath = [IO.Path]::GetFullPath([string]$request.output_path)
    if ($inputPath -eq $outputPath) { throw "The Word adapter cannot overwrite its input." }
    if (-not (Test-Path -LiteralPath $inputPath -PathType Leaf)) {
        throw "Input DOCX does not exist."
    }

    $allowed = [System.Collections.Generic.HashSet[string]]::new(
        [string[]]$request.allowed_field_types,
        [StringComparer]::OrdinalIgnoreCase
    )
    Copy-Item -LiteralPath $inputPath -Destination $outputPath -Force

    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3
    try { $word.Options.UpdateLinksAtOpen = $false } catch {}
    try { $word.Options.SaveNormalPrompt = $false } catch {}

    $document = $word.Documents.Open($outputPath)
    [void]$document.Repaginate()
    Unlock-ApprovedFields $document $allowed
    $spacersRemoved = 0
    $paginationSectionsNormalized = 0
    $displayOffsets = @{}
    $structureMapPath = [IO.Path]::GetFullPath([string]$request.structure_map_path)
    if (Test-Path -LiteralPath $structureMapPath -PathType Leaf) {
        $structureMap = Get-Content -LiteralPath $structureMapPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if (
            $null -ne $structureMap.front_matter -and
            $structureMap.front_matter.approved -eq $true -and
            $null -ne $structureMap.pagination_sections -and
            $structureMap.pagination_sections.approved -eq $true
        ) {
            $paginationResult = Normalize-FrontMatterPagination $document
            $paginationSectionsNormalized = [int]$paginationResult.changed
            $displayOffsets = $paginationResult.display_offsets
        }
        if (
            $null -ne $structureMap.block_spacing -and
            $structureMap.block_spacing.approved -eq $true -and
            $structureMap.block_spacing.same_page_only -eq $true
        ) {
            $spacersRemoved = Remove-PageBoundaryBlockSpacers $document "Monograph Figure Table Spacer"
        }
    }
    $updated = Update-ApprovedFields $document $allowed
    $displayReferenceResult = Normalize-DisplayedPageReferences $document $displayOffsets
    [void]$document.Repaginate()

    $pdfExported = $false
    if ($null -ne $request.pdf_output_path -and [string]$request.pdf_output_path -ne "") {
        $pdfPath = [IO.Path]::GetFullPath([string]$request.pdf_output_path)
        $document.ExportAsFixedFormat($pdfPath, 17)
        $pdfExported = Test-Path -LiteralPath $pdfPath -PathType Leaf
    }
    $document.SaveAs2($outputPath, 12)
    $document.Close($false)
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) | Out-Null
    $document = $null

    $verification = $word.Documents.Open($outputPath)
    [void]$verification.Repaginate()
    $verifiedFields = Count-ApprovedFields $verification $allowed
    $updatedTotal = 0
    foreach ($value in $updated.Values) { $updatedTotal += [int]$value }
    $pageCount = $verification.ComputeStatistics(2)
    $tocCount = $verification.TablesOfContents.Count
    $verification.Close($false)
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($verification) | Out-Null
    $verification = $null

    $wordVersion = [string]$word.Version
    $word.Quit()
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null
    $word = $null

    $result = [ordered]@{
        protocol_version = "1.0"
        status = "success"
        backend = "microsoft_word_com"
        software = "Microsoft Word"
        software_version = $wordVersion
        updated_fields = $updated
        updated_field_types = @($updated.Keys | Sort-Object)
        verified_field_count = $verifiedFields
        toc_count = $tocCount
        page_count = $pageCount
        repaginated = $true
        saved = $true
        field_cache_verified = ($verifiedFields -ge $updatedTotal)
        pdf_exported = $pdfExported
        page_boundary_spacers_removed = $spacersRemoved
        pagination_sections_normalized = $paginationSectionsNormalized
        page_number_display_offsets = $displayOffsets
        displayed_pageref_corrected = [int]$displayReferenceResult.corrected
        toc_locked_for_display_offsets = ([int]$displayReferenceResult.toc_locked -gt 0)
    }
    [Console]::Out.WriteLine(($result | ConvertTo-Json -Depth 5 -Compress))
    exit 0
}
catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
}
finally {
    if ($null -ne $verification) {
        try { $verification.Close($false) } catch {}
        try { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($verification) | Out-Null } catch {}
    }
    if ($null -ne $document) {
        try { $document.Close($false) } catch {}
        try { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) | Out-Null } catch {}
    }
    if ($null -ne $word) {
        try { $word.Quit() } catch {}
        try { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
