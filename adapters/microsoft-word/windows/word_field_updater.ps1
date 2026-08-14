param()

$ErrorActionPreference = "Stop"
$word = $null
$document = $null
$verification = $null

function Get-FieldName([int]$type) {
    switch ($type) {
        3 { "REF" }
        12 { "SEQ" }
        13 { "TOC" }
        26 { "NUMPAGES" }
        33 { "PAGE" }
        34 { "=" }
        37 { "PAGEREF" }
        66 { "SECTIONPAGES" }
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
                    $updateSucceeded = [bool]$field.Update()
                    if (-not $updateSucceeded) {
                        throw "Microsoft Word failed to update an approved $name field."
                    }
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

function Measure-Layout($doc, [string]$spacerStyleName) {
    [void]$doc.Repaginate()
    $sections = @()
    for ($index = 1; $index -le $doc.Sections.Count; $index++) {
        $section = $doc.Sections.Item($index)
        $start = $section.Range.Duplicate
        $start.Collapse(1)
        $finish = $section.Range.Duplicate
        if ($finish.End -gt $finish.Start) { $finish.End-- }
        $finish.Collapse(0)
        $sections += [ordered]@{
            section_index = $index - 1
            first_physical_page = [int]$start.Information(3)
            last_physical_page = [int]$finish.Information(3)
            last_content_page = Get-LastContentPage $section
        }
    }
    $spacerOrdinals = @()
    if (-not [string]::IsNullOrWhiteSpace($spacerStyleName)) {
        $ordinal = 0
        foreach ($paragraph in @($doc.Paragraphs)) {
            $name = ""
            try { $name = [string]$paragraph.Style.NameLocal } catch {}
            if ($name -ne $spacerStyleName) { continue }
            $range = $paragraph.Range.Duplicate
            if ($range.Start -gt 0) {
                $before = $doc.Range($range.Start - 1, $range.Start)
                if ([int]$range.Information(3) -ne [int]$before.Information(3)) {
                    $spacerOrdinals += $ordinal
                }
            }
            $ordinal++
        }
    }
    return [ordered]@{
        sections = $sections
        page_boundary_spacer_ordinals = $spacerOrdinals
        page_count = [int]$doc.ComputeStatistics(2)
    }
}

function Export-RequestedPdf($doc, $request) {
    if ($null -eq $request.pdf_output_path -or [string]$request.pdf_output_path -eq "") {
        return $false
    }
    $pdfPath = [IO.Path]::GetFullPath([string]$request.pdf_output_path)
    $parent = [IO.Path]::GetDirectoryName($pdfPath)
    if (-not [string]::IsNullOrWhiteSpace($parent)) {
        [IO.Directory]::CreateDirectory($parent) | Out-Null
    }
    $doc.ExportAsFixedFormat($pdfPath, 17)
    return (Test-Path -LiteralPath $pdfPath -PathType Leaf)
}

try {
    $requestText = [Console]::In.ReadToEnd()
    $request = $requestText | ConvertFrom-Json
    if ($request.protocol_version -notin @("1.0", "1.1")) {
        throw "Unsupported external field protocol version."
    }
    if ([string]$request.target_software -notmatch "Microsoft (Word|365)") {
        throw "This adapter only supports Microsoft Word target software."
    }
    $operation = [string]$request.operation
    if ([string]::IsNullOrWhiteSpace($operation)) { $operation = "refresh_fields" }
    if ($operation -notin @("measure_layout", "refresh_fields", "verify_only")) {
        throw "Unsupported external field operation."
    }
    $inputPath = [IO.Path]::GetFullPath([string]$request.input_path)
    if (-not (Test-Path -LiteralPath $inputPath -PathType Leaf)) {
        throw "Input DOCX does not exist."
    }
    $allowed = [System.Collections.Generic.HashSet[string]]::new(
        [string[]]$request.allowed_field_types,
        [StringComparer]::OrdinalIgnoreCase
    )

    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3
    try { $word.Options.UpdateLinksAtOpen = $false } catch {}
    try { $word.Options.SaveNormalPrompt = $false } catch {}

    if ($operation -eq "measure_layout") {
        $document = $word.Documents.Open($inputPath, $false, $true)
        $openedReadOnly = [bool]$document.ReadOnly
        if (-not $openedReadOnly) { throw "Measurement input was not opened read-only." }
        $measurement = Measure-Layout $document ([string]$request.block_spacer_style_name)
        $document.Close($false)
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) | Out-Null
        $document = $null
        $wordVersion = [string]$word.Version
        $word.Quit()
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null
        $word = $null
        $result = [ordered]@{
            protocol_version = "1.1"
            operation = "measure_layout"
            status = "success"
            backend = "microsoft_word_com"
            software = "Microsoft Word"
            software_version = $wordVersion
            repaginated = $true
            saved = $false
            read_only_verified = $openedReadOnly
            structural_changes_applied = 0
            page_count = $measurement.page_count
            sections = $measurement.sections
            page_boundary_spacer_ordinals = $measurement.page_boundary_spacer_ordinals
        }
        [Console]::Out.WriteLine(($result | ConvertTo-Json -Depth 6 -Compress))
        exit 0
    }

    if ($operation -eq "verify_only") {
        $document = $word.Documents.Open($inputPath, $false, $true)
        $openedReadOnly = [bool]$document.ReadOnly
        if (-not $openedReadOnly) { throw "Verification input was not opened read-only." }
        [void]$document.Repaginate()
        $verifiedFields = Count-ApprovedFields $document $allowed
        $pageCount = $document.ComputeStatistics(2)
        $tocCount = $document.TablesOfContents.Count
        $pdfExported = Export-RequestedPdf $document $request
        $document.Close($false)
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) | Out-Null
        $document = $null
        $wordVersion = [string]$word.Version
        $word.Quit()
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null
        $word = $null
        $result = [ordered]@{
            protocol_version = "1.1"
            operation = "verify_only"
            status = "success"
            backend = "microsoft_word_com"
            software = "Microsoft Word"
            software_version = $wordVersion
            verified_field_count = $verifiedFields
            toc_count = $tocCount
            page_count = $pageCount
            repaginated = $true
            saved = $false
            read_only_verified = $openedReadOnly
            pdf_exported = $pdfExported
            structural_changes_applied = 0
        }
        [Console]::Out.WriteLine(($result | ConvertTo-Json -Depth 5 -Compress))
        exit 0
    }

    $outputPath = [IO.Path]::GetFullPath([string]$request.output_path)
    if ($inputPath -eq $outputPath) { throw "The Word adapter cannot overwrite its input." }
    Copy-Item -LiteralPath $inputPath -Destination $outputPath -Force
    $document = $word.Documents.Open($outputPath)
    [void]$document.Repaginate()
    Unlock-ApprovedFields $document $allowed
    $updated = Update-ApprovedFields $document $allowed
    [void]$document.Repaginate()
    $pdfExported = Export-RequestedPdf $document $request
    $document.SaveAs2($outputPath, 12)
    $document.Close($false)
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) | Out-Null
    $document = $null

    $verification = $word.Documents.Open($outputPath, $false, $true)
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
        protocol_version = "1.1"
        operation = "refresh_fields"
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
        structural_changes_applied = 0
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
