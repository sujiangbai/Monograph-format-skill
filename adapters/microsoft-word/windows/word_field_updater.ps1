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
    $updated = Update-ApprovedFields $document $allowed
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
