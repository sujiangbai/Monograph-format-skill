-- Explicit opt-in. Input must be this batch's freshly generated synthetic DOCX.
-- No field update, save, print, PDF, template write, or application quit.
-- Pure selection logic; offline tests execute this handler without Word.
on exactPathIndex(candidatePaths, inputPath, allowAbsent)
    set matchIndex to 0
    repeat with candidateIndex from 1 to count of candidatePaths
        set candidatePath to item candidateIndex of candidatePaths
        if class of candidatePath is not text then error "document_path_unavailable" number 7113
        considering case, diacriticals, hyphens, punctuation, white space
            if candidatePath is inputPath then
                if matchIndex is not 0 then error "multiple_exact_path_matches" number 7111
                set matchIndex to candidateIndex
            end if
        end considering
    end repeat
    if matchIndex is 0 and allowAbsent is false then error "no_exact_path_match_after_open" number 7112
    return matchIndex
end exactPathIndex

on exactDocument(inputPath, allowAbsent)
    tell application "Microsoft Word"
        set documentSnapshot to get documents
        set candidatePaths to {}
        repeat with candidateDocument in documentSnapshot
            -- Metadata only. Never read another document's content or use its name.
            set end of candidatePaths to get posix full name of candidateDocument
        end repeat
        set matchIndex to my exactPathIndex(candidatePaths, inputPath, allowAbsent)
        if matchIndex is 0 then return missing value
        set matchedDocument to item matchIndex of documentSnapshot
        -- Recheck the selected reference before granting ownership of the fixture.
        set currentPath to get posix full name of matchedDocument
        my exactPathIndex({currentPath}, inputPath, false)
        return matchedDocument
    end tell
end exactDocument

on requireReadOnly(observedReadOnly)
    if class of observedReadOnly is not boolean then error "not_confirmed_readonly" number 7114
    if observedReadOnly is not true then error "not_confirmed_readonly" number 7114
end requireReadOnly

on closeExactDocument(inputPath)
    tell application "Microsoft Word"
        with timeout of 30 seconds
            set closeTarget to my exactDocument(inputPath, true)
            if closeTarget is missing value then return "no_exact_document_to_close"
            close closeTarget saving no
            if (my exactDocument(inputPath, true)) is not missing value then error "exact_document_remains_after_close" number 7115
            return "exact_document_closed_without_save"
        end timeout
    end tell
end closeExactDocument

on recoverAfterFailure(inputPath, openAttempted, originalSecurity, originalLinks)
    set closeErrors to ""
    set closeOutcome to "open_not_attempted"
    if openAttempted then
        try
            -- Open can create a document and then raise. Still require one exact
            -- path before closing; ambiguity never grants ownership.
            set closeOutcome to my closeExactDocument(inputPath)
        on error closeMessage
            set closeErrors to closeMessage
            set closeOutcome to "close_not_verified"
        end try
    end if
    set restoreErrors to my restoreControls(originalSecurity, originalLinks)
    return "; close_outcome=" & closeOutcome & "; close_errors=" & closeErrors & "; restore_errors=" & restoreErrors
end recoverAfterFailure

on restoreControls(originalSecurity, originalLinks)
    set restoreErrors to ""
    tell application "Microsoft Word"
        with timeout of 30 seconds
            try
                set update links at open of settings to originalLinks
                if (get update links at open of settings) is not originalLinks then error "links_restore_readback"
            on error restoreMessage
                set restoreErrors to restoreErrors & " links: " & restoreMessage
            end try
            try
                set automation security to originalSecurity
                if (get automation security) is not originalSecurity then error "security_restore_readback"
            on error restoreMessage
                set restoreErrors to restoreErrors & " security: " & restoreMessage
            end try
        end timeout
    end tell
    return restoreErrors
end restoreControls

on run argv
    if (count of argv) is not 1 then error "Exactly one generated synthetic DOCX path required" number 7110
    set inputPath to item 1 of argv
    set openAttempted to false
    tell application "Microsoft Word"
        with timeout of 120 seconds
            if (count of documents) is not 0 then error "user_documents_present" number 7101
            if (get background printing status) is not 0 then error "printing_active" number 7102
            set originalSecurity to get automation security
            set originalLinks to get update links at open of settings
            if originalSecurity is missing value then error "security_original_unknown" number 7103
            if class of originalLinks is not boolean then error "links_original_unknown" number 7104
            log {"original_controls", originalSecurity, originalLinks}
            try
                set automation security to msoAutomationSecurityForceDisable
                set update links at open of settings to false
                if (get automation security) is not msoAutomationSecurityForceDisable then error "security_readback_failed" number 7105
                if (get update links at open of settings) is not false then error "links_readback_failed" number 7106
                if (count of documents) is not 0 then error "user_activity_conflict" number 7107
                if (get background printing status) is not 0 then error "printing_conflict" number 7108
                log {"safe_open_controls", (get automation security), (get update links at open of settings)}
                set openAttempted to true
                -- Do not assign the optional runtime reply: a no-result reply
                -- undefines the assigned variable in AppleScript (-2753).
                -- Errors from open still propagate to cleanup and are rethrown.
                -- Word.sdef: named file name is text (5015), distinct from
                -- the direct file argument. 009 tests this one parameter form.
                open file name inputPath read only true add to recent files false
                set ownedDocument to my exactDocument(inputPath, false)
                if (count of documents) is not 1 then error "user_activity_conflict" number 7107
                if (get background printing status) is not 0 then error "printing_conflict" number 7108
                -- Begin raw getter diagnostics: queries only, before the unchanged guard.
                set wasReadOnly to get read only of ownedDocument
                log {"diagnostic_read_only", wasReadOnly, class of wasReadOnly}
                set savedState to get saved of ownedDocument
                log {"diagnostic_saved", savedState, class of savedState}
                set printFields to get update fields at print of settings
                log {"diagnostic_print_fields", printFields, class of printFields}
                set printLinks to get update links at print of settings
                log {"diagnostic_print_links", printLinks, class of printLinks}
                set printCodes to get print field codes of settings
                log {"diagnostic_print_codes", printCodes, class of printCodes}
                my requireReadOnly(wasReadOnly)
                -- End raw getter diagnostics.
                set closeOutcome to my closeExactDocument(inputPath)
                if closeOutcome is not "exact_document_closed_without_save" then error "document_disappeared_before_close" number 7115
                if (count of documents) is not 0 then error "documents_remain_or_user_activity" number 7115
            on error failureMessage number failureNumber
                set recoveryDetails to my recoverAfterFailure(inputPath, openAttempted, originalSecurity, originalLinks)
                error (failureMessage & recoveryDetails) number failureNumber
            end try
            set restoreErrors to my restoreControls(originalSecurity, originalLinks)
            if restoreErrors is not "" then error ("restore_failed:" & restoreErrors) number 7109
            return {"readonly_probe_complete", wasReadOnly, printFields, printLinks, printCodes, savedState, closeOutcome, automation security, update links at open of settings, count of documents, background printing status}
        end timeout
    end tell
end run
