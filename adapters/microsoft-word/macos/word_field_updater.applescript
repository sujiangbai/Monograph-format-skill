-- Bounded protocol 1.1 Microsoft Word adapter for macOS.
-- The Python parent validates paths, approvals, OOXML instructions and artifacts.

property ownedDocumentPath : ""
property ownedDocumentIdentity : ""

on fileIdentity(filePath)
    try
        set identityValue to do shell script "/usr/bin/stat -f '%d:%i:%HT' " & quoted form of filePath
    on error
        error "document_identity_unavailable" number 7113
    end try
    if identityValue does not end with ":Regular File" then error "document_identity_not_regular" number 7113
    return identityValue
end fileIdentity

on ownedPathSpelling(candidatePath)
    if class of candidatePath is not text then error "document_path_unavailable" number 7113
    set expectedPath to my ownedDocumentPath
    if expectedPath is "" then error "document_identity_unbound" number 7113
    set permittedSpelling to false
    considering case, diacriticals, hyphens, punctuation, white space
        if candidatePath is expectedPath then set permittedSpelling to true
        -- Only the observed system /var alias of this already resolved target.
        -- No basename, arbitrary symlink, or global /private replacement.
        if expectedPath starts with "/private/var/" then
            if candidatePath is ("/var/" & text 14 thru -1 of expectedPath) then set permittedSpelling to true
        end if
    end considering
    return permittedSpelling
end ownedPathSpelling

on sameOwnedFile(candidatePath)
    set expectedPath to my ownedDocumentPath
    set expectedIdentity to my ownedDocumentIdentity
    if expectedIdentity is "" then error "document_identity_unbound" number 7113
    if my fileIdentity(expectedPath) is not expectedIdentity then error "document_identity_changed" number 7116
    if my ownedPathSpelling(candidatePath) is false then return false
    if my fileIdentity(candidatePath) is not expectedIdentity then error "document_identity_changed" number 7116
    return true
end sameOwnedFile

on bindCompletedOwnedSave(candidatePath)
    -- Called only after save of the already-held, prechecked document and its
    -- saved=true readback. Word may atomically replace that file while saving.
    if my ownedPathSpelling(candidatePath) is false then error "saved_document_path_changed" number 7116
    set expectedPath to my ownedDocumentPath
    set savedIdentity to my fileIdentity(expectedPath)
    if my fileIdentity(candidatePath) is not savedIdentity then error "saved_document_identity_changed" number 7116
    if my fileIdentity(expectedPath) is not savedIdentity then error "saved_document_identity_changed" number 7116
    set ownedDocumentIdentity to savedIdentity
end bindCompletedOwnedSave

on exactPathIndex(candidatePaths, inputPath, allowAbsent)
    if my sameOwnedFile(inputPath) is false then error "document_expected_path_changed" number 7116
    set matchIndex to 0
    repeat with candidateIndex from 1 to count of candidatePaths
        set candidatePath to item candidateIndex of candidatePaths
        if class of candidatePath is not text then error "document_path_unavailable" number 7113
        if my sameOwnedFile(candidatePath) then
            if matchIndex is not 0 then error "multiple_exact_path_matches" number 7111
            set matchIndex to candidateIndex
        end if
    end repeat
    if matchIndex is 0 and allowAbsent is false then error "no_exact_path_match_after_open" number 7112
    return matchIndex
end exactPathIndex

on exactDocument(inputPath, allowAbsent)
    tell application "Microsoft Word"
        set documentSnapshot to get documents
        set candidatePaths to {}
        repeat with candidateDocument in documentSnapshot
            copy (get posix full name of candidateDocument) to candidatePath
            copy candidatePath to end of candidatePaths
        end repeat
        set matchIndex to my exactPathIndex(candidatePaths, inputPath, allowAbsent)
        if matchIndex is 0 then return missing value
        set matchedDocument to item matchIndex of documentSnapshot
        copy (get posix full name of matchedDocument) to matchedPath
        my exactPathIndex({matchedPath}, inputPath, false)
        return matchedDocument
    end tell
end exactDocument

on closeExactDocument(inputPath)
    tell application "Microsoft Word"
        with timeout of 30 seconds
            set documentSnapshot to get documents
            set candidatePaths to {}
            repeat with candidateDocument in documentSnapshot
                copy (get posix full name of candidateDocument) to end of candidatePaths
            end repeat
            set matchIndex to my exactPathIndex(candidatePaths, inputPath, true)
            if matchIndex is 0 then return "no_exact_document_to_close"
            set closeTarget to item matchIndex of documentSnapshot
            my exactPathIndex({get posix full name of closeTarget}, inputPath, false)
            close closeTarget saving no
            set remainingDocumentSnapshot to get documents
            set remainingCandidatePaths to {}
            repeat with remainingDocument in remainingDocumentSnapshot
                copy (get posix full name of remainingDocument) to end of remainingCandidatePaths
            end repeat
            set remainingMatchIndex to my exactPathIndex(remainingCandidatePaths, inputPath, true)
            if remainingMatchIndex is not 0 then error "exact_document_remains_after_close" number 7115
            return "exact_document_closed_without_save"
        end timeout
    end tell
end closeExactDocument

on assertSafePrintControls(originalPrintFields, originalPrintLinks, originalPrintCodes, stagePrefix, stageState)
    tell application "Microsoft Word"
        set currentStage of stageState to stagePrefix & ".print_fields.get"
        set observedPrintFields to get update fields at print of settings
        if originalPrintFields is missing value then
            set currentStage of stageState to stagePrefix & ".print_fields.expected_missing"
            if observedPrintFields is missing value then
                set currentStage of stageState to stagePrefix & ".print_fields.expected_missing.observed_missing"
            else
                if class of observedPrintFields is not boolean then
                    set currentStage of stageState to stagePrefix & ".print_fields.expected_missing.unknown"
                    error "print_fields_state_unknown" number 7106
                end if
                if observedPrintFields is false then
                    set currentStage of stageState to stagePrefix & ".print_fields.expected_missing.safe_false"
                else
                    set currentStage of stageState to stagePrefix & ".print_fields.expected_missing.unsafe_true"
                    error "print_fields_unsafe_true" number 7106
                end if
            end if
        else
            set currentStage of stageState to stagePrefix & ".print_fields.expected_false"
            if class of observedPrintFields is not boolean then
                set currentStage of stageState to stagePrefix & ".print_fields.expected_false.unknown"
                error "print_fields_state_unknown" number 7106
            end if
            if observedPrintFields is false then
                set currentStage of stageState to stagePrefix & ".print_fields.expected_false.safe_false"
            else
                set currentStage of stageState to stagePrefix & ".print_fields.expected_false.unsafe_true"
                error "print_fields_unsafe_true" number 7106
            end if
        end if
        set currentStage of stageState to stagePrefix & ".print_links.get"
        set observedPrintLinks to get update links at print of settings
        if originalPrintLinks is missing value then
            set currentStage of stageState to stagePrefix & ".print_links.expected_missing"
            if observedPrintLinks is missing value then
                set currentStage of stageState to stagePrefix & ".print_links.expected_missing.observed_missing"
            else
                if class of observedPrintLinks is not boolean then
                    set currentStage of stageState to stagePrefix & ".print_links.expected_missing.unknown"
                    error "print_links_state_unknown" number 7106
                end if
                if observedPrintLinks is false then
                    set currentStage of stageState to stagePrefix & ".print_links.expected_missing.safe_false"
                else
                    set currentStage of stageState to stagePrefix & ".print_links.expected_missing.unsafe_true"
                    error "print_links_unsafe_true" number 7106
                end if
            end if
        else
            set currentStage of stageState to stagePrefix & ".print_links.expected_false"
            if class of observedPrintLinks is not boolean then
                set currentStage of stageState to stagePrefix & ".print_links.expected_false.unknown"
                error "print_links_state_unknown" number 7106
            end if
            if observedPrintLinks is false then
                set currentStage of stageState to stagePrefix & ".print_links.expected_false.safe_false"
            else
                set currentStage of stageState to stagePrefix & ".print_links.expected_false.unsafe_true"
                error "print_links_unsafe_true" number 7106
            end if
        end if
        set currentStage of stageState to stagePrefix & ".print_codes.get"
        set observedPrintCodes to get print field codes of settings
        if originalPrintCodes is missing value then
            set currentStage of stageState to stagePrefix & ".print_codes.expected_missing"
            if observedPrintCodes is missing value then
                set currentStage of stageState to stagePrefix & ".print_codes.expected_missing.observed_missing"
            else
                if class of observedPrintCodes is not boolean then
                    set currentStage of stageState to stagePrefix & ".print_codes.expected_missing.unknown"
                    error "print_codes_state_unknown" number 7106
                end if
                if observedPrintCodes is false then
                    set currentStage of stageState to stagePrefix & ".print_codes.expected_missing.safe_false"
                else
                    set currentStage of stageState to stagePrefix & ".print_codes.expected_missing.unsafe_true"
                    error "print_codes_unsafe_true" number 7106
                end if
            end if
        else
            set currentStage of stageState to stagePrefix & ".print_codes.expected_false"
            if class of observedPrintCodes is not boolean then
                set currentStage of stageState to stagePrefix & ".print_codes.expected_false.unknown"
                error "print_codes_state_unknown" number 7106
            end if
            if observedPrintCodes is false then
                set currentStage of stageState to stagePrefix & ".print_codes.expected_false.safe_false"
            else
                set currentStage of stageState to stagePrefix & ".print_codes.expected_false.unsafe_true"
                error "print_codes_unsafe_true" number 7106
            end if
        end if
    end tell
end assertSafePrintControls

on restoreControls(originalSecurity, originalOpenLinks, originalPrintFields, originalPrintLinks, originalPrintCodes)
    set restoreErrors to ""
    tell application "Microsoft Word"
        with timeout of 30 seconds
            try
                set update links at open of settings to originalOpenLinks
                if (get update links at open of settings) is not originalOpenLinks then error "open_links_restore_readback"
            on error restoreMessage
                set restoreErrors to restoreErrors & " open_links:" & restoreMessage
            end try
            try
                if originalPrintFields is missing value then
                    if (get update fields at print of settings) is not missing value then error "print_fields_missing_restore_readback"
                else
                    set update fields at print of settings to originalPrintFields
                    if (get update fields at print of settings) is not originalPrintFields then error "print_fields_restore_readback"
                end if
            on error restoreMessage
                set restoreErrors to restoreErrors & " print_fields:" & restoreMessage
            end try
            try
                if originalPrintLinks is missing value then
                    if (get update links at print of settings) is not missing value then error "print_links_missing_restore_readback"
                else
                    set update links at print of settings to originalPrintLinks
                    if (get update links at print of settings) is not originalPrintLinks then error "print_links_restore_readback"
                end if
            on error restoreMessage
                set restoreErrors to restoreErrors & " print_links:" & restoreMessage
            end try
            try
                if originalPrintCodes is missing value then
                    if (get print field codes of settings) is not missing value then error "print_codes_missing_restore_readback"
                else
                    set print field codes of settings to originalPrintCodes
                    if (get print field codes of settings) is not originalPrintCodes then error "print_codes_restore_readback"
                end if
            on error restoreMessage
                set restoreErrors to restoreErrors & " print_codes:" & restoreMessage
            end try
            try
                set automation security to originalSecurity
                if (get automation security) is not originalSecurity then error "security_restore_readback"
            on error restoreMessage
                set restoreErrors to restoreErrors & " security:" & restoreMessage
            end try
        end timeout
    end tell
    return restoreErrors
end restoreControls

on jsonValue(value)
    if class of value is list then
        set output to "["
        repeat with i from 1 to count of value
            if i > 1 then set output to output & ","
            set output to output & my jsonValue(contents of item i of value)
        end repeat
        return output & "]"
    else if class of value is integer then
        return value as text
    else if class of value is boolean then
        if value then return "true"
        return "false"
    else if class of value is text then
        set output to "\""
        repeat with c in characters of value
            set ch to c as text
            if ch is "\"" then
                set output to output & "\\\""
            else if ch is "\\" then
                set output to output & "\\\\"
            else if ch is return then
                set output to output & "\\r"
            else if ch is linefeed then
                set output to output & "\\n"
            else if ch is tab then
                set output to output & "\\t"
            else
                if (id of ch) < 32 then error "unexpected_control_character" number 7153
                set output to output & ch
            end if
        end repeat
        return output & "\""
    end if
    error "missing_or_unsupported_result_value" number 7153
end jsonValue

on sameTuple(leftTuple, rightTuple)
    considering case, diacriticals, hyphens, punctuation, white space
        return leftTuple is rightTuple
    end considering
end sameTuple

on pageAt(doc, characterOffset, adjusted, stageState)
    if class of characterOffset is not integer or characterOffset < 0 then error "invalid_page_offset" number 7155
    tell application "Microsoft Word"
        set parentStage to currentStage of stageState
        set currentStage of stageState to parentStage & ".range"
        set pointRange to create range doc start characterOffset end characterOffset
        set actualStart to get start of content of pointRange
        set actualEnd to get end of content of pointRange
        if actualStart is not characterOffset or actualEnd is not characterOffset then error "page_range_mismatch" number 7155
        set currentStage of stageState to parentStage & ".range_information"
        if adjusted then
            set pageRaw to get range information pointRange information type active end adjusted page number
        else
            set pageRaw to get range information pointRange information type active end page number
        end if
        if pageRaw is missing value then error "page_information_missing" number 7155
        set currentStage of stageState to parentStage & ".integer_conversion"
        set pageValue to pageRaw as integer
        if class of pageValue is not integer then error "page_information_not_integer" number 7155
        if pageValue < 1 then error "invalid_page_number" number 7155
        return pageValue
    end tell
end pageAt

on fieldKind(fieldObject)
    tell application "Microsoft Word"
        set kindValue to get field type of fieldObject
        if kindValue is field toc then return "TOC"
        if kindValue is field page then return "PAGE"
        if kindValue is field num pages then return "NUMPAGES"
        if kindValue is field section pages then return "SECTIONPAGES"
        if kindValue is field page ref then return "PAGEREF"
        if kindValue is field ref then return "REF"
        if kindValue is field formula then return "="
        return "UNAPPROVED"
    end tell
end fieldKind

on kindAllowed(kindName, allowedToken)
    return allowedToken contains ("|" & kindName & "|")
end kindAllowed

on topLevelApproved(fieldObjects, allowedToken, stageState)
    tell application "Microsoft Word"
        set descriptorStarts to {}
        set descriptorEnds to {}
        set parentStage to currentStage of stageState
        repeat with fieldObject in fieldObjects
            set currentStage of stageState to parentStage & ".field_code"
            set codeRange to get field code of fieldObject
            set currentStage of stageState to parentStage & ".field_result"
            set resultRange to get result range of fieldObject
            set currentStage of stageState to parentStage & ".field_code_start"
            set codeStartValue to get start of content of codeRange
            if class of codeStartValue is not integer then error "field_code_start_not_integer" number 7158
            set currentStage of stageState to parentStage & ".field_result_end"
            set resultEndValue to get end of content of resultRange
            if class of resultEndValue is not integer then error "field_result_end_not_integer" number 7158
            set currentStage of stageState to parentStage & ".descriptor_start_append"
            copy codeStartValue to end of descriptorStarts
            set currentStage of stageState to parentStage & ".descriptor_end_append"
            copy resultEndValue to end of descriptorEnds
        end repeat
        set currentStage of stageState to parentStage & ".selection_setup"
        set selected to {}
        repeat with i from 1 to count of fieldObjects
            set nested to false
            set currentStage of stageState to parentStage & ".selection_inner_start"
            copy item i of descriptorStarts to innerStartValue
            set currentStage of stageState to parentStage & ".selection_inner_end"
            copy item i of descriptorEnds to innerEndValue
            repeat with j from 1 to count of descriptorStarts
                if i is not j then
                    set currentStage of stageState to parentStage & ".selection_outer_start"
                    copy item j of descriptorStarts to outerStartValue
                    set currentStage of stageState to parentStage & ".selection_outer_end"
                    copy item j of descriptorEnds to outerEndValue
                    set currentStage of stageState to parentStage & ".selection_compare_start"
                    set startInside to innerStartValue > outerStartValue
                    if class of startInside is not boolean then error "nested_start_comparison_unknown" number 7158
                    set currentStage of stageState to parentStage & ".selection_compare_end"
                    set endInside to innerEndValue < outerEndValue
                    if class of endInside is not boolean then error "nested_end_comparison_unknown" number 7158
                    set currentStage of stageState to parentStage & ".selection_mark_nested"
                    if startInside and endInside then set nested to true
                end if
            end repeat
            set currentStage of stageState to parentStage & ".selection_candidate"
            set candidate to item i of fieldObjects
            set currentStage of stageState to parentStage & ".selection_kind"
            set kindName to my fieldKind(candidate)
            set currentStage of stageState to parentStage & ".selection_allowed_kind"
            set allowedKind to my kindAllowed(kindName, allowedToken)
            if class of allowedKind is not boolean then error "allowed_kind_unknown" number 7158
            set currentStage of stageState to parentStage & ".selection_decision"
            set shouldSelect to false
            if nested is false then
                if allowedKind then set shouldSelect to true
            end if
            if class of shouldSelect is not boolean then error "selection_decision_unknown" number 7158
            if shouldSelect then
                set currentStage of stageState to parentStage & ".selection_append"
                copy contents of candidate to end of selected
            end if
        end repeat
        set currentStage of stageState to parentStage & ".return"
        return selected
    end tell
end topLevelApproved

on parseStoryPlan(planToken)
    if planToken is "" then return {}
    set savedDelimiters to AppleScript's text item delimiters
    try
        set AppleScript's text item delimiters to ";"
        set rawEntries to text items of planToken
        set parsedPlan to {}
        set seenOwners to {}
        repeat with rawEntryReference in rawEntries
            set rawEntry to contents of rawEntryReference
            set AppleScript's text item delimiters to ","
            set components to text items of rawEntry
            if (count of components) is not 2 then error "invalid_story_plan" number 7160
            set sectionIndex to item 1 of components as integer
            set storyLabel to item 2 of components
            if sectionIndex < 1 then error "invalid_story_plan" number 7160
            if storyLabel is not in {"header_primary", "header_first", "header_even", "footer_primary", "footer_first", "footer_even"} then error "invalid_story_plan" number 7160
            set ownerToken to (sectionIndex as string) & "," & storyLabel
            if ownerToken is in seenOwners then error "duplicate_story_plan_owner" number 7160
            copy ownerToken to end of seenOwners
            copy {sectionIndex, storyLabel} to end of parsedPlan
            set AppleScript's text item delimiters to ";"
        end repeat
        set AppleScript's text item delimiters to savedDelimiters
        return parsedPlan
    on error failureMessage number failureNumber
        set AppleScript's text item delimiters to savedDelimiters
        error failureMessage number failureNumber
    end try
end parseStoryPlan

on storyFieldIdentities(fieldObjects, stagePrefix, stageState)
    set identities to {}
    repeat with fieldReference in fieldObjects
        set currentStage of stageState to stagePrefix & ".identity.field"
        set fieldObject to contents of fieldReference
        set currentStage of stageState to stagePrefix & ".identity.kind"
        set kindName to my fieldKind(fieldObject)
        tell application "Microsoft Word"
            set currentStage of stageState to stagePrefix & ".identity.code_range"
            set codeRange to get field code of fieldObject
            set currentStage of stageState to stagePrefix & ".identity.code_value"
            set codeText to get content of codeRange
        end tell
        if class of codeText is not text then error "story_field_identity_unavailable" number 7160
        copy {kindName, codeText} to end of identities
    end repeat
    return identities
end storyFieldIdentities

on approvedGroups(doc, allowedToken, storyPlan, stageState)
    set groups to {}
    set groupOwners to {{0, "body"}}
    set storyObservations to {}
    tell application "Microsoft Word"
        set currentStage of stageState to "approved_groups.body_fields"
        set bodyFields to get fields of doc
        set currentStage of stageState to "approved_groups.body_filter"
        set bodyApproved to my topLevelApproved(bodyFields, allowedToken, stageState)
        copy contents of bodyApproved to end of groups
        set currentStage of stageState to "approved_groups.sections"
        set sectionObjects to get sections of doc
        set sectionCount to count of sectionObjects
        repeat with planReference in storyPlan
            set currentStage of stageState to "approved_groups.plan_entry"
            set planEntry to contents of planReference
            set sectionIndex to item 1 of planEntry
            set storyLabel to item 2 of planEntry
            if class of sectionIndex is not integer then error "invalid_story_plan" number 7160
            if sectionIndex < 1 then error "invalid_story_plan" number 7160
            if sectionIndex > sectionCount then error "invalid_story_plan" number 7160
            set currentStage of stageState to "approved_groups.section_object"
            set sectionObject to item sectionIndex of sectionObjects
            set currentStage of stageState to "approved_groups." & storyLabel & ".story_object"
            if storyLabel is "header_primary" then
                set storyObject to get header sectionObject index header footer primary
            else if storyLabel is "header_first" then
                set storyObject to get header sectionObject index header footer first page
            else if storyLabel is "header_even" then
                set storyObject to get header sectionObject index header footer even pages
            else if storyLabel is "footer_primary" then
                set storyObject to get footer sectionObject index header footer primary
            else if storyLabel is "footer_first" then
                set storyObject to get footer sectionObject index header footer first page
            else if storyLabel is "footer_even" then
                set storyObject to get footer sectionObject index header footer even pages
            else
                error "invalid_story_plan" number 7160
            end if
            if storyObject is missing value then error "planned_story_unavailable" number 7160
            if sectionIndex is 1 then
                set currentStage of stageState to "approved_groups." & storyLabel & ".ownership_first_section"
                set ownershipToken to "first_section"
            else
                set currentStage of stageState to "approved_groups." & storyLabel & ".link_state"
                set linkedToPrevious to get link to previous of storyObject
                set currentStage of stageState to "approved_groups." & storyLabel & ".link_state_validation"
                if class of linkedToPrevious is not boolean then error "story_link_state_unknown" number 7158
                if linkedToPrevious is true then
                    error "planned_story_not_independent" number 7160
                else if linkedToPrevious is false then
                    set currentStage of stageState to "approved_groups." & storyLabel & ".ownership_independent"
                    set ownershipToken to "independent"
                else
                    error "story_link_state_unknown" number 7158
                end if
            end if
            set currentStage of stageState to "approved_groups." & storyLabel & ".story_text_object"
            set storyTextObject to get text object of storyObject
            set currentStage of stageState to "approved_groups." & storyLabel & ".story_fields"
            set storyFields to get fields of storyTextObject
            set storyFieldCount to count of storyFields
            if class of storyFieldCount is not integer then error "planned_story_field_count_invalid" number 7160
            if storyFieldCount < 1 then error "planned_story_field_count_invalid" number 7160
            set currentStage of stageState to "approved_groups." & storyLabel & ".story_identity"
            set identityRows to my storyFieldIdentities(storyFields, "approved_groups." & storyLabel, stageState)
            set currentStage of stageState to "approved_groups." & storyLabel & ".story_filter"
            set storyApproved to my topLevelApproved(storyFields, allowedToken, stageState)
            set currentStage of stageState to "approved_groups." & storyLabel & ".append"
            copy contents of storyApproved to end of groups
            copy {sectionIndex, storyLabel} to end of groupOwners
            copy {sectionIndex, storyLabel, ownershipToken, storyFieldCount, identityRows} to end of storyObservations
            set currentStage of stageState to "approved_groups." & storyLabel & ".branch_complete"
        end repeat
        set currentStage of stageState to "approved_groups.story_loop_complete"
        set currentStage of stageState to "approved_groups.section_loop_complete"
    end tell
    set currentStage of stageState to "approved_groups.return_prepare"
    set returnGroups to {}
    repeat with groupReference in groups
        set currentStage of stageState to "approved_groups.return_prepare.group"
        set groupValue to contents of groupReference
        set returnGroup to {}
        repeat with fieldReference in groupValue
            set currentStage of stageState to "approved_groups.return_prepare.field"
            copy contents of fieldReference to end of returnGroup
        end repeat
        copy returnGroup to end of returnGroups
    end repeat
    set currentStage of stageState to "approved_groups.return_execute"
    set returnObservations to {}
    repeat with observationReference in storyObservations
        copy contents of observationReference to end of returnObservations
    end repeat
    return {returnGroups, returnObservations, groupOwners}
end approvedGroups

on snapshotApprovedGroups(doc, groups, stageState, groupOwners)
    set tocTexts to {}
    set tocSpans to {}
    set fieldRows to {}
    set footerRows to {}
    set groupIndex to 0
    set currentStage of stageState to "snapshot.approved_groups.repeat_setup"
    repeat with fieldGroupReference in groups
        set currentStage of stageState to "snapshot.approved_groups.group_contents"
        set fieldGroup to contents of fieldGroupReference
        set groupIndex to groupIndex + 1
        set groupOwner to item groupIndex of groupOwners
        repeat with fieldReference in fieldGroup
            set currentStage of stageState to "snapshot.approved_groups.field_contents"
            set fieldObject to contents of fieldReference
            set currentStage of stageState to "snapshot.approved_groups.field_kind"
            set kindName to my fieldKind(fieldObject)
            set currentStage of stageState to "snapshot.field_rows.code"
            tell application "Microsoft Word"
                set codeRange to get field code of fieldObject
                set codeText to get content of codeRange
            end tell
            set currentStage of stageState to "snapshot.field_rows.result"
            tell application "Microsoft Word"
                set resultObject to get result range of fieldObject
                set resultText to get content of resultObject
            end tell
            if kindName is "TOC" then
                set currentStage of stageState to "snapshot.toc.result"
                if class of resultText is not text then error "toc_result_unavailable" number 7155
                tell application "Microsoft Word"
                    set tocStartOffset to get start of content of resultObject
                end tell
                set tocStartOffset to tocStartOffset as integer
                set currentStage of stageState to "snapshot.toc.page_span.start"
                set startPage to my pageAt(doc, tocStartOffset, false, stageState)
                set currentStage of stageState to "snapshot.toc.page_span.end"
                tell application "Microsoft Word"
                    set finishOffset to get end of content of resultObject
                end tell
                set finishOffset to finishOffset as integer
                if finishOffset > tocStartOffset then set finishOffset to finishOffset - 1
                set finishPage to my pageAt(doc, finishOffset, false, stageState)
                set end of tocTexts to resultText
                set end of tocSpans to finishPage - startPage + 1
            else
                set currentStage of stageState to "snapshot.field_rows.append"
                set end of fieldRows to {kindName, codeText, resultText}
                if (item 2 of groupOwner) is in {"footer_primary", "footer_first", "footer_even"} then
                    if kindName is in {"PAGE", "NUMPAGES"} then
                        copy {item 1 of groupOwner, item 2 of groupOwner, kindName, codeText, resultText} to end of footerRows
                    end if
                end if
            end if
        end repeat
    end repeat
    return {tocTexts, tocSpans, fieldRows, footerRows}
end snapshotApprovedGroups

on updateApprovedGroups(groups, tocOnly, stagePrefix, stageState)
    set currentStage of stageState to stagePrefix & ".repeat_setup"
    repeat with fieldGroupReference in groups
        set currentStage of stageState to stagePrefix & ".group_contents"
        set fieldGroup to contents of fieldGroupReference
        repeat with fieldReference in fieldGroup
            set currentStage of stageState to stagePrefix & ".field_contents"
            set fieldObject to contents of fieldReference
            set currentStage of stageState to stagePrefix & ".field_kind"
            set kindName to my fieldKind(fieldObject)
            set shouldUpdate to false
            if tocOnly is true then
                if kindName is "TOC" then set shouldUpdate to true
            else
                if kindName is not "TOC" then set shouldUpdate to true
            end if
            if shouldUpdate then
                set currentStage of stageState to stagePrefix & ".update"
                tell application "Microsoft Word"
                    set updateSucceeded to update field fieldObject
                end tell
                if updateSucceeded is not true then
                    if tocOnly is true then
                        error "approved_toc_update_failed" number 7156
                    else
                        error "approved_field_update_failed" number 7156
                    end if
                end if
            end if
        end repeat
    end repeat
end updateApprovedGroups

on countApprovedFields(groups, stagePrefix, stageState)
    set verifiedCount to 0
    set currentStage of stageState to stagePrefix & ".repeat_setup"
    repeat with fieldGroupReference in groups
        set currentStage of stageState to stagePrefix & ".group_contents"
        set fieldGroup to contents of fieldGroupReference
        repeat with fieldReference in fieldGroup
            set currentStage of stageState to stagePrefix & ".field_contents"
            set fieldObject to contents of fieldReference
            set currentStage of stageState to stagePrefix & ".count"
            set verifiedCount to verifiedCount + 1
        end repeat
    end repeat
    set currentStage of stageState to stagePrefix & ".return"
    return verifiedCount
end countApprovedFields

on allFieldGroupState(fieldObjects, stagePrefix, stageState)
    set groupState to {}
    set currentStage of stageState to stagePrefix & ".repeat_setup"
    repeat with fieldReference in fieldObjects
        set currentStage of stageState to stagePrefix & ".field_contents"
        set fieldObject to contents of fieldReference
        tell application "Microsoft Word"
            set currentStage of stageState to stagePrefix & ".instruction_range"
            set instructionRange to get field code of fieldObject
            set currentStage of stageState to stagePrefix & ".instruction_value"
            set instructionValue to get content of instructionRange
            set currentStage of stageState to stagePrefix & ".result_range"
            set fieldResultRange to get result range of fieldObject
            set currentStage of stageState to stagePrefix & ".result_value"
            set fieldResultValue to get content of fieldResultRange
        end tell
        if class of instructionValue is not text then error "field_instruction_state_unavailable" number 7159
        if class of fieldResultValue is not text then error "field_result_state_unavailable" number 7159
        set currentStage of stageState to stagePrefix & ".append"
        copy {instructionValue, fieldResultValue} to end of groupState
    end repeat
    set currentStage of stageState to stagePrefix & ".return"
    return groupState
end allFieldGroupState

on allFieldState(doc, storyPlan, stagePrefix, stageState)
    set allStates to {}
    tell application "Microsoft Word"
        set currentStage of stageState to stagePrefix & ".body.fields"
        set bodyFields to get fields of doc
    end tell
    set currentStage of stageState to stagePrefix & ".body.snapshot"
    set bodyState to my allFieldGroupState(bodyFields, stagePrefix & ".body", stageState)
    copy bodyState to end of allStates
    tell application "Microsoft Word"
        set currentStage of stageState to stagePrefix & ".sections"
        set sectionObjects to get sections of doc
        set sectionCount to count of sectionObjects
    end tell
    repeat with planReference in storyPlan
        set currentStage of stageState to stagePrefix & ".plan_entry"
        set planEntry to contents of planReference
        set sectionIndex to item 1 of planEntry
        set storyLabel to item 2 of planEntry
        if class of sectionIndex is not integer then error "invalid_story_plan" number 7160
        if sectionIndex < 1 then error "invalid_story_plan" number 7160
        if sectionIndex > sectionCount then error "invalid_story_plan" number 7160
        tell application "Microsoft Word"
            set currentStage of stageState to stagePrefix & ".section"
            set sectionObject to item sectionIndex of sectionObjects
            set currentStage of stageState to stagePrefix & "." & storyLabel & ".story_object"
            if storyLabel is "header_primary" then
                set storyObject to get header sectionObject index header footer primary
            else if storyLabel is "header_first" then
                set storyObject to get header sectionObject index header footer first page
            else if storyLabel is "header_even" then
                set storyObject to get header sectionObject index header footer even pages
            else if storyLabel is "footer_primary" then
                set storyObject to get footer sectionObject index header footer primary
            else if storyLabel is "footer_first" then
                set storyObject to get footer sectionObject index header footer first page
            else if storyLabel is "footer_even" then
                set storyObject to get footer sectionObject index header footer even pages
            else
                error "invalid_story_plan" number 7160
            end if
            if storyObject is missing value then error "planned_story_unavailable" number 7160
            if sectionIndex > 1 then
                set currentStage of stageState to stagePrefix & "." & storyLabel & ".link_state"
                set linkedToPrevious to get link to previous of storyObject
                if class of linkedToPrevious is not boolean then error "story_link_state_unknown" number 7158
                if linkedToPrevious is true then
                    error "planned_story_not_independent" number 7160
                else if linkedToPrevious is false then
                    set currentStage of stageState to stagePrefix & "." & storyLabel & ".ownership_independent"
                else
                    error "story_link_state_unknown" number 7158
                end if
            end if
            set currentStage of stageState to stagePrefix & "." & storyLabel & ".text_object"
            set storyTextObject to get text object of storyObject
            set currentStage of stageState to stagePrefix & "." & storyLabel & ".fields"
            set storyFields to get fields of storyTextObject
            set storyFieldCount to count of storyFields
            if class of storyFieldCount is not integer then error "planned_story_field_count_invalid" number 7160
            if storyFieldCount < 1 then error "planned_story_field_count_invalid" number 7160
        end tell
        set currentStage of stageState to stagePrefix & "." & storyLabel & ".snapshot"
        set storyState to my allFieldGroupState(storyFields, stagePrefix & "." & storyLabel, stageState)
        copy storyState to end of allStates
    end repeat
    set currentStage of stageState to stagePrefix & ".return"
    return allStates
end allFieldState

on lastContentPage(sectionObject, doc, stageState)
    tell application "Microsoft Word"
        set currentStage of stageState to "snapshot.section.last_content.paragraphs"
        set sectionTextObject to get text object of sectionObject
        set paragraphObjects to get paragraphs of sectionTextObject
        set paragraphCount to count of paragraphObjects
        repeat with paragraphIndex from paragraphCount to 1 by -1
            set currentStage of stageState to "snapshot.section.last_content.paragraph"
            set paragraphObject to item paragraphIndex of paragraphObjects
            set paragraphTextObject to get text object of paragraphObject
            set paragraphText to get content of paragraphTextObject
            set visibleText to paragraphText
            repeat with markerText in {return, character id 7, character id 12, space, tab}
                set AppleScript's text item delimiters to markerText
                set visibleText to text items of visibleText
                set AppleScript's text item delimiters to ""
                set visibleText to visibleText as text
            end repeat
            set AppleScript's text item delimiters to ""
            set inlineShapeCount to count of inline shapes of paragraphTextObject
            if visibleText is not "" or inlineShapeCount > 0 then
                set currentStage of stageState to "snapshot.section.last_content.page"
                set paragraphStart to get start of content of paragraphTextObject
                set paragraphStart to paragraphStart as integer
                return my pageAt(doc, paragraphStart, false, stageState)
            end if
        end repeat
        set currentStage of stageState to "snapshot.section.last_content.fallback"
        set sectionStart to get start of content of sectionTextObject
        set sectionStart to sectionStart as integer
        return my pageAt(doc, sectionStart, false, stageState)
    end tell
end lastContentPage

on captureSnapshot(doc, allowedToken, storyPlan, stageState)
    tell application "Microsoft Word"
        set sectionRows to {}
        set sectionCount to count of sections of doc
        repeat with sectionIndex from 1 to sectionCount
            set currentStage of stageState to "snapshot.section.text"
            set sectionObject to section sectionIndex of doc
            set sectionText to get text object of sectionObject
            set currentStage of stageState to "snapshot.section.start"
            set firstOffset to get start of content of sectionText
            set firstOffset to firstOffset as integer
            set currentStage of stageState to "snapshot.section.end"
            set lastOffset to get end of content of sectionText
            set lastOffset to lastOffset as integer
            if lastOffset > firstOffset then set lastOffset to lastOffset - 1
            set currentStage of stageState to "snapshot.section.page_number_options"
            set footerObject to get footer sectionObject index header footer primary
            set options to get page number options of footerObject
            set startNumber to get starting number of options
            set restartNumbering to get restart numbering at section of options
            if class of startNumber is not integer then error "invalid_section_start" number 7155
            if class of restartNumbering is not boolean then error "invalid_section_restart" number 7155
            set styleValue to get number style of options
            if styleValue is page number style arabic then
                set styleName to "decimal"
            else if styleValue is page number style lowercase roman then
                set styleName to "lowerRoman"
            else
                error "unsupported_page_number_format" number 7155
            end if
            set currentStage of stageState to "snapshot.section.start.physical"
            set firstPhysicalPage to my pageAt(doc, firstOffset, false, stageState)
            set currentStage of stageState to "snapshot.section.end.physical"
            set lastPhysicalPage to my pageAt(doc, lastOffset, false, stageState)
            set currentStage of stageState to "snapshot.section.last_content"
            set contentPage to my lastContentPage(sectionObject, doc, stageState)
            set currentStage of stageState to "snapshot.section.start.logical"
            set firstLogicalPage to my pageAt(doc, firstOffset, true, stageState)
            set end of sectionRows to {sectionIndex - 1, firstPhysicalPage, lastPhysicalPage, contentPage, firstLogicalPage, restartNumbering, startNumber, styleName}
        end repeat
        set currentStage of stageState to "snapshot.total_pages.text"
        set documentText to get text object of doc
        set currentStage of stageState to "snapshot.total_pages.range_information"
        set pageCountRaw to get range information documentText information type number of pages in document
        if pageCountRaw is missing value then error "document_page_count_missing" number 7155
        set currentStage of stageState to "snapshot.total_pages.integer_conversion"
        set pageCount to pageCountRaw as integer
        if class of pageCount is not integer then error "document_page_count_not_integer" number 7155
        if pageCount < 1 then error "invalid_document_page_count" number 7155
        set currentStage of stageState to "snapshot.spacer_rows"
        set spacerRows to {}
        -- Layout queries may materialize pagination-dependent field results.
        -- Sample fields after these queries, never mix an earlier field sample
        -- with the later page statistics. Full tuple comparisons remain strict.
        set currentStage of stageState to "snapshot.approved_groups"
        set approvedResult to my approvedGroups(doc, allowedToken, storyPlan, stageState)
        set groups to item 1 of approvedResult
        set currentStage of stageState to "snapshot.approved_groups.received"
        set currentStage of stageState to "snapshot.approved_groups.first_use"
        set fieldSnapshot to my snapshotApprovedGroups(doc, groups, stageState, item 3 of approvedResult)
        set tocTexts to item 1 of fieldSnapshot
        set tocSpans to item 2 of fieldSnapshot
        set fieldRows to item 3 of fieldSnapshot
        return {tocTexts, tocSpans, sectionRows, pageCount, fieldRows, spacerRows, {item 2 of approvedResult, item 4 of fieldSnapshot}}
    end tell
end captureSnapshot

on countKinds(groups, stagePrefix, stageState)
    set pairs to {}
    set currentStage of stageState to stagePrefix & ".repeat_setup"
    repeat with kindNameReference in {"TOC", "PAGE", "NUMPAGES", "SECTIONPAGES", "PAGEREF", "REF", "="}
        set currentStage of stageState to stagePrefix & ".kind_contents"
        set expectedKind to contents of kindNameReference
        set found to 0
        repeat with fieldGroupReference in groups
            set currentStage of stageState to stagePrefix & ".group_contents"
            set fieldGroup to contents of fieldGroupReference
            repeat with fieldReference in fieldGroup
                set currentStage of stageState to stagePrefix & ".field_contents"
                set fieldObject to contents of fieldReference
                set currentStage of stageState to stagePrefix & ".field_kind"
                set actualKind to my fieldKind(fieldObject)
                set currentStage of stageState to stagePrefix & ".count"
                if actualKind is expectedKind then set found to found + 1
            end repeat
        end repeat
        if found > 0 then set end of pairs to {expectedKind, found}
    end repeat
    return pairs
end countKinds

on calculateFields(doc, allowedToken, storyPlan, stageState)
    tell application "Microsoft Word"
        set snapshots to {}
        set converged to false
        repeat with roundIndex from 1 to 3
            set currentStage of stageState to "refresh.approved_groups.toc"
            set approvedResult to my approvedGroups(doc, allowedToken, storyPlan, stageState)
            set groups to item 1 of approvedResult
            set currentStage of stageState to "refresh.approved_groups.toc.received"
            set currentStage of stageState to "refresh.approved_groups.toc.first_use"
            my updateApprovedGroups(groups, true, "refresh.approved_groups.toc", stageState)
            set currentStage of stageState to "refresh.repaginate"
            repaginate doc
            set currentStage of stageState to "refresh.approved_groups.field_rows"
            set approvedResult to my approvedGroups(doc, allowedToken, storyPlan, stageState)
            set groups to item 1 of approvedResult
            set currentStage of stageState to "refresh.approved_groups.field_rows.received"
            set currentStage of stageState to "refresh.approved_groups.field_rows.first_use"
            my updateApprovedGroups(groups, false, "refresh.approved_groups.field_rows", stageState)
            set currentStage of stageState to "refresh.capture_snapshot"
            set snapshot to my captureSnapshot(doc, allowedToken, storyPlan, stageState)
            set end of snapshots to snapshot
            if roundIndex > 1 then
                set converged to my sameTuple(snapshot, item (roundIndex - 1) of snapshots)
            end if
            if converged then exit repeat
        end repeat
        if not converged then error "full_tuple_not_converged_after_three_rounds" number 7156
        set currentStage of stageState to "refresh.approved_groups.count"
        set approvedResult to my approvedGroups(doc, allowedToken, storyPlan, stageState)
        set groups to item 1 of approvedResult
        set currentStage of stageState to "refresh.approved_groups.count.received"
        set currentStage of stageState to "refresh.approved_groups.count.first_use"
        set updatePairs to my countKinds(groups, "refresh.approved_groups.count", stageState)
        set currentStage of stageState to "refresh.save"
        my exactPathIndex({get posix full name of doc}, my ownedDocumentPath, false)
        save doc
        if (get saved of doc) is not true then error "calculation_copy_not_saved" number 7157
        my bindCompletedOwnedSave(get posix full name of doc)
        return {snapshots, updatePairs}
    end tell
end calculateFields

on run argv
    if (count of argv) is not 5 then error "operation, exact DOCX, PDF path, whitelist and story plan required" number 7110
    set operationName to item 1 of argv
    set inputPath to item 2 of argv
    set pdfPath to item 3 of argv
    set allowedToken to "|" & item 4 of argv & "|"
    set storyPlan to my parseStoryPlan(item 5 of argv)
    if operationName is not in {"measure_layout", "refresh_fields", "verify_only"} then error "unsupported_operation" number 7110
    set openAttempted to false
    set closeOutcome to "open_not_attempted"
    script stageState
        property currentStage : "startup"
    end script
    set ownedDocumentPath to inputPath
    set ownedDocumentIdentity to my fileIdentity(inputPath)
    set inputFileAlias to (POSIX file inputPath) as alias
    set inputPath to POSIX path of inputFileAlias
    set inputFileReference to inputFileAlias
    tell application "Microsoft Word"
        if (count of documents) is not 0 then error "user_documents_present" number 7101
        if (get background printing status) is not 0 then error "printing_active" number 7102
        set originalSecurity to get automation security
        set originalOpenLinks to get update links at open of settings
        set originalPrintFields to get update fields at print of settings
        set originalPrintLinks to get update links at print of settings
        set originalPrintCodes to get print field codes of settings
        if originalSecurity is missing value then error "security_original_unknown" number 7103
        if class of originalOpenLinks is not boolean then error "preference_original_unknown" number 7104
        repeat with flagValue in {originalPrintFields, originalPrintLinks, originalPrintCodes}
            set originalFlag to contents of flagValue
            if originalFlag is not missing value then
                if class of originalFlag is not boolean then error "preference_original_unknown" number 7104
            end if
        end repeat
        try
            with timeout of 480 seconds
                set automation security to msoAutomationSecurityForceDisable
                set update links at open of settings to false
                if originalPrintFields is not missing value then set update fields at print of settings to false
                if originalPrintLinks is not missing value then set update links at print of settings to false
                if originalPrintCodes is not missing value then set print field codes of settings to false
                if (get automation security) is not msoAutomationSecurityForceDisable then error "security_readback_failed" number 7105
                if (get update links at open of settings) is not false then error "safe_preference_readback_failed" number 7106
                set currentStage of stageState to "preopen.safe_controls"
                my assertSafePrintControls(originalPrintFields, originalPrintLinks, originalPrintCodes, "preopen.safe_controls", stageState)
                set currentStage of stageState to "preopen.safe_controls.received"
                set currentStage of stageState to "preopen.activity_check"
                if (count of documents) is not 0 or (get background printing status) is not 0 then error "user_activity_conflict" number 7107
                set currentStage of stageState to "preopen.activity_check.received"
                set openAttempted to true
                if operationName is "refresh_fields" then
                    set currentStage of stageState to "document.open.refresh"
                    open inputFileReference read only false add to recent files false
                    set currentStage of stageState to "document.open.refresh.returned"
                else
                    set currentStage of stageState to "document.open.readonly"
                    open file name inputPath read only true add to recent files false
                    set currentStage of stageState to "document.open.readonly.returned"
                end if
                set currentStage of stageState to "exact_document.lookup"
                set ownedDocument to my exactDocument(inputPath, false)
                set currentStage of stageState to "exact_document.lookup.received"
                set currentStage of stageState to "postopen.activity_check"
                if (count of documents) is not 1 or (get background printing status) is not 0 then error "user_activity_conflict" number 7107
                set currentStage of stageState to "postopen.activity_check.received"
                set currentStage of stageState to "read_only.get"
                set observedReadOnly to get read only of ownedDocument
                set currentStage of stageState to "read_only.get.received"
                set currentStage of stageState to "read_only.validation"
                if class of observedReadOnly is not boolean then error "read_only_unknown" number 7114
                if operationName is "refresh_fields" and observedReadOnly is not false then error "calculation_copy_not_editable" number 7157
                if operationName is not "refresh_fields" and observedReadOnly is not true then error "read_only_not_confirmed" number 7114
                set currentStage of stageState to "read_only.validation.received"
                set savedObservations to {get saved of ownedDocument}
                if class of item 1 of savedObservations is not boolean then error "saved_state_unknown" number 7152
                if operationName is "refresh_fields" then
                    set currentStage of stageState to "refresh.calculate"
                    set calculation to my calculateFields(ownedDocument, allowedToken, storyPlan, stageState)
                    set snapshots to item 1 of calculation
                    set updatePairs to item 2 of calculation
                    set docxSaved to true
                    set pdfExported to false
                else if operationName is "measure_layout" then
                    set currentStage of stageState to "measure.repaginate"
                    repaginate ownedDocument
                    set currentStage of stageState to "measure.capture_snapshot"
                    set snapshots to {my captureSnapshot(ownedDocument, allowedToken, storyPlan, stageState)}
                    set updatePairs to {}
                    set docxSaved to false
                    set pdfExported to false
                else
                    set currentStage of stageState to "verify.repaginate"
                    repaginate ownedDocument
                    set currentStage of stageState to "verify.capture_snapshot.before_pdf"
                    set beforeSnapshot to my captureSnapshot(ownedDocument, allowedToken, storyPlan, stageState)
                    set currentStage of stageState to "verify.all_fields.before_pdf"
                    set beforeAllFieldState to my allFieldState(ownedDocument, storyPlan, "verify.all_fields.before_pdf", stageState)
                    set end of savedObservations to get saved of ownedDocument
                    set currentStage of stageState to "verify.pdf_export"
                    save as ownedDocument file name pdfPath file format format PDF add to recent files false
                    set currentStage of stageState to "verify.pdf_export.reacquire"
                    set ownedDocument to my exactDocument(inputPath, false)
                    set currentStage of stageState to "verify.pdf_export.read_only"
                    if (get read only of ownedDocument) is not true then error "read_only_lost_after_pdf" number 7114
                    set currentStage of stageState to "verify.pdf_export.saved"
                    set end of savedObservations to get saved of ownedDocument
                    set currentStage of stageState to "verify.all_fields.after_pdf"
                    set afterAllFieldState to my allFieldState(ownedDocument, storyPlan, "verify.all_fields.after_pdf", stageState)
                    set currentStage of stageState to "verify.all_fields.compare"
                    set allFieldsStable to my sameTuple(beforeAllFieldState, afterAllFieldState)
                    if allFieldsStable is not true then error "verify_all_fields_changed" number 7159
                    set currentStage of stageState to "verify.capture_snapshot.after_pdf"
                    set afterSnapshot to my captureSnapshot(ownedDocument, allowedToken, storyPlan, stageState)
                    set end of savedObservations to get saved of ownedDocument
                    set snapshots to {beforeSnapshot, afterSnapshot}
                    set updatePairs to {}
                    set docxSaved to false
                    set pdfExported to true
                end if
                repeat with observedSaved in savedObservations
                    if class of contents of observedSaved is not boolean then error "saved_state_unknown" number 7152
                end repeat
                set finalSnapshot to item -1 of snapshots
                set pageCount to item 4 of finalSnapshot
                set tocCount to count of item 1 of finalSnapshot
                set verifiedCount to 0
                set currentStage of stageState to "verify.approved_groups"
                set approvedResult to my approvedGroups(ownedDocument, allowedToken, storyPlan, stageState)
                set groups to item 1 of approvedResult
                set storyObservations to item 2 of approvedResult
                set currentStage of stageState to "verify.approved_groups.received"
                set currentStage of stageState to "verify.approved_groups.first_use"
                set verifiedCount to my countApprovedFields(groups, "verify.approved_groups", stageState)
                set currentStage of stageState to "verify.approved_groups.count_received"
                set currentStage of stageState to "post_operation.word_version.get"
                set wordVersion to get version
                set currentStage of stageState to "post_operation.word_version.received"
                set currentStage of stageState to "post_operation.safe_controls"
                my assertSafePrintControls(originalPrintFields, originalPrintLinks, originalPrintCodes, "post_operation.safe_controls", stageState)
                set currentStage of stageState to "post_operation.safe_controls.received"
                set currentStage of stageState to "document.close.normal"
                set closeOutcome to my closeExactDocument(inputPath)
                set currentStage of stageState to "document.close.normal.returned"
                if closeOutcome is not "exact_document_closed_without_save" then error "document_close_unconfirmed" number 7115
                if (count of documents) is not 0 or (get background printing status) is not 0 then error "documents_or_printing_remain" number 7115
            end timeout
        on error failureMessage number failureNumber
            set failureStage to currentStage of stageState
            set closeErrors to ""
            if openAttempted then
                try
                    set currentStage of stageState to "document.close.cleanup"
                    set closeOutcome to my closeExactDocument(inputPath)
                    set currentStage of stageState to "document.close.cleanup.returned"
                on error closeMessage
                    set currentStage of stageState to "document.close.cleanup.failed"
                    set closeOutcome to "close_not_verified"
                    set closeErrors to closeMessage
                end try
            end if
            set restoreErrors to my restoreControls(originalSecurity, originalOpenLinks, originalPrintFields, originalPrintLinks, originalPrintCodes)
            set closeFailed to closeErrors is not ""
            set restoreFailed to restoreErrors is not ""
            error ("stage=" & failureStage & "; error_number=" & (failureNumber as string) & "; close_outcome=" & closeOutcome & "; close_failed=" & (closeFailed as string) & "; restore_failed=" & (restoreFailed as string)) number failureNumber
        end try
        set restoreErrors to my restoreControls(originalSecurity, originalOpenLinks, originalPrintFields, originalPrintLinks, originalPrintCodes)
        if restoreErrors is not "" then error ("restore_failed:" & restoreErrors) number 7109
        return my jsonValue({"success", operationName, wordVersion, docxSaved, observedReadOnly, pdfExported, closeOutcome, true, 0, pageCount, tocCount, verifiedCount, updatePairs, snapshots, savedObservations, storyObservations})
    end tell
end run
