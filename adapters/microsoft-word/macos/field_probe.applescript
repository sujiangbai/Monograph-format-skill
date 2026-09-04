-- Synthetic 010 only. The Python entry adds the unchanged 009 ownership/control
-- handlers and run body; this file never opens a document on its own.
on exactText(actual, expected, labelText)
    if class of actual is not text then error (labelText & "_not_text") number 7150
    considering case, diacriticals, hyphens, punctuation, white space
        if actual is not expected then error (labelText & "_mismatch") number 7151
    end considering
end exactText

on requireTrue(actual, labelText)
    if class of actual is not boolean then error (labelText & "_not_boolean") number 7152
    if actual is not true then error (labelText & "_false") number 7152
end requireTrue

on sameTuple(leftTuple, rightTuple)
    -- Outside tell Word: Word's "case" property shadows AppleScript's
    -- comparison consideration inside the application terminology scope.
    considering case, diacriticals, hyphens, punctuation, white space
        return leftTuple is rightTuple
    end considering
end sameTuple

on jsonValue(value)
    if class of value is list then
        set output to "["
        repeat with i from 1 to count of value
            if i > 1 then set output to output & ","
            set output to output & my jsonValue(item i of value)
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
    error "missing_or_unsupported_snapshot_value" number 7153
end jsonValue

on fieldDescriptors(fieldObjects, scopeName)
    tell application "Microsoft Word"
        set descriptors to {}
        repeat with fieldObject in fieldObjects
            set codeRange to get field code of fieldObject
            set resultRange to get result range of fieldObject
            set descriptor to {(get content of codeRange), (get field type of fieldObject), (get start of content of codeRange), (get end of content of codeRange), (get start of content of resultRange), (get end of content of resultRange)}
            set end of descriptors to descriptor
            log {"field_descriptor", scopeName, count of descriptors, descriptor}
        end repeat
        return descriptors
    end tell
end fieldDescriptors

on fixtureTopLevelOrdinals(descriptors, scopedDescriptors, expectedCodes, expectedTypes)
    -- Fixture-only partition, observed in 010 inventory observation 02.
    -- Never authorize updates by field type alone or cross-container references.
    set childCount to count of scopedDescriptors
    if childCount is not 0 and childCount is not 2 then error "toc_child_count" number 7154
    if (count of descriptors) is not 6 + childCount then error "body_field_count" number 7154
    set outerDescriptor to item 1 of descriptors
    if (count of outerDescriptor) is not 6 then error "field_descriptor_shape" number 7154
    set tocStart to item 5 of outerDescriptor
    set tocEnd to item 6 of outerDescriptor
    set outsideOrdinals to {}
    set insideCount to 0
    set previousOutsideEnd to -1
    set previousChildEnd to -1
    repeat with i from 1 to count of descriptors
        set descriptor to item i of descriptors
        if (count of descriptor) is not 6 then error "field_descriptor_shape" number 7154
        set {codeText, typeValue, codeStart, codeEnd, resultStart, resultEnd} to descriptor
        repeat with boundValue in items 3 thru 6 of descriptor
            if class of boundValue is not integer then error "field_range_type" number 7154
        end repeat
        if codeStart < 1 or codeStart > codeEnd or codeEnd >= resultStart or resultStart > resultEnd then error "field_range_order" number 7154
        set matches to 0
        repeat with otherDescriptor in descriptors
            if my sameTuple(descriptor, contents of otherDescriptor) then set matches to matches + 1
        end repeat
        if matches is not 1 then error "field_descriptor_not_unique" number 7154
        set scopedMatches to 0
        repeat with childDescriptor in scopedDescriptors
            if my sameTuple(descriptor, contents of childDescriptor) then set scopedMatches to scopedMatches + 1
        end repeat
        set insideResult to codeStart > tocStart and resultEnd < tocEnd
        if i is not 1 and insideResult and scopedMatches is 1 then
            if typeValue is not item 4 of expectedTypes then error "toc_child_not_pageref" number 7154
            if codeStart <= previousChildEnd + 1 then error "toc_child_overlap_or_order" number 7154
            set insideCount to insideCount + 1
            if not my sameTuple(descriptor, item insideCount of scopedDescriptors) then error "toc_child_scope_order" number 7154
            set previousChildEnd to resultEnd
        else
            if scopedMatches is not 0 then error "field_scope_conflict" number 7154
            if i is not 1 and (insideResult or codeStart <= tocEnd or codeStart <= previousOutsideEnd + 1) then error "field_cross_boundary_or_order" number 7154
            set end of outsideOrdinals to i
            set ordinal to count of outsideOrdinals
            if ordinal > 6 then error "extra_top_level_field" number 7154
            my exactText(codeText, item ordinal of expectedCodes, "body_field_code_" & ordinal)
            if typeValue is not item ordinal of expectedTypes then error "body_field_type" number 7154
            set previousOutsideEnd to resultEnd
        end if
    end repeat
    if (count of outsideOrdinals) is not 6 or insideCount is not childCount then error "field_partition_incomplete" number 7154
    return outsideOrdinals
end fixtureTopLevelOrdinals

on fieldInventory(doc)
    tell application "Microsoft Word"
        if (count of sections of doc) is not 2 then error "section_count" number 7154
        if (count of tables of contents of doc) is not 1 then error "toc_count" number 7154
        set allBodyFields to get fields of doc
        -- This exact list is generated from the pinned sample's approved manifest.
        set expectedCodes to __BODY_CODES__
        set expectedTypes to {field toc, field num pages, field section pages, field page ref, field ref, field quote}
        if (count of allBodyFields) < 1 then error "body_field_count" number 7154
        my exactText((get content of field code of item 1 of allBodyFields), item 1 of expectedCodes, "outer_toc_code")
        set tocResultRange to get result range of item 1 of allBodyFields
        set scopedFields to get fields of tocResultRange
        set bodyDescriptors to my fieldDescriptors(allBodyFields, "body")
        set scopedDescriptors to my fieldDescriptors(scopedFields, "toc_result")
        set bodyOrdinals to my fixtureTopLevelOrdinals(bodyDescriptors, scopedDescriptors, expectedCodes, expectedTypes)
        set bodyFields to {}
        repeat with ordinal in bodyOrdinals
            set end of bodyFields to item (contents of ordinal) of allBodyFields
        end repeat
        log {"body_field_partition", count of allBodyFields, count of bodyFields, count of scopedFields}
        repeat with i from 1 to 6
            my exactText((get content of field code of item i of bodyFields), item i of expectedCodes, "body_field_code_" & i)
        end repeat
        set selectedFields to items 1 thru 5 of bodyFields
        repeat with sectionIndex from 1 to 2
            set footerObject to get footer (section sectionIndex of doc) index header footer primary
            set footerFields to get fields of text object of footerObject
            if (count of footerFields) is not 1 then error "footer_field_count" number 7154
            my exactText((get content of field code of item 1 of footerFields), " PAGE ", "footer_field_code")
            set end of selectedFields to item 1 of footerFields
        end repeat
        set tocObject to table of contents 1 of doc
        my requireTrue((get use heading styles of tocObject), "toc_heading_sources")
        if (get use fields of tocObject) is not false then error "toc_unapproved_field_sources" number 7154
        if (get upper heading level of tocObject) is not 1 then error "toc_upper_level" number 7154
        if (get lower heading level of tocObject) is not 1 then error "toc_lower_level" number 7154
        my requireTrue((get include page numbers of tocObject), "toc_page_numbers")
        -- 010 observation 05: paragraph.style returns a Word style object.
        -- Resolve the built-in identity in this document; never compare names.
        set headingStyle to get Word style style heading1 of doc
        if class of headingStyle is not Word style then error "heading_style_class" number 7154
        my requireTrue((get built in of headingStyle), "heading_style_builtin")
        if (get style type of headingStyle) is not style type paragraph then error "heading_style_type" number 7154
        set headings to {}
        repeat with paragraphObject in (get paragraphs of doc)
            set observedStyle to get style of paragraphObject
            if observedStyle is headingStyle then
                set end of headings to get content of text object of paragraphObject
            end if
        end repeat
        if (count of headings) is not 2 then error "heading_count" number 7154
        my exactText(item 1 of headings, "Chapter Alpha" & return, "heading_1")
        my exactText(item 2 of headings, "Chapter Beta" & return, "heading_2")
        my exactText((get content of text object of bookmark "probe_alpha" of doc), "Chapter Alpha", "alpha_bookmark")
        my exactText((get content of text object of bookmark "probe_beta" of doc), "Chapter Beta", "beta_bookmark")
        my exactText((get content of result range of item 6 of bodyFields), "DO NOT REFRESH", "unapproved_quote")
        log {"approved_inventory", count of selectedFields, "unapproved_quote_unchanged"}
        return selectedFields
    end tell
end fieldInventory

on pageAt(doc, snapshotCharacterOffset, adjusted)
    if class of snapshotCharacterOffset is not integer or snapshotCharacterOffset < 0 then error "invalid_snapshot_offset" number 7155
    tell application "Microsoft Word"
        log {"page_at_before", snapshotCharacterOffset, class of snapshotCharacterOffset, adjusted}
        set pointRange to create range doc start snapshotCharacterOffset end snapshotCharacterOffset
        set actualStart to get start of content of pointRange
        set actualEnd to get end of content of pointRange
        log {"page_at_range", actualStart, actualEnd}
        if actualStart is not snapshotCharacterOffset or actualEnd is not snapshotCharacterOffset then error "snapshot_range_mismatch" number 7155
        if adjusted then
            set resultValue to get range information pointRange information type active end adjusted page number
        else
            set resultValue to get range information pointRange information type active end page number
        end if
        log {"page_at_after", resultValue, class of resultValue}
        -- Dictionary returns text. Numeric conversion is explicit, not a default.
        set pageValue to resultValue as integer
        if pageValue < 1 then error "invalid_page_number" number 7155
        return pageValue
    end tell
end pageAt

on captureSnapshot(doc, approvedFields)
    tell application "Microsoft Word"
        log {"snapshot_stage", "toc_range_before"}
        set tocRange to get result range of item 1 of approvedFields
        set tocText to get content of tocRange
        if class of tocText is not text then error "toc_result_unavailable" number 7155
        set tocStart to my pageAt(doc, (get start of content of tocRange), false)
        set tocEnd to my pageAt(doc, ((get end of content of tocRange) - 1), false)
        log {"snapshot_stage", "toc_range_after", tocStart, tocEnd}
        log {"snapshot_stage", "body_start_before"}
        set bodyStart to get start of content of text object of section 2 of doc
        set bodyPhysical to my pageAt(doc, bodyStart, false)
        set bodyLogical to my pageAt(doc, bodyStart, true)
        log {"snapshot_stage", "body_start_after", bodyPhysical, bodyLogical}
        set sectionPages to {}
        repeat with i from 1 to 2
            log {"snapshot_stage", "section_options_before", i}
            set footerObject to get footer (section i of doc) index header footer primary
            set options to get page number options of footerObject
            set startNumber to get starting number of options
            log {"snapshot_section_start", i, startNumber, class of startNumber}
            if class of startNumber is not integer or startNumber is not 1 then error "section_start_changed" number 7155
            my requireTrue((get restart numbering at section of options), "section_restart")
            set styleValue to get number style of options
            log {"snapshot_section_style", i, styleValue, class of styleValue}
            if i is 1 and styleValue is page number style lowercase roman then
                set styleName to "lowerRoman"
            else if i is 2 and styleValue is page number style arabic then
                set styleName to "decimal"
            else
                error "section_number_format_changed" number 7155
            end if
            set end of sectionPages to {i, startNumber, styleName}
            log {"snapshot_stage", "section_options_after", i}
        end repeat
        log {"snapshot_stage", "total_pages_before"}
        set totalPages to (get range information (text object of doc) information type number of pages in document) as integer
        log {"snapshot_stage", "total_pages_after", totalPages, class of totalPages}
        if totalPages < 1 then error "invalid_total_pages" number 7155
        set fieldResults to {}
        repeat with i from 2 to count of approvedFields
            log {"snapshot_stage", "scalar_result_before", i}
            set end of fieldResults to {i, (get content of field code of item i of approvedFields), (get content of result range of item i of approvedFields)}
            log {"snapshot_stage", "scalar_result_after", i, class of item 3 of item (count of fieldResults) of fieldResults}
        end repeat
        set headingPages to {}
        repeat with bookmarkName in {"probe_alpha", "probe_beta"}
            log {"snapshot_stage", "bookmark_page_before", contents of bookmarkName}
            set pos to get start of bookmark of bookmark (bookmarkName as text) of doc
            set end of headingPages to my pageAt(doc, pos, true)
            log {"snapshot_stage", "bookmark_page_after", contents of bookmarkName, item (count of headingPages) of headingPages}
        end repeat
        return {tocText, tocEnd - tocStart + 1, bodyPhysical, bodyLogical, sectionPages, totalPages, fieldResults, headingPages}
    end tell
end captureSnapshot

on calculateFields(doc)
    tell application "Microsoft Word"
        set snapshots to {}
        set converged to false
        repeat with roundIndex from 1 to 3
            -- Complete authorization check occurs before any update each round.
            set approvedFields to my fieldInventory(doc)
            log {"round_begin", roundIndex}
            my requireTrue((update field (item 1 of approvedFields)), "toc_update")
            log {"toc_updated", roundIndex}
            repaginate doc
            -- Re-resolve fields after TOC mutation; never refresh a collection.
            set approvedFields to my fieldInventory(doc)
            repeat with i from 2 to count of approvedFields
                my requireTrue((update field (item i of approvedFields)), "approved_field_update_" & i)
                log {"approved_field_updated", roundIndex, i}
            end repeat
            set approvedFields to my fieldInventory(doc)
            set snapshot to my captureSnapshot(doc, approvedFields)
            set end of snapshots to snapshot
            log {"snapshot_complete", roundIndex}
            if roundIndex > 1 then
                set converged to my sameTuple(snapshot, item (roundIndex - 1) of snapshots)
            end if
            if converged then exit repeat
        end repeat
        if not converged then error "full_tuple_not_converged_after_three_rounds" number 7156
        save doc
        my requireTrue((get saved of doc), "calculation_saved")
        log {"calculation_saved", count of snapshots}
        return snapshots
    end tell
end calculateFields
