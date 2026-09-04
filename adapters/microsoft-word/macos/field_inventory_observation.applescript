-- One synthetic observation only: no scalar update, save or result-text log.
on observeFieldDescriptors(fieldObjects, phaseName, scopeName)
    tell application "Microsoft Word"
        set descriptors to {}
        repeat with i from 1 to count of fieldObjects
            set fieldObject to item i of fieldObjects
            log {"field_observation_before", phaseName, scopeName, i}
            set codeRange to get field code of fieldObject
            set codeText to get content of codeRange
            set typeValue to get field type of fieldObject
            log {"field_code_type", phaseName, scopeName, i, codeText, typeValue, class of typeValue}
            set resultRange to get result range of fieldObject
            set codeStart to get start of content of codeRange
            set codeEnd to get end of content of codeRange
            set resultStart to get start of content of resultRange
            set resultEnd to get end of content of resultRange
            log {"field_bounds", phaseName, scopeName, i, codeStart, codeEnd, resultStart, resultEnd}
            set end of descriptors to {codeText, typeValue, codeStart, codeEnd, resultStart, resultEnd}
        end repeat
        return descriptors
    end tell
end observeFieldDescriptors

on observeBodyFields(doc, phaseName)
    tell application "Microsoft Word"
        set observedFields to get fields of doc
        log {"body_field_count_observed", phaseName, count of observedFields}
        if (count of observedFields) < 1 then error "observation_no_toc" number 7158
        set tocField to item 1 of observedFields
        my exactText((get content of field code of tocField), " TOC \\o \"1-1\" ", "observed_outer_toc")
        set tocRange to get result range of tocField
        set tocStart to get start of content of tocRange
        set tocEnd to get end of content of tocRange
        set tocChildren to get fields of tocRange
        log {"toc_result_bounds", phaseName, tocStart, tocEnd, "scoped_fields", count of tocChildren}
        -- Collect both complete lists before any association/partition decision.
        set descriptors to my observeFieldDescriptors(observedFields, phaseName, "body")
        set scopedDescriptors to my observeFieldDescriptors(tocChildren, phaseName, "toc_result")
        log {"field_lists_complete", phaseName}
        if class of tocStart is not integer or class of tocEnd is not integer or tocStart > tocEnd then error "observation_toc_range" number 7158
        set outsideSignatures to {}
        set insideCount to 0
        repeat with i from 1 to count of descriptors
            set descriptor to item i of descriptors
            repeat with boundValue in items 3 thru 6 of descriptor
                if class of boundValue is not integer then error "observation_noninteger_range" number 7158
            end repeat
            set {codeText, typeValue, codeStart, codeEnd, resultStart, resultEnd} to descriptor
            if codeStart > codeEnd or codeEnd >= resultStart or resultStart > resultEnd then error "observation_range_order" number 7158
            set insideResult to codeStart > tocStart and resultEnd < tocEnd
            set scopedMatches to 0
            repeat with childDescriptor in scopedDescriptors
                if my sameTuple(descriptor, contents of childDescriptor) then set scopedMatches to scopedMatches + 1
            end repeat
            set bodyMatches to 0
            repeat with otherDescriptor in descriptors
                if my sameTuple(descriptor, contents of otherDescriptor) then set bodyMatches to bodyMatches + 1
            end repeat
            log {"field_descriptor_membership", phaseName, i, insideResult, scopedMatches, bodyMatches}
            if bodyMatches is not 1 then error "observation_duplicate_body_descriptor" number 7158
            if i is 1 then
                if scopedMatches is not 0 then error "observation_outer_toc_in_child_list" number 7158
                set end of outsideSignatures to {codeText, typeValue}
            else if insideResult and scopedMatches is 1 then
                set insideCount to insideCount + 1
            else if (not insideResult) and scopedMatches is 0 and (resultEnd < tocStart or codeStart > tocEnd) then
                set end of outsideSignatures to {codeText, typeValue}
            else
                error "observation_unknown_or_cross_boundary" number 7158
            end if
        end repeat
        if insideCount is not (count of scopedDescriptors) then error "observation_unmatched_scoped_field" number 7158
        log {"body_field_partition", phaseName, count of observedFields, count of outsideSignatures, insideCount}
        return {outsideSignatures, count of observedFields, insideCount}
    end tell
end observeBodyFields

on observeBodyFieldInventory(doc)
    tell application "Microsoft Word"
        set approvedFields to my fieldInventory(doc)
        set beforeFields to my observeBodyFields(doc, "before")
        my requireTrue((update field (item 1 of approvedFields)), "observation_toc_update")
        log {"observation_toc_updated", 1}
        repaginate doc
        set afterFields to my observeBodyFields(doc, "after")
        my requireTrue(my sameTuple(item 1 of beforeFields, item 1 of afterFields), "observation_non_toc_signatures")
        log {"non_toc_signatures_unchanged", true}
        return {item 2 of beforeFields, item 2 of afterFields, item 3 of afterFields}
    end tell
end observeBodyFieldInventory
