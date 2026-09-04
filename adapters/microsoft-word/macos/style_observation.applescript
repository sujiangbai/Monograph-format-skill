-- One explicitly authorized, read-only observation on the fixed 010 copy.
-- The entry composes this with 009's unchanged open/control/cleanup path.
on observeHeadingStyles(doc)
    tell application "Microsoft Word"
        log {"observation_stage", "paragraph_collection_before"}
        set paragraphSnapshot to get paragraphs of doc
        log {"paragraph_collection_after", class of paragraphSnapshot, count of paragraphSnapshot}
        set directMatches to 0
        set approvedParagraphs to {}
        log {"heading_enum", style heading1, class of style heading1}
        set alphaStart to get start of bookmark of bookmark "probe_alpha" of doc
        set betaStart to get start of bookmark of bookmark "probe_beta" of doc
        set alphaText to get content of text object of bookmark "probe_alpha" of doc
        set betaText to get content of text object of bookmark "probe_beta" of doc
        log {"bookmark_source_check", my sameTuple(alphaText, "Chapter Alpha"), my sameTuple(betaText, "Chapter Beta"), alphaStart, betaStart}
        set i to 0
        repeat with paragraphObject in paragraphSnapshot
            set i to i + 1
            log {"observation_stage", i, "paragraph_reference_before_class"}
            log {"paragraph_reference_class", i, class of paragraphObject}
            log {"observation_stage", i, "direct_paragraph_style_before"}
            -- Word.sdef also declares style (1695) on paragraph itself. The
            -- 04 observation located -1728 in the nested text-object path.
            set observedStyle to get style of paragraphObject
            log {"paragraph_style_raw", i, observedStyle, class of observedStyle}
            set directMatch to observedStyle is style heading1
            log {"paragraph_style_enum_match", i, directMatch}
            if directMatch then set directMatches to directMatches + 1
            log {"observation_stage", i, "nested_start_before"}
            set startPosition to get start of content of text object of paragraphObject
            log {"paragraph_start_after", i, startPosition, class of startPosition}
            set expectedText to ""
            set sourceOrdinal to 0
            if startPosition is alphaStart then
                set expectedText to "Chapter Alpha" & return
                set sourceOrdinal to 1
            else if startPosition is betaStart then
                set expectedText to "Chapter Beta" & return
                set sourceOrdinal to 2
            end if
            log {"paragraph_style", i, observedStyle, class of observedStyle, directMatch, sourceOrdinal}
            if sourceOrdinal is not 0 then
                log {"observation_stage", i, "approved_content_before"}
                set contentMatches to my sameTuple((get content of text object of paragraphObject), expectedText)
                log {"approved_paragraph_correspondence", sourceOrdinal, i, startPosition, contentMatches}
                set end of approvedParagraphs to {sourceOrdinal, i, observedStyle, contentMatches}
            end if
        end repeat
        log {"heading_observation_counts", count of paragraphSnapshot, directMatches, count of approvedParagraphs}
        -- One dictionary-backed target-object lookup; no fallback API probing.
        log {"observation_stage", "target_style_lookup_before"}
        set targetStyle to get Word style style heading1 of doc
        log {"target_style_lookup_after", targetStyle, class of targetStyle}
        set targetName to get name local of targetStyle
        set targetBuiltIn to get built in of targetStyle
        set targetType to get style type of targetStyle
        log {"heading_target_style", targetStyle, class of targetStyle, targetName, targetBuiltIn, targetType}
        repeat with entry in approvedParagraphs
            set observedStyle to item 3 of entry
            log {"approved_style_comparisons", item 1 of entry, (observedStyle is targetStyle), (observedStyle is targetName)}
        end repeat
    end tell
end observeHeadingStyles
