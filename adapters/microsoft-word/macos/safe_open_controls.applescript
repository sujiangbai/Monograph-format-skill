-- Opt-in synthetic capability probe, not a production field updater.
-- Does not open documents or modify printing options, templates, or GUI policy.
-- Constants/properties come from the installed Microsoft Word Word.sdef.
on run
    tell application "Microsoft Word"
        with timeout of 30 seconds
            if (count of documents) is not 0 then error "user_documents_present" number 7101
            if (get background printing status) is not 0 then error "printing_active" number 7102
            set originalSecurity to get automation security
            set originalLinks to get update links at open of settings
            if originalSecurity is missing value then error "security_original_unknown" number 7103
            if class of originalLinks is not boolean then error "links_original_unknown" number 7104
            try
                set automation security to msoAutomationSecurityForceDisable
                set update links at open of settings to false
                set safeSecurity to get automation security
                set safeLinks to get update links at open of settings
                if safeSecurity is not msoAutomationSecurityForceDisable then error "security_readback_failed" number 7105
                if safeLinks is not false then error "links_readback_failed" number 7106
                if (count of documents) is not 0 then error "user_activity_conflict" number 7107
                if (get background printing status) is not 0 then error "printing_conflict" number 7108
            on error failureMessage number failureNumber
                -- Attempt each restoration independently, even if the other fails.
                set restoreErrors to ""
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
                error (failureMessage & "; restore_errors=" & restoreErrors) number failureNumber
            end try
            -- The normal restoration also attempts both properties on any error.
            set restoreErrors to ""
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
            if restoreErrors is not "" then error ("restore_failed:" & restoreErrors) number 7109
            return {"controls_roundtrip_pass", originalSecurity, originalLinks, safeSecurity, safeLinks, automation security, update links at open of settings, count of documents, background printing status}
        end timeout
    end tell
end run
