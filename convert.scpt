tell application "Microsoft Excel"
    activate
    delay 3 -- Give Excel time to fully start
    
    -- Set up POSIX paths
    set inputDir to "/Users/nak/Documents/open-excel/input"
    
    try
        -- Get files using POSIX path
        set excelFiles to paragraphs of (do shell script "ls \"" & inputDir & "\"")
        
        repeat with currentFile in excelFiles
            if currentFile ends with ".xls" then
                try
                    -- Construct and verify file path
                    set currentPath to (inputDir & "/" & currentFile)
                    
                    -- Check if file exists
                    do shell script "test -f " & quoted form of currentPath
                    
                    -- Open workbook with full error handling
                    set currentWorkbook to open workbook workbook file name (POSIX file currentPath)
                    delay 2
                    
                    -- Save and close with explicit saving
                    tell currentWorkbook
                        save
                        delay 1
                        close saving yes
                        delay 1
                    end tell
                    
                on error errMsg
                    log "Error with file " & currentFile & ": " & errMsg
                end try
            end if
        end repeat
        
    on error errMsg
        log "Error accessing directory: " & errMsg
    end try
    
    -- Quit Excel
    quit saving no
end tell