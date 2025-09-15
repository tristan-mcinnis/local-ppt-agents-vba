#!/usr/bin/osascript
-- Load a generated VBA script into PowerPoint and run it
-- Usage: osascript scripts/run_vba.applescript TEMPLATE_PPTX VBA_SCRIPT
on run argv
    if (count of argv) < 2 then
        error "Usage: run_vba.applescript TEMPLATE_PPTX VBA_SCRIPT"
    end if
    set pptPath to POSIX file (item 1 of argv)
    set vbaPath to POSIX file (item 2 of argv)
    set vbaCode to read vbaPath
    tell application "Microsoft PowerPoint"
        activate
        open pptPath
        do visual basic vbaCode
        run vb macro "Main"
    end tell
end run
