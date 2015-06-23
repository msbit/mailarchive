#!/usr/bin/env osascript
tell application "Mail"
  display dialog "Export selected message(s) as Raw Source to .eml files?" with title "Mail Archive" with icon note
  if the button returned of the result is "OK" then
    set theFolder to choose folder with prompt "Save Exported Messages to..." without invisibles
    set selectedMessages to selection
    if (count of selectedMessages) is equal to 0 then
      display alert "No Messages Selected" message "Select the messages you want to collect before running this script."
    end if
    repeat with theMessage in selectedMessages
      set theFilename to (subject of theMessage) as string
      set theSource to source of theMessage
      -- display alert theSource
    end repeat
  end if
end tell
