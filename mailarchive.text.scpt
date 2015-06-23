#!/usr/bin/env osascript
tell application "Mail"
  set selectedMessages to selection
  if (count of selectedMessages) is equal to 0 then
    display alert "No Messages Selected" message "Select the messages you want to collect before running this script."
  end if
  repeat with theMessage in selectedMessages
    set theFilename to (subject of theMessage) as string
    set theSource to source of theMessage
    display alert theSource
    -- save theMessage in theFilename as raw
  end repeat
end tell
