#!/usr/bin/env osascript

tell application "Mail"
  if selection is not {} then
    set _selection to selection
    mail attachments of item 1 of _selection
  end if
end tell
