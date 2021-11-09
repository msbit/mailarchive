#!/usr/bin/env osascript

property _creatorID : "emal"
property _extension : "eml"
property _messageTitle : "Export selected emails to a folder"

tell application "Mail"
  try
    display dialog "Export selected message(s) as Raw Source to " & _extension & " files?" with title "Mail Archive" with icon note
  on error
    return
  end try

  set _messages to selection
  set _count to count of _messages
  if the _count is equal to 0 then
    display alert "No Messages Selected" message "Select the messages you want to collect before running this script."
    return
  end if

  set _doneCount to 0
  set _folder to choose folder
  repeat with _message in _messages
    set _attachments to mail attachments of _message
    if _attachments is not {} then
    -- repeat with _attachment in mail attachments of _message
    --   log _attachment
    -- end repeat
    end if
    set _subject to my safeName(subject of _message) as Unicode text
    set _formattedDate to my dateFormat(date sent of _message)
    set _filenameBase to (_folder & _formattedDate & " " & _subject) as Unicode text

    set _source to source of _message
    set _destination to (_filenameBase & "." & _extension) as Unicode text
    try
      set _fileID to open for access _destination with write permission
      set eof of the _fileID to 0
      write _source to _fileID
      close access _fileID
      tell application "Finder"
        try
          set the creator type of file _destination to _creatorID
        on error
          log "Could not set creator type"
        end try
      end tell
    on error _error number _number
      set _errorMessage to "Can't write message" & return & "Error " & _error & return & "Number " & _number
      log _errorMessage
      display dialog _errorMessage
      try
        close _fileID
      end try
    end try
    set _doneCount to _doneCount + 1
  end repeat
  display dialog "Successfully exported " & _doneCount & " of " & _count & " messages." & return & return & "Destination folder:" & return & (_folder as Unicode text) buttons {"OK", "Open Folder"} default button 1 with title _messageTitle with icon note
  if the button returned of the result is "Open Folder" then
    tell application "Finder" to open folder _folder
  end if
end tell

property _months : { January, February, March, April, May, June, July, August, September, October, November, December }

on pad(input)
  text -2 thru -1 of ("00" & input)
end pad

on dateFormat(_date)
  repeat with n from 1 to count of _months
    set _month to item n of _months
    if month of _date is _month then
      set output to "" & (year of _date) & "-" & pad(n) & "-" & pad(day of _date)
      set output to output & " " & pad(hours of _date) & "-" & pad(minutes of _date) & "-" & pad(seconds of _date)
      return output
    end if
  end repeat

  "XXXX-XX-XX XX-XX-XX"
end dateFormat

on range(i, j)
  set _i to ASCII number i
  set _j to ASCII number j
  set output to {}
  repeat while _i <= _j
    set output to output & ASCII character _i
    set _i to _i + 1
  end repeat

  return output
end range

property _maxLength : 255
property _safeChars : range("a", "z") & range("A", "Z") & { "-", "_", " " }

on safeName(_name)
  set output to ""
  repeat with i from 1 to the length of _name
    if item i of _name is in _safeChars then
      set output to output & character i of _name
    else
      set output to output & "_"
    end if
  end repeat
  if the length of output > _maxLength then
    set output to (characters 1 through _maxLength of output) as Unicode text
  end if
  return output
end safeName
