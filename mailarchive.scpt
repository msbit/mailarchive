#!/usr/bin/env osascript

property _creatorID : "emal"
property _extension : "eml"
property _maxLength : 255
property _messageTitle : "Export selected emails to a folder"

tell application "Mail"
  display dialog "Export selected message(s) as Raw Source to " & _extension & " files?" with title "Mail Archive" with icon note
  if the button returned of the result is "OK" then
    set _messages to selection
    set _count to count of _messages
    if the _count is equal to 0 then
      display alert "No Messages Selected" message "Select the messages you want to collect before running this script."
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
      set _subject to my _cleanName(subject of _message) as Unicode text
      set _formattedDate to my _dateFormat(date sent of _message)
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
  end if
end tell

to _dateFormat(_date)
  set _result to ""
  set _result to _result & year of _date & "-"
  if (month of _date is January) then
    set _result to _result & "01-"
  else if (month of _date is February) then
    set _result to _result & "02-"
  else if (month of _date is March) then
    set _result to _result & "03-"
  else if (month of _date is April) then
    set _result to _result & "04-"
  else if (month of _date is May) then
    set _result to _result & "05-"
  else if (month of _date is June) then
    set _result to _result & "06-"
  else if (month of _date is July) then
    set _result to _result & "07-"
  else if (month of _date is August) then
    set _result to _result & "08-"
  else if (month of _date is September) then
    set _result to _result & "09-"
  else if (month of _date is October) then
    set _result to _result & "10-"
  else if (month of _date is November) then
    set _result to _result & "11-"
  else if (month of _date is December) then
    set _result to _result & "12-"
  end if
  set _result to _result & day of _date & " "
  set _result to _result & hours of _date & "-"
  set _result to _result & minutes of _date & "-"
  set _result to _result & seconds of _date
  return _result as Unicode text
end _dateFormat

to _cleanName(_name)
  set _result to ""
  set _goodChars to {"a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "-", "_", " "}
  repeat with i from 1 to the length of _name
    if item i of _name is in _goodChars then
      set _result to _result & character i of _name
    else
      set _result to _result & "_" 
    end if
  end repeat
  if the length of _result > _maxLength then
    set _result to (characters 1 through _maxLength of _result) as Unicode text
  end if
  return _result
end _cleanName
