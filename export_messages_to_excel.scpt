tell application "Microsoft Excel"
	set LinkRemoval to make new workbook
	set theSheet to active sheet of LinkRemoval
	set formula of range "E1" of theSheet to "Message"
	set formula of range "D1" of theSheet to "Subject"
	set formula of range "C1" of theSheet to "To"
	set formula of range "B1" of theSheet to "From"
	set formula of range "A1" of theSheet to "Date"
end tell

tell application "Mail"
	set theRow to 2
	set theAccount to "American Airlines"
	get account theAccount
	set theMessages to messages of inbox
	repeat with aMessage in theMessages
		my SetDate(date received of aMessage, theRow, theSheet)
		my SetFrom(sender of aMessage, theRow, theSheet)
		my SetTo(address of recipient of aMessage, theRow, theSheet)
		my SetSubject(subject of aMessage, theRow, theSheet)
		my SetMessage(content of aMessage, theRow, theSheet)
		set theRow to theRow + 1
	end repeat
end tell

on SetDate(theDate, theRow, theSheet)
	tell application "Microsoft Excel"
		set theRange to "A" & theRow
		set formula of range theRange of theSheet to theDate
	end tell
end SetDate

on SetFrom(theSender, theRow, theSheet)
	tell application "Microsoft Excel"
		set theRange to "B" & theRow
		set formula of range theRange of theSheet to theSender
	end tell
end SetFrom

on SetTo(theRecipient, theRow, theSheet)
	tell application "Microsoft Excel"
		set theRange to "C" & theRow
		set formula of range theRange of theSheet to theRecipient
	end tell
end SetTo

on SetSubject(theSubject, theRow, theSheet)
	tell application "Microsoft Excel"
		set theRange to "D" & theRow
		set formula of range theRange of theSheet to theSubject
	end tell
end SetSubject

on SetMessage(theMessage, theRow, theSheet)
	tell application "Microsoft Excel"
		set theRange to "E" & theRow
		set formula of range theRange of theSheet to theMessage
	end tell
end SetMessage
