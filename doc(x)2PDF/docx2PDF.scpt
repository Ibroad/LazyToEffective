on run {input, parameters}
	tell application id "com.microsoft.Word"
		activate
		repeat with aFile in input
			
			open aFile
			set extension to name extension of (info for aFile)
			set out to (aFile as text)
			set out to (text 1 thru ((out's length) - (extension's length) - 1) of out) & ".pdf"
			
			tell active document
				save as it file name out file format format PDF
				close saving no
			end tell
			
		end repeat
		quit
	end tell
end run
