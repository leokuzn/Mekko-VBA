on run argv
	
	set portMekko to "0"
	set startMekkoExcel to true
	if not argv = {} then
		set portMekko to ((item 1 of argv) as string)
	end if
	
	set xlsbName_ to "__MekkoExcel__.xlsb"
	set xlsbMekko to "\"" & xlsbName_ & "!ThisWorkbook."
	
	tell application "Finder"
		set dirSupport_ to ((path to application support from user domain) as text) & "Mekko"
		
		if not (exists folder dirSupport_) then
			copy ("Does not exist folder: " & POSIX path of dirSupport_) to stderr
			set startMekkoExcel to false
		else
			set xlsmMekko_ to dirSupport_ & ":__MekkoExcel__.xlsm"
			if exists file xlsmMekko_ then
				if application "Microsoft Excel" is running then
					tell application "Microsoft Excel" to quit
				end if
				delete file xlsmMekko_
			end if
			
			set xlsbFile_ to dirSupport_ & ":" & xlsbName_
			if not (exists file xlsbFile_) then
				if application "Microsoft Excel" is running then
					tell application "Microsoft Excel" to quit
				end if
				set sourcePath_ to ((path to me as text) & "::" & xlsbName_)
				if not (exists file sourcePath_) then
					copy ("Does not exist resource file: " & POSIX path of sourcePath_) to stderr
					set startMekkoExcel to false
				else
					copy file sourcePath_ to folder dirSupport_
				end if
			end if
		end if
		
	end tell
	
	if startMekkoExcel = true then
		tell application "Microsoft Excel"
			activate
			if not portMekko = "0" then
				run VB macro "__MekkoExcel__.xlsb!ThisWorkbook.SetMekkoReceiverPort" arg1 portMekko
				run VB macro "__MekkoExcel__.xlsb!ThisWorkbook.HideByMekko"
				run VB macro "__MekkoExcel__.xlsb!ThisWorkbook.CleanupByMekko"
			end if
			run VB macro "__MekkoExcel__.xlsb!ThisWorkbook.UnhideByMekko"
		end tell
	end if
end run
