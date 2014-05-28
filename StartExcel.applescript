on run argv
	
	set portMekko to "0"
	set commandMekko to ""
	set startMekkoExcel to true
	
	if (count of argv) > 0 then
		set portMekko to ((item 1 of argv) as string)
		if (count of argv) > 1 then
			set commandMekko to ((item 2 of argv) as string)
		end if
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
			set sourcePath_ to ((path to me as text) & "::" & xlsbName_)
			set copy_xlsbFile_ to false
			
			if not (exists file sourcePath_) then
				copy ("Does not exist resource file: " & POSIX path of sourcePath_) to stderr
				set startMekkoExcel to false
			else if not (exists file xlsbFile_) then
				set copy_xlsbFile_ to true
			else if not (size of (info for (POSIX path of sourcePath_))) = (size of (info for (POSIX path of xlsbFile_))) then
				set copy_xlsbFile_ to true
			else
				set myHash to do shell script ("md5 -q " & "\"" & POSIX path of sourcePath_ & "\"")
				set xlsbHash to do shell script ("md5 -q " & "\"" & POSIX path of xlsbFile_ & "\"")
				if not myHash = xlsbHash then
					set copy_xlsbFile_ to true
				end if
			end if
			
			if copy_xlsbFile_ then
				if application "Microsoft Excel" is running then
					tell application "Microsoft Excel" to quit
				end if
				duplicate file sourcePath_ to folder dirSupport_ replacing yes
			end if
			
		end if
		
	end tell
	
	if startMekkoExcel = true then
		tell application "Microsoft Excel"
			activate
			if not portMekko = "0" then
				run VB macro "__MekkoExcel__.xlsb!ThisWorkbook.SetMekkoReceiverPort" arg1 portMekko
			end if
			
			if commandMekko = "CleanupByMekko" then
				run VB macro "__MekkoExcel__.xlsb!ThisWorkbook.HideByMekko"
				run VB macro "__MekkoExcel__.xlsb!ThisWorkbook.CleanupByMekko"
				run VB macro "__MekkoExcel__.xlsb!ThisWorkbook.UnhideByMekko"
				
			else if commandMekko = "SendContentToMekko" then
				run VB macro "__MekkoExcel__.xlsb!ThisWorkbook.UnhideByMekko"
				run VB macro "__MekkoExcel__.xlsb!ThisWorkbook.SendContentToMekko"
				
			else
				run VB macro "__MekkoExcel__.xlsb!ThisWorkbook.UnhideByMekko"
			end if
			
		end tell
	end if
end run
