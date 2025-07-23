####################################################################################################
#                      Mephisto's Active Directory User Password Change Script                     #
####################################################################################################
# The purpose of this script is to facilitate the updating of a large volume of AD User passwords  #
####################################################################################################
# This script will provide the option to get a list of users from Active Directory or load a CSV   #
# containing a list of users. Once the user list is loaded, you may mark users as "Included" in    #
# the password change operation. Once the user selection is complete, you will be given the option #
# to initiate a password reset for the selected users. You can also export the user selections to  #
# a CSV file for future use.                                                                       #
####################################################################################################
# This script generates randomized passwords for AD Users based on a set of characters. Once the   #
# password reset operation has completed, you will be prompted to export the new user credentials  #
# in a report in CSV format. If the export of new user credentials fails, the report will be       #
# printed to the PowerShell console for you to record manually.                                    #
####################################################################################################
function GeneratePassword($PWLenMin, $PWLenMax){#Function to generate a random password of a given length or range with at least one of each: upper case, lower case, number, symbol.
	$UpperChars = [char[]]'ABCDEFGHJKMNPQRSTUVWXYZ'#Valid upper case letters.
	$LowerChars = [char[]]'abcdefghjkmnpqrstuvwxyz'#Valid lower case letters.
	$NumChars = [char[]]'123456789'#Valid numbers.
	$SymChars = [char[]]'+?=@_$!*()^'#Valid symbols.
	$CurrentNewPW = [string]''#Initialize variable to store newly constructed password.
	if ($PWLenMin -eq $PWLenMax){#If the user chose to use a single value for the length of the new passwords...
		$CurrentNewPWLen = $PWLenMin#Defines the length of the resultant password depending on the desired value(s) input by the user.
	}#End if
	else{#If the user chose to use a range of values for the length of the new passwords...
		$CurrentNewPWLen = Get-Random -Minimum $PWLenMin -Maximum $PWLenMax #Defines the length of the resultant password depending on the desired value(s) input by the user.
	}#End else
	$CharsetIncludes = [string]''#Variable to keep track of which character sets have already been used in the constructed password.
	while ($CurrentNewPW.Length -lt $CurrentNewPWLen){#While the current length of the constructed password is less than the desired length...
		$CurrentCharSet = "U", "L", "N", "S" | Get-Random #Select a character set at random for the next character in the password.
		if (!($CharsetIncludes -match $CurrentCharSet)){$CharsetIncludes = $CharsetIncludes + $CurrentCharSet}#If the character set selected has yet to be used, add it to the list $CharsetIncludes.
		$CurrentNewChar = ''#Initialize variable to hold the next character of the constructed password.
		switch ($CurrentCharSet) {#Switch to check which character set was selected.
			U {$CurrentNewChar = $($UpperChars | Get-Random); $CurrentNewPW = $CurrentNewPW + $CurrentNewChar}#If the upper case character set is selected, select a character at random from that set and add it to the constructed password.
			L {$CurrentNewChar = $($LowerChars | Get-Random); $CurrentNewPW = $CurrentNewPW + $CurrentNewChar}#If the lower case character set is selected, select a character at random from that set and add it to the constructed password.
			N {$CurrentNewChar = $($NumChars | Get-Random); $CurrentNewPW = $CurrentNewPW + $CurrentNewChar}#If the number character set is selected, select a character at random from that set and add it to the constructed password.
			S {$CurrentNewChar = $($SymChars | Get-Random); $CurrentNewPW = $CurrentNewPW + $CurrentNewChar}#If the symbol character set is selected, select a character at random from that set and add it to the constructed password.
		}#End switch
	}#End while
	if ($CharsetIncludes.Length -eq 4){#If the list $CharsetIncludes is 4 characters long (and therefore all 4 character sets have been included in the constructed password)...
		return $CurrentNewPW#Return the constructed password.
	}#End if
	else{#If the list $CharsetIncludes is less than 4 characters long (and therefore all 4 character sets have NOT been included)...
		GeneratePassword $PWLenMin $PWLenMax #Generate a new password by calling this function.
	}#End else
}#End GeneratePassword function
function PWLengthInputCheck(){#Function to validate the user input for the desired length of the new passwords.
	$InputNewPWLength = Read-Host "Please enter the desired length of the new passwords (8-100, or R)"#Ask the user to input a value between 8 and 100 for the new password length, or "R" to select a range.
	if ($InputNewPWLength -in 8..100 -or $InputNewPWLength -eq "R"){#If the user input a valid response between 8 and 100 or "R"...
		return $InputNewPWLength#Return the value to be used for the length of the new passwords, or "R" for a range.
	}#End if
	else{#If the user did NOT enter a value between 8 and 100 or "R"...
		Write-Host "ERROR: Given password length `"$InputNewPWLength`" is invalid!"#Print a line explaining that the value entered is invalid.
		PWLengthInputCheck #Call this function to have the user provide new input for the desired length of the new passwords.
	}#End else
}#End PWLengthInputCheck Function.
function Report($ReportData){#Function for writing the report containing the new user credentials.
	if ($InputCSV -ne "N/A"){#If the user selected to load a CSV file...
		$ReportPathOnly = [System.IO.Path]::GetDirectoryName($InputCSV)#Get the path of the input CSV for use in the output file.
	}#End if
	else{#If the user DID NOT select to load a CSV file...
		$ReportPathOnly = Read-Host "Provide a location to save the CSV report containing the new user credentials."#Prompt the user to enter a path to store the new user credentials.
		if (!([string]::IsNullOrWhiteSpace($ReportPathOnly))){#If the user provided a non-blank path...
			while (!(Test-Path $ReportPathOnly)){#While the path provided by the user is invalid...
				Write-Host "ERROR: Invalid destination path!"#Print a line explaining that the given path is invalid.
				$ReportPathOnly = Read-Host "Provide a location to save the CSV report containing the new user credentials."#Prompt the user to enter a path to store the new user credentials.
				if (Test-Path -Path $ReportPathOnly -PathType Leaf){#If the user provided a path including a filename...
					$ReportPathOnly = [System.IO.Path]::GetDirectoryName($ReportPathOnly)#Strip out only the filepath, excluding the file name.
					if ([string]::IsNullOrWhiteSpace($ReportPathOnly)){#If the resulting file path is blank...
						Write-Host "ERROR: Invalid destination path!"#Print a line explaining that the given path is invalid.
						Report $ReportData #Call the Report function, passing the array containing the new user credentials so that the user can provide a valid output path.
					}#End if
				}#End if
			}#End while
		}#End if
		else{#If the user provided a blank path...
			Write-Host "ERROR: Invalid destination path!"#Print a line explaining that the provided path is invalid.
			Report $ReportData #Call the Report function, passing the array containing the new user credentials so that the user can provide a valid output path.
		}#End else
	}#End else
	$DateTime = (Get-Date).ToUniversalTime().ToString("MMddyyyyTHHmmssZ")#Get the current date and time.
	$ReportFilename = "PWChange-$DateTime.csv"#Set the filename of the report.
	$ReportPath = "$ReportPathOnly\$ReportFilename"#Set the full filename and path of the output report.
	$FilenameAppendCounter = 0 #Set the default value for the number we will append to the end of the filename of the report if the generated filename is already in use.
	while (Test-Path $ReportPath){#While the generated filename is already in use by an existing file...
		$FilenameAppendCounter++ #Increment the counter $FilenameAppendCounter
		$ReportFilename = "PWChange-$DateTime-$FilenameAppendCounter.csv"#Set the filename of the report with the appended counter.
		$ReportPath = "$ReportPathOnly\$ReportFilename"#Set the full filename and path of the output report.
	}#End while
	$DefaultReportPath = Read-Host "Would you like to write the report to $ReportPath ? (y,n)"#Ask the user if they want to use the default path and filename for the output report.
	if ($DefaultReportPath -eq "y" -or $DefaultReportPath -eq "Y"){#If the user responds that they do want to use the default report path...
		Try{#Try to output new credentials to report...
			$ReportData | Export-CSV -Path $ReportPath -NoTypeInformation #Export the array containing the new user credentials to a CSV file.
		}#End try
		Catch{#If an error occurs while attempting to output the report to a file...
			Write-Host "**ERROR WRITING REPORT** Please record the below credentials manually!***"#Print a line explaining that there was an error while exporting the new credentials to a report.
			$ReportData#Dump the array containing the objects with the new user credentials to the screen.
			exit #Stop execution.
		}#End catch
	}#End if
	else{#If the user replies that they DO NOT want to use the current path and filename for the output report...
		Report $ReportData #Call the Report function, passing the new user credentials, to get a new output path from the user.
	}#End else
}#End Report function
function HelpPage(){#Function for printing the help menu.
	cls #Clear the console screen.
	Write-Host "================================================================================"#Print UI border.
	Write-Host "|                                   Help Menu                                  |"#Print help menu title.
	Write-Host "================================================================================"#Print UI border.
	Write-Host "| This script is designed to get a list of users from Active Directory and     |"#Print help menu line.
	Write-Host "| allow you to select which users are to have their passwords updated.         |"#Print help menu line.
	Write-Host "================================================================================"#Print UI border.
	Write-Host "|                            User Account Selection                            |"#Print help menu subtitle.
	Write-Host "================================================================================"#Print UI border.
	Write-Host "| - = Go to previous user account.                                             |"#Print help menu line.
	Write-Host "| + = Go to next user account.                                                 |"#Print help menu line.
	Write-Host "| # = Go to specified index (Index must be in range)                           |"#Print help menu line.
	Write-Host "| ? = Search for user accounts by SamAccountName.                              |"#Print help menu line.
	Write-Host "| x = Toggle user account selection.                                           |"#Print help menu line.
	Write-Host "| S = Show user account selection summary.                                     |"#Print help menu line.
	Write-Host "| W = Write the current dataset to a CSV file.                                 |"#Print help menu line.
	Write-Host "| . = Continue with the password reset operation with the current selections.  |"#Print help menu line.
	Write-Host "| QUIT = End script execution without saving current selections.               |"#Print help menu line.
	Write-Host "| H = Show this help menu.                                                     |"#Print help menu line.
	Write-Host "================================================================================"#Print UI border.
	Read-Host "ENTER TO CONTINUE..."#Prompt the user to press Enter key to continue.
}#End HelpPage function.
function Summary($InputTable){#Function for displaying a summary of all Active Directory Users and their inclusion status.
	$TotalSelected = 0 #Initialize a variable to keep track of how many users are marked as selected.
	Write-Host "================================================================================"#Print UI border.
	Write-Host "|                        User Account Selection Summary                        |"#Print User Account Selection title.
	Write-Host "================================================================================"#Print UI border.
	foreach ($Row in $InputTable){#For every item in the list of user accounts...
		Write-Host "Active Directory User: "$Row.AccountName #Print a line displaying the current Active Directory user SamAccountName.
		Write-Host "Index: "$([array]::indexof($InputTable,$Row) + 1) #Print a line showing the index number of the current Active Directory user.
		Write-Host "Last Password Change: "$Row.LastPWChange #Print a line showing the date of the last password change for the current Active Directory user.
		Write-Host "Active Directory User Enabled: "$Row.UserEnabled #Print a line showing if the current Active Directory user is enabled.
		Write-Host "Included: "$Row.Included #Print a line showing if the current Active Directory user is to be included.
		Write-Host "================================================================================"#Print UI border.
		if ($Row.Included -eq "y"){#If the current Active Directory user is marked to be included...
			$TotalSelected++ #Increment the variable for keeping track of how many users are currently marked as included.
		}#End if
	}#End foreach
	Write-Host " "#Print UI empty line.
	Write-Host "********************************************************************************"#Print UI border.
	Write-Host "Total Active Directory Users: "$InputTable.Length #Print a line showing the total number of Active Directory Users.
	Write-Host "Total Included: "$TotalSelected #Print a line showing how many Active Directory Users are marked to be included.
	Write-Host "********************************************************************************"#Print UI border.
}#End Summary Function
function SearchUsers($UserList){#Function for searching for Active Directory users by SamAccountName.
	$QueryItems = @()#Initialize an array to keep the list of matches for the users query.
	$Query = Read-Host "Enter search query."#Prompt the user to enter a search term.
	if ($Query -match '^[a-zA-Z0-9._@-]{1,20}$'){#If the user entered a string between 1 and 20 characters long containing only valid characters for Active Directory usernames.
		foreach ($UserAccount in $UserList){#For each Active Directory user in the list of users...
			if ($UserAccount.AccountName -match $Query){#If the current Active Directory user SamAccountName matches the search query...
				$QueryItems += $UserAccount.AccountName #Add the matched users SamAccountName to the array $QueryItems.
			}#End if
		}#End foreach
	}#End if
	else{#If the user DID NOT enter a string between 1 and 20 characters long containing only valid characters for Active Directory usernames.
		Write-Host "Invalid input"#Print a line telling the user that their input is invalid.
		sleep 2 #Wait 2 seconds.
		SearchUsers $UserTable #Call the function SearchUsers and pass it the array full of user objects to allow the user to try searching again.
	}#End else
	if ($QueryItems.Length -lt 1){#If the search query turned up 0 results...
		Write-Host "No matches found for query `"$Query`"."#Print a line telling the user that their query returned no results.
		sleep 2 #Wait 2 seconds.
		SearchUsers $UserList #Call the function SearchUsers and pass it the array full of user objects to allow the user to try searching again.
	}#End if
	elseif ($QueryItems.Length -eq 1){#If the search query turned up only 1 result...
		$SelectedUserObject = $UserList | where {$_.AccountName -eq $($QueryItems[0])}#Set the $SelectedUserObject variable equal to the only matching Active Directory user object.
		return $([array]::indexof($UserList,$SelectedUserObject))#Return the index value for the $UserList array where the value is equal to the matching Active Directory user object.
	}#End elseif
	else{#If the search query turned up multiple results...
		Write-Host "Multiple users found for query `"$Query`":"#Print a line explaining that the search terms resulted in multiple matches.
		foreach ($QueryResult in $QueryItems){#For each matching Active Directory user...
			Write-Host $([array]::indexof($QueryItems,$QueryResult) + 1)". "$QueryResult #Print a line showing the SamAccountName of the matching Active Directory user object and a numerical identifier.
		}#End foreach
		$SelectedEntry = Read-Host "Please select an entry. (1..$($QueryItems.Length))"#Ask the user to select from one of the printed matchs.
		if ($SelectedEntry -in 1..$($QueryItems.Length)){#If the user selects a valid entry from the printed list of matching Active Directory user objects...
			$SelectedUserObject = $UserList | where {$_.AccountName -eq $($QueryItems[$($SelectedEntry - 1)])}#Set the $SelectedUserObject variable equal to the matching Active Directory user object selected by the user.
			return $([array]::indexof($UserList,$SelectedUserObject))#Return the index value for the $UserList array where the value is equal to the matching Active Directory user object selected by the user.
		}#End if
		else{#If the user DOES NOT select a valid entry from the printed list of matching Active Directory user objects..
			Write-Host "ERROR: Selection `"$SelectedEntry`" out of range!"#Print a line explaining that the user has entered an invalid selection.
			sleep 2 #Wait 2 seconds
			SearchUsers $UserList #Call the function SearchUsers and pass it the array full of user objects to allow the user to try searching again.
		}#End else
	}#End else
}#End SearchUsers Function
function WriteSelection($CurrentDataset){#Function for writing the list of Active Directory users to a CSV file.
	cls #Clear the console screen.
	$OutputAppendCounter = 0 #Initialize a variable for appending output filename if the default filename is in use.
	$UserListFilename = "ADUserList.csv"#Initialize a variable to contain the filename to use for the list of Active Directory users.
	Write-Host "================================================================================"#Print UI border.
	Write-Host "|                            Selection Table Output                            |"#Print the title for the Selection Table Output screen.
	Write-Host "================================================================================"#Print UI border.
	$UserListFilePath = Read-Host "Please provide an output path for the user table."#Ask the user to provide a path to save the CSV file containing the list of Active Directory users.
	if ([string]::IsNullOrWhiteSpace($UserListFilePath)){#If the user entered nothing for the output path...
		Write-Host "ERROR: Invalid output path!"#Print a line telling the user that the path they provided is not found.
		sleep 2 #Wait 2 seconds.
		WriteSelection $CurrentDataset #Call the WriteSelection function and pass it the current list of Active Directory users so the user can try again to provide a valid path for output.
	}#End if
	if (!(Test-Path $UserListFilePath)){#If the user DID NOT enter a valid path...
		Write-Host "ERROR: Specified folder `"$UserListFilePath`" not found!"#Print a line telling the user that the path they provided is not found.
		sleep 2 #Wait 2 seconds.
		WriteSelection $CurrentDataset #Call the WriteSelection function and pass it the current list of Active Directory users so the user can try again to provide a valid path for output.
	}#End if
	else{#If the user provided a valid path...
		if (Test-Path "$UserListFilePath\$UserListFilename"){#If the default filename is in use in the directory provided by the user...
			$OverwriteUserList = Read-Host "File `"$UserListFilePath\$UserListFilename`" already exists, would you like to overwrite it? (y,n)"#Prompt the user if they want to overwrite the existing file.
			if ($OverwriteUserList -ne "y" -and $OverwriteUserList -ne "Y"){#If the user responds that they DO NOT want to overwrite the existing file...
				while (Test-Path "$UserListFilePath\$UserListFilename"){#While the filename we are trying to write to already exists...
					$OutputAppendCounter++ #Add one to the variable $OutputAppendCounter.
					$UserListFilename = "ADUserList-$OutputAppendCounter.csv"#Set the new output filename to contain the appended number before the file extension.
				}#End while
			}#End if
		}#End if
		Try{#Try to output the list of Active Directory users to a CSV file...
			$CurrentDataset | Export-CSV -Path $UserListFilePath\$UserListFilename -NoTypeInformation #Export the list of Active Directory users to the location set by the user.
			Write-Host "User selection table exported to `"$UserListFilePath\$UserListFilename`"."#Print a line showing that the list of Active Directory users was exported to a file at the given location.
		}#End try
		Catch{#If an error occurs while trying to output the list of Active Directory users to a CSV file...
			Write-Host "ERROR: Failed to output user selection table to `"$UserListFilePath\$UserListFilename`"!"#Print a line telling the user that the export operation failed.
		}#End catch
	}#End else
	Read-Host "ENTER TO CONTINUE..."#Prompt the user to press Enter key to continue.
}#End WriteSelection Function
function TableSelection($List){#Function for allowing the user to browse the list of Active Directory users and select/deselect them for inclusion in the operation.
	cls #Clear the console screen.
	$HasIncludedProperty = [bool]($List[0].PSobject.Properties.name -match 'Included')#Check to see if the list of Active Directory users has the "Included" property.
	if (!($HasIncludedProperty)){#If the list of Active Directory users DOES NOT have the "Included" property.
		$InputDefaultSelection = ""#Initialize a variable to contain 'y' or 'n' depending on if the user wants to include (y) or exclude (n) all accounts in the operation by default.
		while ($InputDefaultSelection -ne "y" -and $InputDefaultSelection -ne "n"){#While the user has not selected a valid default inclusion status (y,n)...
			$InputDefaultSelection = Read-Host "Would you like to include all Active Directory users by default? (y,n)"#Ask the user if they want to include all Active Directory users by default.
		}#End while
		$List | foreach {if (!($_.Included)){$_ | Add-Member -MemberType NoteProperty -Name "Included" -Value $InputDefaultSelection}}#If it doesn't already exist, add a property to each Active Directory user object named "Included" to indicate if they are to be included in the operation and set the value of this property according to the users preference.
	}#End if
	$ListMaxIndex = $($($List.Length) - 1)#Initialize a variable to contain the maximum possible index value for the list of Active Directory users.
	$Index=0 #Initialize a variable to use as an index for iterating through the list of Active Directory users.
	while ($Index -in 0..$ListMaxIndex){#While the value of the variable $Index is within the allowed range of values for the list of Active Directory users...
		cls #Clear the console screen.
		$CurrentItem = $List | where {$_.AccountName -eq $($List[$Index]).AccountName}#Initialize a variable to contain the current Active Directory user object.
		Write-Host "================================================================================"#Print UI border.
		Write-Host "|                            User Account Selection                            |"#Print the title for the "User Account Selection" screen.
		Write-Host "================================================================================"#Print UI border.
		Write-Host " Current Active Directory User: "$($List[$Index].AccountName) #Print a line showing the SamAccountName for the currently selected Active Directory user.
		Write-Host " Current Index: "$($Index + 1) #Print a line showing the index value for the currently selected Active Directory user.
		Write-Host " Item Selected: "$($CurrentItem.Included) #Print a line showing whether the current Active Directory user is marked as included in the operation.
		Write-Host " Total Items: "$List.Length #Print a line showing the total number of Active Directory users in the list.
		Write-Host "================================================================================"#Print UI border.
		Write-Host "| - = Previous, + = Next, # = Go to, x = Toggle Select, H = Help, . = Done     |"#Print basic operations key.
		Write-Host "================================================================================"#Print UI border.
		$Reply = Read-Host "> "#Prompt the user for what to do next.
		switch ($Reply) {#Switch to check what the user wants to do next based on their input.
			'+' {$Index++; Break}#If the user enters "+", go to the next Active Directory user in the list.
			'-' {$Index--; Break}#If the user enters "-", go to the previous Active Directory user in the list.
			'x' {if($CurrentItem.Included-eq "n"){$CurrentItem.Included = 'y'}else{$CurrentItem.Included = 'n'}; Break}#If the user enters "x", toggle the "Included" flag for the current Active Directory user.
			'H' {HelpPage; Break}#If the user enters "H", call the Help function.
			'S' {cls; Summary $List; Read-Host "ENTER TO CONTINUE..."; Break}#If the user enters "S", call the Summary function and pass it the current list of Active Directory users.
			'?' {$Index = $(SearchUsers $List); Break}#If the user enters "?", Set $Index equal to to the value returned from the SearchUsers function, passing it the current list of Active Directory users.
			'#' {$DesiredIndex = Read-Host "Please enter the index you want to jump to. (1..$($List.Length))"; if ($DesiredIndex -in 1..$($List.Length)){$Index = $($DesiredIndex - 1)}else{Write-Host "ERROR: Selection `"$DesiredIndex`" is out of range!"; sleep 2}}#If the user enters "#", prompt the user to enter a valid index value to jump to in the list of Active Directory users. If the user enters a valid index value, set $Index to the appropriate value, otherwise, Print a line telling the user that their input is not a valid value for the index.
			'W' {WriteSelection $List; Break}#If the user enters "W", Call the WriteSelection function and pass it the current list of Active Directory users.
			'.' {cls; Summary $List; $ConfirmRun = Read-Host "Are you sure you want to initiate a password reset for the currently selected users? (y,n)"; if ($ConfirmRun -eq 'y' -or $ConfirmRun -eq 'Y'){Main $List; exit}; Break}#If the user enters ".", clear the console screen, print a summary of the current Acitve Directory user list and ask the user if they want to initiate a password reset for the selected users. If the user replies 'y' or 'Y', call the Main function passing the current user list. When the Main function returns, end execution.
			'QUIT' {$ConfirmQuit = Read-Host "Are you sure you want to exit? (y,n)"; if ($ConfirmQuit = 'y' -or $ConfirmQuit -eq 'Y'){exit}; Break}#If the user enters "QUIT", ask the user if they really want to stop execution. If they respond "y" or "Y", stop execution.
		}#End switch
		if ($Index -gt $ListMaxIndex){#If the current value of $Index is more than allowed...
			$Index-- #Subtract 1 from it's value.
		}#End if
		elseif ($Index -lt 0){#If the current value of $Index is less than allowed...
			$Index++ #Add 1 to its value.
		}#End else
	}#End while
}#End TableSelection Function
function Main(){#Function for resetting Active Directory user passwords for selected users.
	Param([Parameter(Mandatory=$false)][array]$ADUserData)#Get the list of Active Directory users if it was passed to this function, assign it the name $ADUserData
	$InputCSV = "N/A" #Initialize the variable for holding the filename and path for the loaded CSV with the placeholder value "N/A".
	$TotalResets = 0 #Initialize the variable for keeping track of how many passwords have been reset.
	$TotalAlreadyChanged = 0 #Initialize the variable for keeping track of how many user accounts are found to already have changed their password since the given date.
	$InvalidUsers = 0 #Initialize the variable for keeping track of how many users listed in the given CSV do not appear in Active Directory.
	$InputPWLength = 0 #Initialize the variable for new password length and give it a value of 8.
	$PWLenMin = 0 #Initialize the variable for the minimum desired length for new passwords.
	$PWLenMax = 0 #Initialize the variable for the maximum desired length for new passwords.
	$NewUserCredentials = @()#Initialize an array to store the objects containing the new credentials for the Active Directory users.
	if (!($ADUserData)){#If the list of Active Directory users is empty...
		$ADUserData = @()#Initialize an array to store the Active Directory user objects.
		$InputCSV = Read-Host "Enter path to the CSV file containing the list of Active Directory users"#Get the file path of the CSV containing Active Directory user accounts from the user.
		if (Test-Path $InputCSV){#If the path provided by the user for the CSV containing Active Directory usernames exists...
			$ADUserDataCSV = Import-Csv -Path $InputCSV #Import the content of the CSV containing Active Directory usernames.
		}#End if
		else{#If the path provided by the user for the CSV containing Active Directory usernames DOES NOT exist...
			Write-Host "ERROR: CSV file containing the list of Active Directory users not found!"#Print a line explaining that the given CSV file is not found.
			exit #Stop execution.
		}#End else
		$CSVADUserEnabled = "UNKNOWN"#Initialize a variable to contain the "Enabled" status of the current Active Directory user, set the default value to "UNKNOWN".
		$CSVADULastPWChange = "UNKNOWN"#Initialize a variable to contain the Last password change of the current Active Directory user, set the default value to "UNKNOWN".
		foreach ($CSVADUser in $ADUserDataCSV){#for every user listed in the loaded CSV...
			if (!(([string]::IsNullOrWhiteSpace($($CSVADUser.AccountName)))) -and !(([string]::IsNullOrWhiteSpace($($CSVADUser.Included))))){#If the "AccountName" or "Included" values are NOT missing...
				if (!(([string]::IsNullOrWhiteSpace($($CSVADUser.UserEnabled))))){#If the "UserEnabled" value is NOT missing...
					$CSVADUserEnabled = $CSVADUser.UserEnabled #Set the value of the variable $CSVADUserEnabled equal to the value of the "UserEnabled" value for the current user.
				}#End if
				if (!(([string]::IsNullOrWhiteSpace($($CSVADUser.LastPWChange))))){#If the "LastPWChange" value is NOT missing...
					$CSVADULastPWChange = $CSVADUser.LastPWChange #Set the value of the variable $CSVADULastPWChange equal to the value of the "LastPWChange" value for the current user.
				}#End if
				$ADUserData += [PSCustomObject]@{"AccountName"=$CSVADUser.AccountName;"LastPWChange"=$CSVADULastPWChange;"UserEnabled"=$CSVADUserEnabled;"Included"=$CSVADUser.Included}#Build the user object for the current user and add it to the array $ADUserData.
			}#End if
		}#End foreach
		cls #Clear the console screen.
		Summary $ADUserData #Call the Summary function and pass it the current list of Active Directory user objects.
		$ModifyLoaded = Read-Host "Would you like to modify the loaded Active Directory user selections? (y,n)"#Ask the user if they want to modify the loaded selections.
		if ($ModifyLoaded -eq "y" -or $ModifyLoaded -eq "Y"){#If the user responds that they want to edit the loaded selections...
			$ADUserData = TableSelection $ADUserData #Set the new value of the array containing the Active Directory user objects equal to the returned value from the TableSelection function, passing it the current array containing the Active Directory user objects.
		}#End if
	}#End if
	$Verbose = $False #Initialize the variable for verbose logging.
	$VerboseReply = Read-Host "Enable verbose change logging? (y,n)"#Prompt user to enable verbose logging.
	if ($VerboseReply -eq 'y' -or $VerboseReply -eq 'Y'){#If the user replies that they want to enable verbose logging...
		$Verbose = $True #Set the variable for verbose logging to True.
		$CurrentDate = (Get-Date).ToUniversalTime().ToString("MMddyyyy")#Get the current date in MonthDayYear format.
		$VerboseLogfileName = "ADPWCH-VERBLOG-$CurrentDate.txt"#Initialize the variable for the filename of the verbose log.
		Try{#Try to read and write to the location for the verbose log.
			if (!(Test-Path "$PSScriptRoot\$VerboseLogfileName")){#If the filename set for the verbose log is not already in use...
				Add-Content -Path "$PSScriptRoot\$VerboseLogfileName" "######START VERBOSE LOG FOR ADPWCH.ps1######"#Write the header to the verbose log.
			}#End if
			else{#If the filename set for the verbose log is already in use...
				$VerbLogOverwriteReply = Read-Host "File $PSScriptRoot\$VerboseLogfileName already exists, overwrite? (y,n)"#Prompt the user to overwrite the file currently using the filename set for the verbose log.
				if ($VerbLogOverwriteReply -eq 'y' -or $VerbLogOverwriteReply -eq 'Y'){#If the user replies that they do want to overwrite the existing file...
					Remove-Item -Path "$PSScriptRoot\$VerboseLogfileName"#Delete the existing file that is using the filename set for the verbose log.
					Add-Content -Path "$PSScriptRoot\$VerboseLogfileName" "######START VERBOSE LOG FOR ADPWCH.ps1######"#Write the header to the verbose log.
				}#End if
				else{#If the user replies that they do not want to overwrite the file currently using the filename set for the verbose log.
					Write-Host "Execution cancelled by user."#Print a line explaining that the script execution was cancelled by the user.
					exit #Stop execution.
				}#End else
			}#End else
		}#End try
		Catch{#If an error is raised during the verbose file read/write operation...
			Write-Host "ERROR: Failed to enable verbose log, check permissions in $PSScriptRoot"#Print a line explaining that there was an error creating the verbose log file in the script directory.
			exit #Stop execution.
		}#End catch
	}#End if
	$InputMinPWDate = Read-Host "Enter the earliest date acceptable for last password change (YYYY-MM-DD)"#Get a date from the user to use as the earliest acceptable last change date for user passwords.
	Try{#Try to convert the users input into a date.
		$MinLastChangeDate = [datetime]::Parse($InputMinPWDate)#Parse the user provided date into a standard date format.
	}#End try
	Catch{#If an error occurs while trying to parse the user provided date into a standard date format.
		Write-Host "ERROR: Bad date!"#Print a line explaining that there was an error trying to parse the date provided.
		exit #Stop execution.
	}#End catch
	$InputWriteEnabled = Read-Host "Would you like to write changes to Active Directory? (y,n)"#Ask the user if they want to allow this script to change user passwords on Active Directory during execution.
	if ($InputWriteEnabled -eq "y" -or $InputWriteEnabled -eq "Y"){#If the user responds that they want to allow changes to Active Directory...
		$WriteEnabled = $True#Set the variable that will be used to determine if we will write changes to Active Directory to true.
		$InputPWLength = PWLengthInputCheck #Call the function to ask the user to provide a value for the length of the new passwords.
		if ($InputPWLength -eq "R"){#If the user has selected that they would like to use a range for the length of new passwords...
			$ValidRange = $False#Initialize a variable to keep track of if the user has entered a valid range for password lengths, set the default value to "False".
			while (!($ValidRange)){#While the variable indicating that the user has entered a valid range of values for the length of the new passwords...
				$PWLenMin = Read-Host "Input MINimum desired length for new passwords (8-99)"#Ask the user to input a value for the minimum desired password length.
				$PWLenMax = Read-Host "Input MAXimum desired length for new passwords (9-100)"#Ask the user to input a value for the maximum desired password length.
				if ([int]$PWLenMin -ge [int]$PWLenMax){#If the user entered a minimum length greater than the maximum length...
					Write-Host "ERROR: Minimum length must not be equal or larger than maximum length!"#Print a line telling the user that minimum password length must be less than maximum password length.
				}#End if
				elseif (!($PWLenMin -in 8..99)){#If the selected minimum password length is not in range...
					Write-Host "ERROR: Minimum password length must be between 8 and 99 inclusive!"#Print a line explaining that the selected minimum password length value is invalid.
					sleep 2 #Wait 2 seconds.
				}#End elseif
				elseif (!($PWLenMax -in 9..100)){#If the selected maximum password length is not in range...
					Write-Host "ERROR: Maximum password length must be between 9 and 100 inclusive!"#Print a line explaining that the selected maximum password length value is invalid.
					sleep 2 #Wait 2 seconds.
				}#End elseif
				else{#If the user has entered valid values for the minimum and maximum desired password length...
					$ValidRange = $True#Set the variable $ValidRange to "True".
				}#End else
				cls #Clear the console screen.
			}#End while
		}#End if
		else{#If the user has selected that they would like to use a single value for the length of new passwords...
			$PWLenMin = $InputPWLength#Set the value of the minimum password length to the value input by the user.
			$PWLenMax = $InputPWLength#Set the value of the maximum password length to the value input by the user.
		}#End else
	}#End if
	else{#If the user responds that they DO NOT want to allow changes to Active Directory...
		$WriteEnabled = $False#Set the variable that will be used to determine if we will write changes to Active Directory to false.
	}#End else
	cls #Clear the console screen.
	Write-Host "================================================================================"#Print UI border.
	Write-Host "***************************Please Confirm Selections****************************"#Print UI header.
	Write-Host "================================================================================"#Print UI border.
	Write-Host "CSV Path: $InputCSV"#Print the path of the file that the user provided to be used as the source of Active Directory users.
	Write-Host "Oldest Acceptable Password Date: $MinLastChangeDate"#Print the date that the user provided to be used as the earliest acceptable last change date for user passwords.
	if ($Verbose){#If verbose logging is enabled...
		Write-Host "Verbose change logging enabled: True"#Print a line showing that verbose logging is enabled.
	}#End if
	else{#If verbose logging is not enabled...
		Write-Host "Verbose change logging enabled: False"#Print a line showing that verbose logging is disabled.
	}#End else
	if ($WriteEnabled){#If the user responded that they DO want to enable changes to Active Directory...
		if ($InputPWLength -eq "R"){#If the user chose to use a range for the length of the new user passwords...
			Write-Host "New Password Minimum Length: $PWLenMin"#Print a line showing the minimum desired password length.
			Write-Host "New Password Minimum Length: $PWLenMax"#Print a line showing the maximum desired password length.
		}#End if
		else{#If the user selected to use a single value for the new user password length...
			Write-Host "New Password Length: $InputPWLength"#Print a line showing the desired length of new passwords.
		}#End else
		Write-Host "********************LIVE WRITE TO ACTIVE DIRECTORY ENABLED**********************" -ForegroundColor red -BackgroundColor white #Print using red text on a white background a warning that write is enabled.
		Write-Host "****************WARNING: USER ACCOUNT PASSWORDS WILL BE CHANGED*****************" -ForegroundColor red -BackgroundColor white #Print using red text on a white background a warning that Active Directory user passwords will be changed if execution continues.
	}#End if
	else{#If the user responded that they DO NOT want to enable changes to Active Directory...
		Write-Host "Active Directory Write Enabled: FALSE"#Print a line showing that changes will not be written to Active Directory.
	}#End else
	Write-Host "================================================================================"#Print UI border.
	Write-Host " "#Print UI empty line.
	$RunNow = Read-Host "Continue? (y,n)"#Ask user if they want to contine with execution.
	if($RunNow -eq "y" -or $RunNow -eq "Y"){#If user responded that they DO want to contine execution...
		foreach ($ADUserObject in $ADUserData){#For every user listed in the array containing Active Directory usernames...
			$UserIncluded = "False"#Initialize a variable to record if a user is included in the operations of this script.
			$UserFound = "N/A"#Initialize a variable to keep track if the current user was found in Active Directory.
			$CurrentUserModified = "False"#Initialize a variable to keep track if the current user password has been reset.
			$NewPassword = "N/A"#Set a default value for the variable $NewPassword to "N/A" to include in the report if the current user password is not changed.
			$CUDisplayName = "N/A"#Set a default value for the variable $CUDisplayName to "N/A" to include in the report if the current user is not found in Active Directory.
			$CULastChange = "N/A"#Set a default value for the variable $CULastChange to "N/A" to include in the report if the current user is not found in Active Directory.
			if (!([string]::IsNullOrWhiteSpace($ADUserObject.AccountName))){#If the current Active Directory username is not blank...
				if ($ADUserObject.Included -eq "y" -or $ADUserObject.Included -eq "Y"){#If the current user is marked to be included in the password change...
				$UserIncluded = "True"#Set the variable $UserIncluded to "True"
					Try{#Try to get the current user information from the Active Directory.
						$CurrentADUser = Get-AdUser -Identity $ADUserObject.AccountName -Properties PasswordLastSet -ErrorAction SilentlyContinue #Query the Active Directory regarding the last time the account password was changed for the given user.
						$CUDisplayName = $CurrentADUser.Name #Set the $CUDIsplayName variable equal to the display name for the user account in Active Directory.
						$CULastChange = $CurrentADUser.PasswordLastSet #Set the $CULastChange variable equal to the last date that the current user password was changed as reported by Active Directory.
						$UserFound = "True"#Set the $UserFound variable to "True" to indicate that the user was found in Active Directory.
					}#End try
					Catch{#If there is an error while attempting to get the information from Active Directory regarding the current user.
						Write-Host "The user `"$($ADUserObject.AccountName)`" not found in Active Directory."#Print a confirmation that the current user does not exist in Active Directory.
						$InvalidUsers++#Increment the number of accounts not found in Active Directory.
						$UserFound = "False"#Set the $UserFound variable to "False" to indicate that the user WAS NOT found in Active Directory.
					}#End catch
					if ($UserFound -eq "True"){#If the current user was found in Active Directory...
						if ($CurrentADUser.PasswordLastSet -lt $MinLastChangeDate) {#If the users last password change date is before the given earliest acceptable last change date for user passwords...
							if ($WriteEnabled){#If the user previously responded that they DO want to write changes to Active Directory...
								$NewPassword = $(GeneratePassword $PWLenMin $PWLenMax) #Call the function named "GeneratePassword" and pass the minimum and maximum desired length for the new password, set the returned value as the value of the variable $NewPassword which will be used as the current users new password.
								Try{#Try to set the new user password on Active Directory.
									Set-ADAccountPassword -Identity $ADUserObject.AccountName -NewPassword (ConvertTo-SecureString -AsPlainText $NewPassword -Force)#Set the current Active Directory user's password to the new password we generated.
									$TimeOfChange = (Get-Date).ToUniversalTime().ToString("MMddyyyyTHHmmssZ")#Get the current time and date in MonthDayYearHourMinuteSecond format.
									if ($Verbose){#If verbose logging is enabled...
										$LogLine = $ADUserObject.AccountName #Initialize a variable to contain the log line and set it equal to the acocunt name of the current user.
										$LogLine = "User $LogLine password updated to $NewPassword at $TimeOfChange"#Craft the rest of the verbose log entry line.
										Add-Content -Path "$PSScriptRoot\$VerboseLogfileName" $LogLine #Write the verbose log line to the verbose log.
									}#End if
									Write-Host "The password for account `"$($ADUserObject.AccountName)`" has been reset."#Print a confirmation that the current Active Directory user password has been updated.
									$CurrentUserModified = "True"#Set the variable that we are using to keep track of if we changed the user password to "True".
								}#End try
								Catch{#If an error is raised during the password reset operation for the current user.
									Write-Host "ERROR: Failed to set password '$NewPassword' for account $($ADUserObject.AccountName)"#Print a line explaining that the password reset operation failed.
									exit #Stop execution.
								}#End catch
							}#End if
							else{#If the user previously responded that they DO NOT want to write changes to Active Directory...
								Write-Host "The password for account `"$($ADUserObject.AccountName)`" would have been reset."#Print a confirmation that the current user's password would have been updated if write to Active Directoy was enabled.
							}#End else
							$TotalResets++#Increment the number of passwords that have been reset (or would have been if write was enabled).
						}#End if
						else{#If the users last password change date is after the given earliest acceptable last change date for user passwords...
							Write-Host "The password for account `"$($ADUserObject.AccountName)`" was already changed on $($CurrentADUser.PasswordLastSet)"#Print a confirmation that the current user's password was found to have been already changed after the given earliest acceptable last change date.
							$TotalAlreadyChanged++#Increment the number of accounts that were found to have already changed their password after the given earliest acceptable last change date.
						}#End else
					}#End if
				}#End if
				else{#If the current user is NOT marked to be included in the password change...
					Write-Host "Account `"$($ADUserObject.AccountName)`" is not included in this operation."#Print a line stating that the current user is marked to be excluded from this scripts operation.
				}#End else
				$NewUserCredentials += [PSCustomObject]@{"Username"=$($ADUserObject.AccountName);"UserIncluded"=$($UserIncluded);"UserFound"=$($UserFound);"DisplayName"=$($CUDisplayName);"LastPWChange"=$($CULastChange);"PasswordChanged"=$($CurrentUserModified);"NewPassword"=$($NewPassword)}#Build the object containing the new credentials for the current user and add it to the array.
			}#End if
		}#End foreach
	}#End if
	else{#If the user responds that they DO NOT want to contine with execution.
		Write-Host "Execution cancelled by user."#Print a line explaining that the script execution was cancelled by the user.
		exit #Stop execution.
	}#End else
	Write-Host "================================================================================"#Print UI border.
	Write-Host "Total Passwords Reset (If write to Active Directory enabled): $TotalResets"#Print a line showing the total number of passwords reset (if write to Active Directory enabled).
	Write-Host "Total Passwords Already Changed Since $MinLastChangeDate`: $TotalAlreadyChanged"#Print a line showing the total number of Active Directory accounts found to have already changed their passwords since the given earliest acceptable last change date.
	Write-Host "Total Users Not Found: $InvalidUsers"#Print a line showing the total number of users in the given CSV file that were not found in Active Directory.
	Write-Host "================================================================================"#Print UI border.
	Report $NewUserCredentials #Call the function to write the report containing the new user credentials to a file, passing the array of user objects and the loaded user list to it.
}#End Main function
#Start Execution Here
cls #Clear console screen.
$ADUserList = @()#Initialize an array to store the obhects containing the Active Directory user information.
$GenerateUserList = Read-Host "Would you like to connect to Active Directory to get a list of users? (y,n)"#Ask the user if they want to connect to Active Directory to get a list of users.
if ($GenerateUserList -eq 'y' -or $GenerateUserList -eq 'Y'){#If the user replies that they want to connect to Active Directory to get a list of users...
	Try{#Try to get a list of users from Active Directory.
		Get-AdUser -Filter * | Select-Object SamAccountName | foreach {$CurrentUser = Get-AdUser -Identity $_.SamAccountName -Properties PasswordLastSet -ErrorAction SilentlyContinue; $ADUserList += [PSCustomObject]@{"AccountName"=$($_.SamAccountName);"LastPWChange"=$($CurrentUser.PasswordLastSet);"UserEnabled"=$($CurrentUser.Enabled)}};#Contact Active Directory and get a list of users, add the list of users to a array containing user objects for each user.
	}#End try
	Catch{#If an error occurs while attempting to connect to Active Directory.
		Write-Host "ERROR: Failed to connect to Active Directory!"#Print a line explaining that there was an error connecting to Active Directory.
		exit #Stop execution.
	}#End catch
	TableSelection $ADUserList #Call the function TableSelection to allow the user to review and edit selections for the list of Active Directory users.
}#End if
else{#If the user replies that they DO NOT want to connect to Active Directory...
	Main #Call the Main function.
}#End else
