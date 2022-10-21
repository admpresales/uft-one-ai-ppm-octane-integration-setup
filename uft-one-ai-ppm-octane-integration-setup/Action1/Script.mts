'===========================================================
'20201008 - Initial creation
'20201012 - Increased the Exist timeout for synchronization of intially bringing up the Shared Space customization.  Base 20 seconds wasn't enough.
'20201012 - Replaced regenerate access step and copy step with traditional OR, small resolution screens causing recognition issues
'20201012 - Replaced PPM Password entry with traditional OR, small resolution screens causing recognition issues
'20201012 - Updated the click on the Shared Space to have a reattempt, up to 3 tries
'20201014 - Added logic to handle if the sa@nga Octane user is brought into Settings upon login instead of as a normal user.
'20201015 - Corrected logic failure where failure occurs, but reporter event was set to micPass
'20201020 - DJ: Updated to handle changes coming in UFT One 15.0.2
'20201120 - DJ: Increased the timeout for bringing up the Shared Space configuration to be 300 seconds instead of 120 seconds.
'20201120 - DJ: Commented out the msgbox command which can cause problems with a script being executed from Jenkins
'20210111 - DJ: The traditional OR of the settings and return to application icon needed to be updated due to changes in Octane.  Updated these
'				statements to be AI statements (15.0.2 now recognizes them correctly).  GREAT example of when to transition a traditional OR statement
'				to be an AI statement, script was broken due to an application change.
'20210209 - DJ: Updated to start the mediaserver service on the UFT One host machine if it isn't running
'20221021 - DJ: Updated because the text for the shared space changed, plus converted some statements that were properties-based to AI with enhancements to the AI engine.
'				Now the entire script is AI, no properties-based statements.
'===========================================================

Dim BrowserExecutable, ParsedClipboard, ParsedClientID, ParsedClientSecret, Counter, rc, oShell

Set oShell = CreateObject ("WSCript.shell")
oShell.run "powershell -command ""Start-Service mediaserver"""
Set oShell = Nothing

While Browser("CreationTime:=0").Exist(0)   												'Loop to close all open browsers
	Browser("CreationTime:=0").Close 
Wend
BrowserExecutable = DataTable.Value("BrowserName") & ".exe"
SystemUtil.Run BrowserExecutable,"","","",3													'launch the browser specified in the data table
Set AppContext=Browser("CreationTime:=0")													'Set the variable for what application (in this case the browser) we are acting upon

'===========================================================================================
'BP:  Navigate to the Octane login page
'===========================================================================================

AppContext.ClearCache																		'Clear the browser cache to ensure you're getting the latest forms from the application
AppContext.Navigate DataTable.Value("OctaneURL")											'Navigate to the application URL
AppContext.Maximize																			'Maximize the application to give the best chance that the fields will be visible on the screen
AppContext.Sync																				'Wait for the browser to stop spinning
AIUtil.SetContext AppContext																'Tell the AI engine to point at the application

'===========================================================================================
'BP:  Log into Octane
'===========================================================================================
AIUtil("input", "Name").SetText DataTable.Value("OctaneUserID")
AIUtil("input", "Password").Type DataTable.Value("OctanePassword")
AIUtil("button", "Login").Click
AppContext.Sync																				'Wait for the browser to stop spinning
if AIUtil("search").Exist(60) Then
	Reporter.ReportEvent micPass, "Log into Octane", "The search icon displayed within 60 seconds"
Else
	Reporter.ReportEvent micFail, "Log into Octane", "The search icon did not display within 60 seconds"
End If

'===========================================================================================
'BP:  Click the settings icon, AI not recognizing, feedback submitted
'===========================================================================================
If AIUtil.FindTextBlock("SETTINGS").Exist(1) Then
	'===========================================================================================
	'BP:  Click the return to main application icon, non-standard visual element, AI not an option
	'===========================================================================================
	AIUtil("gear_settings", micAnyText, micFromTop, 1).Click
	AppContext.Sync																				'Wait for the browser to stop spinning
End If

AIUtil("gear_settings", micAnyText, micFromTop, 1).Click

'===========================================================================================
'BP:  Click the Spaces text in the drop down menu
'===========================================================================================
AIUtil.FindTextBlock("Spaces").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Default Shared Space text 
'===========================================================================================
Counter = 0
Do
	AIUtil.FindText("Default Shared Workspace").Click
	AppContext.Sync																				'Wait for the browser to stop spinning
	Counter = Counter + 1
	wait(1)
	If Counter >=3 Then
		'msgbox("Something is broken, the Epic hasn't shown up")
		Reporter.ReportEvent micFail, "Click the Default Shared Workspace text", "The Epic text didn't display within " & Counter & " attempts."
		Exit Do
	End If
Loop Until AIUtil.FindTextBlock("Epic").Exist(300)
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the API ACCESS text
'===========================================================================================
AIUtil.FindText("API ACCESS").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the ppm text 
'===========================================================================================
AIUtil.FindTextBlock("ppm").Click

'===========================================================================================
'BP:  Click the Regen text 
'===========================================================================================
AIUtil("document").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Copy text 
'===========================================================================================
AIUtil.FindTextBlock("Copy").Click

'===========================================================================================
'BP:  Parse the clipboard to get the client ID and client secret
'===========================================================================================
Set MyClipboard = CreateObject("Mercury.Clipboard")
ParsedClipboard = Split (MyClipboard.GetText)
ParsedClientID = ParsedClipboard(2)
ParsedClientID = Left(ParsedClientID, Len(ParsedClientID) - 7)
ParsedClientSecret = ParsedClipboard(4)

'===========================================================================================
'BP:  Click the OK text, this script will NOT paste the ID and secret into Octane and save it for security reasons
'===========================================================================================
AIUtil.FindTextBlock("0K").Click

'===========================================================================================
'BP:  Click the return to main application icon, non-standard visual element, AI not an option
'===========================================================================================
AIUtil("gear_settings").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the user avatar icon, non-standard visual element, AI not an option
'===========================================================================================
AIUtil("button", micAnyText, micWithAnchorOnLeft, AIUtil("down_triangle")).Click

'===========================================================================================
'BP:  Click the Logout button
'===========================================================================================
AIUtil("button", "Logout").Click
AppContext.Sync																				'Wait for the browser to stop spinning
rc = AIUtil("input", "Password").Exist

'===========================================================================================
'BP:  Navigate to the PPM login page
'===========================================================================================
AppContext.Navigate DataTable.Value("PPMURL")											'Navigate to the application URL
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Log into PPM with admin privileges
'===========================================================================================
AIUtil("input", "Usemame").Type DataTable.Value("PPMUserID")
AIUtil("input", micAnyText, micWithAnchorOnLeft, AIUtil.FindTextBlock("Password")).Type DataTable.Value("PPMPassword")
AIUtil("button", "Sign-In").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Search for integrations
'===========================================================================================
AIUtil("search").Search "integrations"

'===========================================================================================
'BP:  Click the Integrations (OPEN) text block to navigate to the integrations page
'===========================================================================================
AIUtil.FindTextBlock("Integrations (OPEN)").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Enterprise Agile text
'===========================================================================================
AIUtil.FindTextBlock("Enterprise Agile").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Octane text
'===========================================================================================
AIUtil.FindTextBlock("Octane").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Enter the Client ID into the Client ID field.  Using traditional OR as the OCR is recognizing extra characters for the label
'===========================================================================================
AIUtil("text_box", "Client ID").SetText ParsedClientID

'===========================================================================================
'BP:  Enter the Client Secret into the Client Secret field.  Using traditional OR as the OCR is recognizing extra characters for the label and need to clear the value first
'===========================================================================================
AIUtil("text_box", "Client Secret").SetText ParsedClientSecret

'===========================================================================================
'BP:  Click the Save button
'===========================================================================================
AIUtil("button", "SAVE").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click Advanced text
'===========================================================================================
AIUtil.FindTextBlock("Advanced").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click Select Agile Projects link
'===========================================================================================
AIUtil.FindTextBlock("Select Agile Projects").Click
AppContext.Sync																				'Wait for the browser to stop spinning
If AIUtil.FindText("Failed to get available agile projects ").Exist(5) Then
	Reporter.ReportEvent micFail, "Exercise Integration", "The error message of Failed to get available agile projects displayed, integration broken"
Else
	Reporter.ReportEvent micPass, "Exercise Integration", "The error message didn't display"
End If

'===========================================================================================
'BP:  Click the Cancel text
'===========================================================================================
AIUtil("button", "CANCEL", micFromTop, 1).Click
'AIUtil.FindText("CANCEL").Click (In UFT 2022, the underlying cancel button is also being recognized, so had to change to first from top)

'===========================================================================================
'BP:  CLick the profile icon
'===========================================================================================
AIUtil("profile", micAnyText, micFromTop, 1).Click

'===========================================================================================
'BP:  CLick the Sign Out text
'===========================================================================================
AIUtil.FindText("Sign Out").Click
AppContext.Sync																				'Wait for the browser to stop spinning
rc = AIUtil("button", "Sign-In").Exist

AppContext.Close


