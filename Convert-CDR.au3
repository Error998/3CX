; #INDEX# =======================================================================================================================
; Title .........: ConvertCDRtoCSV
; AutoIt Version : 3.3.16.1
; Language ......: English
; Description ...: Converts a 3CX CDR.log file to a CSV file and adds column header on the first row
; Author(s) .....: Jock Henke
; Requirements...: AutoIt v3.3 +, Developed/Tested on Windows 11
; ===============================================================================================================================

#include <FileConstants.au3>

; CSV Header - Defaults to all fields selected in 3CX
Const $Header = "historyid,callid,duration,time-start,time-answered,time-end,reason-terminated,from-no,to-no,from-dn,to-dn,dial-no,reason-changed,final-number,final-dn,bill-code,bill-rate,bill-cost,bill-name,chain,missed-queue-calls,from-dispname,to-dispname,final-dispname,from-type,to-type,final-type"

ConvertCRDtoCSV()

Func ConvertCRDtoCSV()
        ; Create a constant variables in Local scope of the filepath that will be read/written to.
        Local Const $sFilePath = ".\CDR.log"
	Local Const $sFileWritePath = ".\CDR.csv"
	Local $sFileRead
	
	
	; Open the file for reading
        Local $hFileReadOpen = FileOpen($sFilePath, $FO_READ)
        If $hFileReadOpen = -1 Then
                ConsoleWrite("An error occurred opening CDR.log")
                Return False
        EndIf
	
	; Open the file for writing
        Local $hFileWriteOpen = FileOpen($sFileWritePath, $FO_OVERWRITE )
        If $hFileWriteOpen = -1 Then
                ConsoleWrite("An error occurred opening CDR.csv")
                Return False
        EndIf
	
	; Write CSV Header
	FileWriteLine($hFileWriteOpen, $Header)
	
        While 1
		; Read the log file
		$sFileRead = FileReadLine($hFileReadOpen)
		; Check if end of file has been reached
		If @error = -1 then ExitLoop
		
		; Skip blank lines
		If $sFileRead = "" Then ContinueLoop
		
		; Write each line to the CSV file
		FileWriteLine($hFileWriteOpen, $sFileRead)
	WEnd
		
	; Close the handles returned by FileOpen.
        FileClose($hFileReadOpen)
	FileClose($hFileWriteOpen)
	
EndFunc  
