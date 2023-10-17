; #INDEX# =======================================================================================================================
; Title ...............: ConvertCDRtoCSV
; Version ...........: 1.2
; AutoIt Version..: 3.3.16.1
; Language .......: English
; Description .....: Converts a 3CX CDR.log file to a CSV file and adds column header on the first row
; Author(s) .......: Jock Henke
; Requirements..: AutoIt v3.3 +, Developed/Tested on Windows 11
; ===============================================================================================================================

#include <FileConstants.au3>
#include <Date.au3>


; CSV Header - Defaults to all fields selected in 3CX
Const $Header = "HistoryID,CallID,Duration,TimeStart,TimeAnswered,TimeEnd,ReasonTerminated,FromNo,ToNo,FromDN,ToDN,DialNo,ReasonChanged,FinalNumber,FinalDN,BillCode,BillRate,BillCost,BillName,Chain,MissedQueueCalls,FromDispname,ToDispname,FinalDispname,FromType,ToType,FinalType"

ConvertCRDtoCSV()
MidgeCallsAnswered()

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
			
		; Split into seperate fields
		$aFields = StringSplit($sFileRead, ",")
		
		; Wrong number of fields skip record
		If $aFields[0] <> 27 Then ContinueLoop
		
		; Calculate the call record date difference, we only want todays data
		$iDateCalc = _DateDiff('D', $aFields[4], _NowCalc())
		If $iDateCalc > 0 Then ContinueLoop

		
		; Write each line to the CSV file
		FileWriteLine($hFileWriteOpen, $sFileRead)
	WEnd
		
	; Close the handles returned by FileOpen.
        FileClose($hFileReadOpen)
	FileClose($hFileWriteOpen)
	
EndFunc  


Func MidgeCallsAnswered()
	; Create a constant variables in Local scope of the filepath that will be read/written to.
    Local Const $sFilePath = ".\CDR.log"
	Local Const $sFileWritePath = ".\CDRMidge.csv"
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
	
	
	
	$aFields = StringSplit($Header, ",")
	;For $x = 1 To $aFields[0]
	;	ConsoleWrite($x & " - " & $aFields[$x] & @CRLF)
	;Next

	; Headers
	;~ 1 - MidgeInChain *
	;~ 2 - CallID
	;~ 3 - Duration *
	;~ 4 - TimeStart *
	;~ 5 - TimeAnswered *
	;~ 6 - TimeEnd *
	;~ 7 - ReasonTerminated *
	;~ 8 - FromNo *
	;~ 9 - ToNo *
	;~ 10 - FromDN *
	;~ 11 - ToDN *
	;~ 12 - DialNo *
	;~ 13 - ReasonChanged
	;~ 14 - FinalNumber *
	;~ 15 - FinalDN *
	;~ 16 - BillCode
	;~ 17 - BillRate
	;~ 18 - BillCost
	;~ 19 - BillName
	;~ 20 - Chain *
	;~ 21 - MissedQueueCalls *
	;~ 22 - FromDispname *
	;~ 23 - ToDispname *
	;~ 24 - FinalDispname *
	;~ 25 - FromType *
	;~ 26 - ToType *
	;~ 27 - FinalType *
	
	$aFields[1] = "MidgeInChain"
	; Write CSV Header
	$NewHeader = $aFields[1] & "," & $aFields[3] & "," & $aFields[4] & "," & $aFields[5] & "," & $aFields[6] & "," & $aFields[7] & "," & $aFields[8] & "," & $aFields[9] & "," & $aFields[10] & "," & $aFields[11] & "," &$aFields[12] & "," &$aFields[14] & "," &$aFields[15] & "," & $aFields[20] & "," & $aFields[21] & "," & $aFields[22] & "," & $aFields[23] & "," & $aFields[24] & "," & $aFields[25] & "," & $aFields[26] & "," & $aFields[27] 
	
	FileWriteLine($hFileWriteOpen, $NewHeader)
		
    While 1
		; Read the log file
		$sFileRead = FileReadLine($hFileReadOpen)
		; Check if end of file has been reached
		If @error = -1 then ExitLoop
		
		; Skip blank lines
		If $sFileRead = "" Then ContinueLoop
		
		; Split into seperate fields
		$aFields = StringSplit($sFileRead, ",")
		
		; Wrong number of fields skip record
		If $aFields[0] <> 27 Then ContinueLoop
		
		; Calculate the call record date difference, we only want todays data
		$iDateCalc = _DateDiff('D', $aFields[4], _NowCalc())
		If $iDateCalc > 0 Then ContinueLoop
		
		
		; Did Midge answer the phone? 
		If StringInStr($aFields[20], "Ext.207") Then
			$aFields[1]  = "True"
		Else
			$aFields[1]  = "False"
		EndIf
		
		$data = $aFields[1] & "," & $aFields[3] & "," & $aFields[4] & "," & $aFields[5] & "," & $aFields[6] & "," & $aFields[7] & "," & $aFields[8] & "," & $aFields[9] & "," & $aFields[10] & "," & $aFields[11] & "," &$aFields[12] & "," &$aFields[14] & "," &$aFields[15] & "," & $aFields[20] & "," & $aFields[21] & "," & $aFields[22] & "," & $aFields[23] & "," & $aFields[24] & "," & $aFields[25] & "," & $aFields[26] & "," & $aFields[27] 
		
		; Write each line to the CSV file
		FileWriteLine($hFileWriteOpen, $data)
	WEnd
		
	; Close the handles returned by FileOpen.
        FileClose($hFileReadOpen)
	FileClose($hFileWriteOpen)
	
EndFunc  
