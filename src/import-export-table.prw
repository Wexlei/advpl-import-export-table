/**
 * The MIT License (MIT)
 *
 * Copyright (c) 2018 NG Informatica - TOTVS Software Partner
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */

#Include "Totvs.ch"
#Include "TopConn.ch"
#Include "Fileio.ch"

/**
 * ImpExpTable is a simple interface for import and export tables.
 *
 * @author Wexlei Silveira
 * @since 27/11/2018
**/
User Function ImpExpTable()

	Local cAppSrvr := GetAdv97()
	Local cCodEmp  := GetPvProfString("ONSTART", "Empresa", "", cAppSrvr)
	Local cCodFil  := GetPvProfString("ONSTART", "Filial", "", cAppSrvr)

	RPCSetType(3)
	RPCSetEnv(cCodEmp, cCodFil, "", "", "MNT", GetEnvServer())

	Init()

Return nil

/**
 * Init shows the visual interface and call the import/export functions
 *
 * @author Wexlei Silveira
 * @since 27/11/2018
**/
Static Function Init()

	Local oDlg
	Local aOptions := {"Import file to target table", "Export table to file"}
	Local nOption  := 1
	Local lNext    := .F.

	Define MsDialog oDlg Title "Import/export table" FROM 000,000 TO 120, 200 Colors 0, 16777215 Pixel

	@ 005, 006 Say OemtoAnsi("Select an option bellow and click Next to continue") Size 044, 007 OF oDlg Colors CLR_HBLUE Pixel

	TRadMenu():New(25, 7, aOptions, {|u| IIf (PCount() == 0, nOption, nOption := u)}, oDlg,,,,,,,, 80, 35,,,, .T.)

	@ 050, 006 Button oBt Prompt "Next" Action (lNext := .T., oDlg:End()) Size 30,08 Of oDlg Pixel
	@ 050, 038 Button oBt Prompt "Exit" Action (lNext := .F., oDlg:End()) Size 30,08 Of oDlg Pixel

	Activate MsDialog oDlg Centered

	If lNext
		If nOption == 1
			File2Table()
		Else
			Table2File()
		EndIf
	EndIf

Return nil

/**
 * File2Table is a funtion for importing tables from a file.
 *
 * TODO: Make a better description
 *
 * @Obs. Make sure the data is exactly how it should be imported,
 * specially in cases with leading zeros (e.g.: 000009).
 * @author Wexlei Silveira
 * @since 19/11/2018
 * @return nil
 *
**/
Static Function File2Table()

	Local oDlg
	Local cFile    := Replicate(" ", 100)
	Local cTable   := "   "
	Local lExit    := .F.
	Local cMask    := "CSV (comma-separated values) (*.csv)|*.csv|Text file (*.txt)|*.txt"

	Define MsDialog oDlg Title "Import table" FROM 000,000 To 120, 210 Colors 0, 16777215 Pixel

	@ 005, 006 Say OemtoAnsi("File:") Size 044, 007 Of oDlg Colors CLR_HBLUE Pixel
	@ 012, 006 MsGet cFile Picture "@!" Size 072, 008 Of oDlg HasButton Pixel
	@ 012, 080 Button oBt Prompt "..." Action cFile := cGetFile(cMask, "Open",,, .T.,, .T., .F.) Size 13,10 Of oDlg Pixel

	@ 025, 006 Say OemtoAnsi("Target table:") Size 044, 007 OF oDlg Colors CLR_HBLUE Pixel
	@ 032, 006 MsGet cTable Picture "@!" Size 072, 010 Of oDlg HasButton Pixel

	@ 050, 006 Button oBt Prompt "Back" Action (lExit := .T., oDlg:End()) Size 30,08 Of oDlg Pixel
	@ 050, 038 Button oBt Prompt "Import" Action (fImport(cFile, cTable)) Size 30,08 Of oDlg Pixel
	@ 050, 070 Button oBt Prompt "Exit" Action (oDlg:End()) Size 30,08 Of oDlg Pixel

	Activate MsDialog oDlg Centered

	If lExit
		Init()
	EndIf

Return nil

/**
 * fImport imports the selected file to a table.
 *
 * @param cFile, Character, full path of the file to be imported
 * @param cTable, Character, target table name
 * @author Wexlei Silveira
 * @since 19/11/2018
 * @return nil
**/
Static Function fImport(cFile, cTable)

	Local cLog   := "File2Table.log"
	Local nIndex := 0

	Private aTable  := {}
	Private aStruct := {}
	Private aErrors := {}

	If Empty(cFile)
		MessageBox("File path must be filled.", "Attention", 48)
		Return .F.
	ElseIf !File(cFile)
		MessageBox("File not found.", "Attention", 48)
		Return .F.
	ElseIf Empty(cTable)
		MessageBox("Table name must be filled.", "Attention", 48)
		Return .F.
	ElseIf !ChkFile(cTable)
		MessageBox("Table not found.", "Attention", 48)
		Return .F.
	EndIf

	nHandle := FCreate(cLog)
	FT_FUSE(cLog)

	FWrite(nHandle, "Importation started in: " + cValToChar(Date()) + " - " + Time() + CRLF)
	dbSelectArea(cTable)
	dbSetOrder(1)
	aStruct := DBStruct() // Loads table structure

	If !fLoadFile(cFile) // Loads all data from file to array aTable

		MessageBox("Layout not compatible with target table.", "Attention", 48) // TODO: Specify line error
		Return nil

	EndIf

	If !Empty(aTable)

		If fValidInfo() // Validates all loaded info

			fImportInfo(cTable) // Execute importation

			FWrite(nHandle, "Importation completed successfully." + CRLF)

		Else

			For nIndex := 1 To Len(aErrors)

				FWrite(nHandle, aErrors[nIndex]) // Fills the log file with errors when found

			Next nIndex

		EndIf

	EndIf

	FWrite(nHandle, "Finalized in: " + cValToChar(Date()) + " - " + Time() + CRLF + Replicate("-", 50) + CRLF)
	FClose(nHandle) // Finalizes and closes log file

	If Len(aErrors) > 0
		MessageBox("Errors where found during de proccess, check the log for details.", "Attention", 48)
	Else
		MessageBox("Importation completed successfully.", "Attention", 64)
	EndIf

	(cTable)->(DbCloseArea())

Return nil

/**
 * fLoadFile loads the selected file to an array.
 *
 * @author Wexlei Silveira
 * @since 19/11/2018
 * @param cFile, Character, full path of the file to be imported
 * @return lRet, Boolean, true if the file was successfully loaded
**/
Static Function fLoadFile(cFile)

	Local nHandle := FT_FUse(cFile)
	Local nCount  := 1
	Local lRet    := .T.

	fSeek(nHandle, 0, FS_END)
	FT_FGoTop()
	FT_fSkip() // Skip the first line, for it's a header

	While (!FT_FEof())

		aAdd(aTable, StrTokArr2(FT_FREADLN(), ";", .T.))

		If Len(aTable[nCount]) < Len(aStruct) // Validates layout
			lRet := .F.
			Exit
		EndIf

		nCount++

		FT_fSkip()

	End

	FT_fUse()
	fClose(nHandle)

Return lRet

/**
 * fValidInfo makes basic tests on the loaded data.
 *
 * @author Wexlei Silveira
 * @since 19/11/2018
 * @return Empty(aErrors), Boolean, true if it passes the validation tests
**/
Static Function fValidInfo()

	Local nLine   := 0
	Local nColumn := 0

	// Validation
	For nLine := 1 To Len(aTable) // Line by line loop

		For nColumn := 1 To Len(aStruct[nLine]) // Column by column loop

			If X3Obrigat(aStruct[nLine, 1]) // If it's a required field, it must be filled

				If aTable[nLine, nColumn] == ""
					aAdd(aErrors, "Field " + aStruct[nLine, 1] + " at line " + cValToChar(nLine + 1) + " must be filled." + CRLF)
				EndIf

			EndIf

		Next nColumn

	Next nLine

Return Empty(aErrors)

/**
 * fImportInfo inserts the data directly on the table.
 *
 * @param cTable, Character, target table name
 * @author Wexlei Silveira
 * @since 19/11/2018
 * @return nil
**/
Static Function fImportInfo(cTable)

	Local nLine     := 0
	Local nColumn   := 0
	Local cIndexKey := fDefIndex(cTable) // Assemble the expression of the index key

	For nLine := 1 To Len(aTable) // Line by line loop

		If((cTable)->(MsSeek(&cIndexKey))) // Deletes the existing line, to avoid duplicity
			RecLock(cTable, .F.)
			DbDelete()
			(cTable)->(MsUnLock())
		EndIf

		RecLock(cTable, .T.)
		For nColumn := 1 To Len(aStruct) // Column by column loop

			If aStruct[nColumn, 2] == "D"
				&(cTable + "->" + aStruct[nColumn, 1]) := SToD(aTable[nLine, nColumn])
			ElseIf aStruct[nColumn, 2] == "N"
				&(cTable + "->" + aStruct[nColumn, 1]) := Val(aTable[nLine, nColumn])
			Else
				&(cTable + "->" + aStruct[nColumn, 1]) := aTable[nLine, nColumn]
			EndIf

		Next nColumn
		(cTable)->(MsUnLock())

	Next nLine

Return nil

/**
 * fDefIndex Defines an unique index key as an expression to be dynamically executed.
 *
 * FIXME: It may occur that the table does not contain an index equal to the unique
 * index key, and the default index (1) allows more than one record to be found,
 * in which case a record already entered will be deleted.
 *
 * @param cTable, Character, target table name
 * @author Wexlei Silveira
 * @since 19/11/2018
 * @return cIndex, Character, the whole index key as an expression for macroexecution
**/
Static Function fDefIndex(cTable)

	Local nIndex    := 0
	Local cIndex    := ""
	Local cUnique   := (cTable)->(IndexKey())
	Local aIndexKey := {}

	dbSelectArea("SX2")
	dbSetOrder(1)
	If MsSeek(cTable)
		cUnique := AllTrim(SX2->X2_UNICO)
	EndIf
	SX2->(DbCloseArea())

	dbSelectArea("SIX")
	dbSetOrder(1)
	MsSeek(cTable)
	While SIX->(!EoF()) .And. SIX->INDICE == cTable

		If AllTrim(SIX->CHAVE) == cUnique // Gets a key that is equal to the unique key

			(cTable)->(dbSetOrder(Val(SIX->ORDEM)))
			Exit

		EndIf

		SIX->(DbSkip())

	End
	SIX->(DbCloseArea())

	cUnique := StrTran(cUnique, "DTOS(", "")
	cUnique := StrTran(cUnique, "STR(", "")
	cUnique := StrTran(cUnique, ")", "")
	aIndexKey := StrTokArr2(cUnique, "+")

	For nIndex := 1 To Len(aIndexKey)

		cIndex += IIf(nIndex > 1, "+", "") + "aTable[nLine, " + cValToChar(aScan(aStruct, {|x| AllTrim(Upper(x[1])) == aIndexKey[nIndex]})) + "]"

	Next nIndex

Return cIndex

/**
 * Table2File is a funtion for exporting tables to a file.
 *
 * @author Wexlei Silveira
 * @since 06/11/2018
 * @return nil
**/
Static Function Table2File()

	Local oDlg
	Local cTable   := "   "
	Local lExit    := .F.

	Define MsDialog oDlg Title "Export table" FROM 000,000 To 90, 210 Colors 0, 16777215 Pixel

	@ 005, 006 Say OemtoAnsi("Table name:") Size 044, 007 Of oDlg Colors CLR_HBLUE Pixel
	@ 015, 006 MsGet cTable Picture "@!" Size 072, 010 Of oDlg HasButton Pixel

	@ 035, 006 Button oOk Prompt "Back" Action (lExit := .T., oDlg:End()) Size 30,08 Of oDlg Pixel
	@ 035, 038 Button oOk Prompt "Export" Action (fExport(cTable)) Size 30,08 Of oDlg Pixel
	@ 035, 070 Button oOk Prompt "Exit" Action (oDlg:End()) Size 30,08 Of oDlg Pixel

	Activate MsDialog oDlg Centered

	If lExit
		Init()
	EndIf

Return

/**
 * fExport exports the selected table to a file on Temp directory.

 * @param cTable, Character, target table name
 * @author Wexlei Silveira
 * @since 06/11/2018
 * @return nil
**/
Static Function fExport(cTable)

	Local aArea      := GetArea()
	Local lAddColumn := .T.
	Local oFWMsExcel := FWMsExcelEx():New() // XML object
	Local oExcel     := MsExcel():New() // Opens a new connection with Excel
	Local cDateTime  := cValToChar(Year(Date())) + cValToChar(Month(Date())) + StrZero(Day(Date()),2) + StrTran(Time(),":","")
	Local cFile      := GetTempPath() + cTable + cDateTime + '.xml'
	Local aStruct    := {}
	Local aLine      := {}
	Local cChar      := ""
	Local nIndex     := 0

	If Empty(cTable)
		MessageBox("Table name must be filled.", "Attention", 48)
		Return .F.
	ElseIf !ChkFile(cTable)
		MessageBox("Table not found.", "Attention", 48)
		Return .F.
	EndIf

	// Loads the table structure
	DbSelectArea(cTable)
	DbSetOrder(1)
	aStruct := DBStruct()

	oFWMsExcel:AddWorkSheet(cTable) // Creates the sheet
	oFWMsExcel:AddTable(cTable, "Fields") // Creates the table

	// Creates lines and columns
	(cTable)->(DbGoTop())
	While !((cTable)->(EoF()))

		For nIndex := 1 To Len(aStruct)

			// Adds the value for each cell of the corresponding line from the array
			If aStruct[nIndex, 2] == "D"
				// Converts date fields fields in order to avoid errors on XML
				aAdd(aLine, DToS(&(cTable + "->" + aStruct[nIndex, 1])))

			ElseIf aStruct[nIndex, 2] == "C"
				// Replaces reserved XML Characters for XML escape characters
				cChar := StrTran(&(cTable + "->" + aStruct[nIndex, 1]), "&", "&amp")
				cChar := StrTran(cChar, '"', "&quot")
				cChar := StrTran(cChar, "'", "&apos")
				cChar := StrTran(cChar, "<", "&lt")
				cChar := StrTran(cChar, ">", "&gt")
				aAdd(aLine, cChar)

			Else

				aAdd(aLine, &(cTable + "->" + aStruct[nIndex, 1]))

			EndIf

				If lAddColumn // Creates columns
					oFWMsExcel:AddColumn(cTable, "Fields", aStruct[nIndex, 1], 1)
				EndIf

		Next nIndex

		lAddColumn := .F.

		oFWMsExcel:AddRow(cTable, "Fields", aLine) // Creates lines

		aLine := {}

		(cTable)->(DbSkip())

	EndDo

	// Activates and generates XML
	oFWMsExcel:Activate()
	oFWMsExcel:GetXMLFile(cFile)
	oExcel:WorkBooks:Open(cFile) // Opens the sheet
	oExcel:SetVisible(.T.)       // Shows the sheet
	oExcel:Destroy()             // Ends the task manager proccess

	(cTable)->(DbCloseArea())
	RestArea(aArea)

Return .T.
