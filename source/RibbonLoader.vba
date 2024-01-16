Sub LoadCustRibbon()

    Dim hFile As Long
    Dim path As String, fileName As String, ribbonXML As String, user As String
    hFile = FreeFile
    user = Environ("Username")
    path = "C:\Users\" & user & "\AppData\Local\Microsoft\Office\"
    fileName = "Excel.officeUI"
    
    ribbonXML = "<mso:customUI xmlns:mso='http://schemas.microsoft.com/office/2009/07/customui'>" & vbNewLine
    ribbonXML = ribbonXML + "  <mso:ribbon>" & vbNewLine
    ribbonXML = ribbonXML + "    <mso:qat/>" & vbNewLine
    ribbonXML = ribbonXML + "    <mso:tabs>" & vbNewLine
    ribbonXML = ribbonXML + "      <mso:tab id='casinoTab' label='Zoo Lottery' insertBeforeQ='mso:TabFormat'>" & vbNewLine
    ribbonXML = ribbonXML + "        <mso:group id='groupRunCasino' label='Zoo Camp Lottery' autoScale='true'>" & vbNewLine
    ribbonXML = ribbonXML + "          <mso:button id='initCasino'" & vbNewLine
    ribbonXML = ribbonXML + "                      imageMso='ExportTextFile'" & vbNewLine
    ribbonXML = ribbonXML + "                      size='large'" & vbNewLine
    ribbonXML = ribbonXML + "                      label='Initialize'" & vbNewLine
    ribbonXML = ribbonXML + "                      onAction='RibbonActions.cmdInitialize_onAction'" & vbNewLine
    ribbonXML = ribbonXML + "                      screentip='Initialize Spreadsheet'" & vbNewLine
    ribbonXML = ribbonXML + "                      supertip='Initialize the sheet.'/>" & vbNewLine
    
    ribbonXML = ribbonXML + "          <mso:button id='runCasinoConfig'" & vbNewLine
    ribbonXML = ribbonXML + "                      imageMso='ExportTextFile'" & vbNewLine
    ribbonXML = ribbonXML + "                      size='large'" & vbNewLine
    ribbonXML = ribbonXML + "                      label='Generate Camp Config'" & vbNewLine
    ribbonXML = ribbonXML + "                      onAction='RibbonActions.cmdGenCampConfig_onAction'" & vbNewLine
    ribbonXML = ribbonXML + "                      screentip='Camp Config'" & vbNewLine
    ribbonXML = ribbonXML + "                      supertip='Generate the camp config sheet.'/>" & vbNewLine
    
    ribbonXML = ribbonXML + "          <mso:button id='runCasinoLottery'" & vbNewLine
    ribbonXML = ribbonXML + "                      imageMso='ExportTextFile'" & vbNewLine
    ribbonXML = ribbonXML + "                      size='large'" & vbNewLine
    ribbonXML = ribbonXML + "                      label='Run Lottery'" & vbNewLine
    ribbonXML = ribbonXML + "                      onAction='cmdRollDice_onAction'" & vbNewLine
    ribbonXML = ribbonXML + "                      screentip='Roll Dice'" & vbNewLine
    ribbonXML = ribbonXML + "                      supertip='Run the Casino Roll Dice process.'/>" & vbNewLine
    ribbonXML = ribbonXML + "        </mso:group>" & vbNewLine
    
    ribbonXML = ribbonXML + "        <mso:group id='groupCasinoSettings' label='Settings' autoScale='true'>" & vbNewLine
    ribbonXML = ribbonXML + "          <mso:button id='runCasinoSettings'" & vbNewLine
    ribbonXML = ribbonXML + "                      imageMso='ControlsGallery'" & vbNewLine
    ribbonXML = ribbonXML + "                      size='large'" & vbNewLine
    ribbonXML = ribbonXML + "                      label='Lottery Settings'" & vbNewLine
    ribbonXML = ribbonXML + "                      onAction='cmdCasinoSettings_onAction'" & vbNewLine
    ribbonXML = ribbonXML + "                      screentip='Settings for Casino'" & vbNewLine
    ribbonXML = ribbonXML + "                      supertip='Configure the way Casino will run.'/>" & vbNewLine
    ribbonXML = ribbonXML + "        </mso:group>" & vbNewLine
    ribbonXML = ribbonXML + "      </mso:tab>" & vbNewLine
    ribbonXML = ribbonXML + "    </mso:tabs>" & vbNewLine
    ribbonXML = ribbonXML + "  </mso:ribbon>" & vbNewLine
    ribbonXML = ribbonXML + "</mso:customUI>"
    
    ribbonXML = Replace(ribbonXML, """", "")
    
    Open path & fileName For Output Access Write As hFile
    Print #hFile, ribbonXML
    Close hFile

End Sub

Sub ClearCustRibbon()

    Dim hFile As Long
    Dim path As String, fileName As String, ribbonXML As String, user As String
    
    hFile = FreeFile
    user = Environ("Username")
    path = "C:\Users\" & user & "\AppData\Local\Microsoft\Office\"
    fileName = "Excel.officeUI"
    
    ribbonXML = "<mso:customUI           xmlns:mso=""http://schemas.microsoft.com/office/2009/07/customui"">" & _
    "<mso:ribbon></mso:ribbon></mso:customUI>"
    
    Open path & fileName For Output Access Write As hFile
    Print #hFile, ribbonXML
    Close hFile

End Sub

