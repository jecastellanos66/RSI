Option Strict Off
Option Explicit On
Friend Class clsINI_RW
	
	Private m_strFileToOpen As String
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		Dim strFileName As String
		
		strFileName = VB6.Format(Now, "mm_dd_YYYY")
		strFileName = "C:\RSI\Ini\" & strFileName & ".ini"
		'Initialize the path of the iniFile to be open (By Default the ini file for today)
		Me.PathIniFile = strFileName
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'Purpose:       The purpose of this routine is to verify that the ini file exists
	'               if it dosn't exist we will create it with this routine
	Public Sub InitFile()
		'------------------------------------------------------------------------------------------------
		'------------------------------------------------------------------------------------------------
		'Purpose:       The purpose of this routine is to verify that the ini file exists
		'               if the ini file dosn't exist we will create it with this routine
		'
		'
		'Parameters:    N/A
		'
		'
		'Usage / Assumptions:   N/A
		'
		'
		'Example:       N/A
		'
		'
		'Nov-26, 2001 - EMJ - Initial Version
		'------------------------------------------------------------------------------------------------
		'------------------------------------------------------------------------------------------------
		
		
		Dim fileNum As Short
		Dim fNewFile As Boolean
		Dim intCtr As Short
		
		fileNum = FreeFile 'getting free file number
		
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If Dir(m_strFileToOpen) = "" Then
			FileOpen(fileNum, m_strFileToOpen, OpenMode.Output)
			fNewFile = True
		End If
		If fNewFile Then 'if it's a new file then have to fill it up with section and entries
			'in order to write to it
			PrintLine(fileNum, "[MESSAGE]")
			'        For intCtr = 1 To g_intMaxNumbOfRaces
			'          Print #fileNum, "WO" & intCtr & "="
			'          Print #fileNum, "FN" & intCtr & "="
			'          Print #fileNum, "WR" & intCtr & "="
			'          Print #fileNum, "RO" & intCtr & "="
			'          Print #fileNum, "RS" & intCtr & "="
			'          Print #fileNum, "RE" & intCtr & "="
			'          Print #fileNum, "PB" & intCtr & "="
			'          Print #fileNum, "WP" & intCtr & "="
			'          Print #fileNum, "TT" & intCtr & "="
			'          Print #fileNum, "RI" & intCtr & "="
			'        Next intCtr
		End If
		
		fNewFile = False
		FileClose(fileNum)
		
	End Sub
	
	'Parameters:    strNewEntry - new entry or current entry to write information for
	'               strSection  - section to write information to
	'               strData     - Data to be placed at entry level
	Public Sub WriteData(ByRef strNewEntry As String, ByRef strSection As String, ByRef strData As String, ByRef strPath As String)
		
		'------------------------------------------------------------------------------------------------
		'------------------------------------------------------------------------------------------------
		'Purpose:       The purpose of this routine is to write information to INI file
		'               this routine can also be used to add new entries to the section
		'
		'
		'Parameters:    strNewEntry - new entry or current entry to write information for
		'               strSection  - section to write information to
		'               strData     - Data to be placed at entry level
		'
		'
		'Usage / Assumptions:   N/A
		'
		'
		'Example:       N/A
		'
		'
		'Nov-26, 2001 - EMJ - Initial Version
		'------------------------------------------------------------------------------------------------
		'------------------------------------------------------------------------------------------------
		
		Dim iniClient As IniFile.itmIniFile
		
		iniClient = New IniFile.itmIniFile 'new instance of ini object
		
		iniClient.FileName = strPath
		
		With iniClient
			.Section = strSection 'setting up the section to look into
			.Entry = strNewEntry 'setting up the entry to modify
			.Data = strData 'when writing information to the ini file must
			'set the data propery before the writedata method is called
			'setting up information to write
			.WriteData()
		End With
		
		'UPGRADE_NOTE: Object iniClient may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		iniClient = Nothing
		
	End Sub
	
	'Parameters:    strEntry - enrty to read from
	'               strSection - section where entry data should be read from
	
	
	Public Function ReadData(ByRef strEntry As String, ByRef strSection As String) As String
		
		'------------------------------------------------------------------------------------------------
		'------------------------------------------------------------------------------------------------
		'Purpose:       The purpose of this routine is to read information from INI file
		'
		'
		'Parameters:    strEntry - enrty to read from
		'               strSection - section where entry data should be read from
		'
		'
		'Usage / Assumptions:   N/A
		'
		'
		'Example:       N/A
		'
		'
		'Nov-26, 2001 - EMJ - Initial Version
		'------------------------------------------------------------------------------------------------
		'------------------------------------------------------------------------------------------------
		
		Dim strMessage As String
		Dim iniClient As IniFile.itmIniFile
		iniClient = New IniFile.itmIniFile 'new instance of ini object
		
		iniClient.FileName = m_strFileToOpen
		
		With iniClient
			.Section = strSection 'setting up the section to look into
			.Entry = strEntry 'setting up the entry to modify
			.GetData(IniFile.IniDataTypes.StringValue) 'must first use the getdata method before we
			'can read from the ini file with the data property
			'UPGRADE_WARNING: Couldn't resolve default property of object iniClient.Data. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ReadData = .Data 'reading from the data property
		End With
		
		'UPGRADE_NOTE: Object iniClient may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		iniClient = Nothing
		
	End Function
	
	'property Let/Get procedures corresponding to the private member variable m_strFileToOpen
	
	Public Property PathIniFile() As String
		Get
			PathIniFile = m_strFileToOpen
		End Get
		Set(ByVal Value As String)
			m_strFileToOpen = Value
		End Set
	End Property
	
	'Read-Only property
	'This property will verify if the ini file exists.
	Public ReadOnly Property NewFile() As Boolean
		Get
			
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			If Dir(Me.PathIniFile) = "" Then
				NewFile = True
			End If
			
		End Get
	End Property
	
	Public Function ReadEntries(ByRef strSection As String) As String
		Dim strMessage As String
		Dim iniClient As IniFile.itmIniFile
		Dim strFileName As String
		
		strFileName = Me.PathIniFile
		
		If strFileName <> "" Then
			
			iniClient = New IniFile.itmIniFile 'new instance of ini object
			
			iniClient.FileName = strFileName 'set the 'FileName' property before using the GetEntries method.
			
			With iniClient
				.Section = strSection 'set the 'Section' property before using the GetEntries method.
				.GetEntries()
				'UPGRADE_WARNING: Couldn't resolve default property of object iniClient.Data. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ReadEntries = .Data 'the Data property will contain a string with entries seperated by null characters and terminated by two null characters.
			End With
			
			'UPGRADE_NOTE: Object iniClient may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			iniClient = Nothing
			
		End If
		
	End Function
	
	Public Sub WriteEntries(ByRef strSection As String, ByRef strData As String)
		'this method is used to create sections and entries.
		Dim iniClient As IniFile.itmIniFile
		Dim strFileName As String
		
		strFileName = Me.PathIniFile
		
		If strFileName <> "" Then
			iniClient = New IniFile.itmIniFile 'new instance of ini object
			
			iniClient.FileName = strFileName
			
			With iniClient
				.Section = strSection 'setting up the section to look into
				.Data = strData 'Set the Data property before using the WriteEntries method.
				'The Data property must be a string containing the entries
				'and data seperated by null characters and terminated
				'by two null characters. The entry and data must be formatted
				'as 'entry=data' within the string.
				.WriteEntries()
			End With
			
			'UPGRADE_NOTE: Object iniClient may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			iniClient = Nothing
			
		End If
		
	End Sub
End Class