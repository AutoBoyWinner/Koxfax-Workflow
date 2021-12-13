'LoadAssembly:System.Web.Extensions.dll
'LoadAssembly:System.Data.dll
'LoadAssembly:System.Data.DataSetExtensions.dll
'LoadAssembly:System.Xml.dll
' Be sure to copy the dll System.WebExtensions.dll into the AutoStore directory in "Program Files"
' SEE README.md for more details

Option Strict Off

Imports System
Imports NSi.AutoStore.Capture.DataModel
Imports System.Net
Imports System.Web
Imports System.IO
Imports System.Text
Imports Microsoft.VisualBasic.Strings
Imports System.Web.Script.Serialization
Imports System.Reflection
'Imports Newtonsoft.Json
Imports System.XML
Imports System.Data
Imports System.Data.odbc
Imports System.Diagnostics
Imports System.Data.DataSetExtensions

Module Script
	Const SECURITY_TOKEN As String = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Im5PbzNaRHJPRFhFSzFqS1doWHNsSFJfS1hFZyIsImtpZCI6Im5PbzNaRHJPRFhFSzFqS1doWHNsSFJfS1hFZyJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvc3Rhc2luZGV2LnNoYXJlcG9pbnQuY29tQDFiNDQ3NDdlLTQyM2ItNDE5Ny1iYTI4LTExNmJmYWE2YjU0ZiIsImlzcyI6IjAwMDAwMDAxLTAwMDAtMDAwMC1jMDAwLTAwMDAwMDAwMDAwMEAxYjQ0NzQ3ZS00MjNiLTQxOTctYmEyOC0xMTZiZmFhNmI1NGYiLCJpYXQiOjE2MjkyMDM1MzQsIm5iZiI6MTYyOTIwMzUzNCwiZXhwIjoxNjI5MjkwMjM0LCJpZGVudGl0eXByb3ZpZGVyIjoiMDAwMDAwMDEtMDAwMC0wMDAwLWMwMDAtMDAwMDAwMDAwMDAwQDFiNDQ3NDdlLTQyM2ItNDE5Ny1iYTI4LTExNmJmYWE2YjU0ZiIsIm5hbWVpZCI6IjFmYjlhYzczLWZmODgtNGZlZS1iMGQyLTZlMmE4ZGMwZGFiNEAxYjQ0NzQ3ZS00MjNiLTQxOTctYmEyOC0xMTZiZmFhNmI1NGYiLCJvaWQiOiI1NmEyMzYxMC1iY2M0LTQ3MzgtYTdkZS0wYTI2MGQ0ZjU1ODgiLCJzdWIiOiI1NmEyMzYxMC1iY2M0LTQ3MzgtYTdkZS0wYTI2MGQ0ZjU1ODgiLCJ0cnVzdGVkZm9yZGVsZWdhdGlvbiI6ImZhbHNlIn0.eqbMPNg6yhVwdXHE4f4jUfbV05Ya1a4GeSSvhTYMvhOwL4EVc0S63JD_h0bDjF9ExGtW6mBAp9mGdEH_5zGwzP-yEWyC9EhY6xJUW-qL8rkaLoD5z1tpVhsd-ibXzetcQRCbAj15k-9_XlA46DRnutb1r4yRmMZ5X-025JzhFvrfh1z-tPmNi2pRfzADHbQKE2ed_QuqN1A5MCB3YNQzv5UbQ8MlBvkvNMU_gcyuKjo3_lBUQH5Bie_EQOtBBMO_wAV58FFFNgb-YnvRNgJKAQ9DgJwFZSIfadrmaQTjQPMlmWASehzB3JdtOGFoV3oWX1jJrbMLyUj2H2dM5HaHJg"
	Const ODBC_CONNECTION_STRING As String = "DSN=sqltest;Trusted_Connection=YES"
    Sub Form_OnLoad(ByVal eventData As MFPEventData)
		Call InitializeInterface(eventData)
	End Sub
	
	Function ShowStatus(ByVal eventData As MFPEventData, ByVal strMSG As String)
		Dim status As TextField = eventData.Form.Fields.GetField("status")
		status.Value = strMSG	
		
	End Function
	
	Function ShowInterface(ByVal eventData As MFPEventData, ByVal bShowFlag As Boolean)
		eventData.Form.Fields.GetField("patientID").IsHidden =Not bShowFlag
		eventData.Form.Fields.GetField("nom").IsHidden =Not bShowFlag
		eventData.Form.Fields.GetField("prenom").IsHidden =Not bShowFlag
		eventData.Form.Fields.GetField("patientList").IsHidden =Not bShowFlag
		eventData.Form.Fields.GetField("searchBtn").IsHidden =Not bShowFlag
		eventData.Form.Fields.GetField("gobackBtn").IsHidden = bShowFlag
		eventData.Form.Fields.GetField("dateofbirth").IsHidden = bShowFlag
		eventData.Form.Fields.GetField("documentType").IsHidden = bShowFlag
		eventData.Form.Fields.GetField("visitdate").IsHidden = bShowFlag 
	End Function
	Sub Form_OnSubmit(ByVal eventData As MFPEventData)
        'TODO add code here to execute when the user presses OK in the form
    End Sub
	Function SetCurID(ByVal eventData As MFPEventData, ByVal id As String)
		eventData.Form.Fields.GetField("curID").Value = id
	End Function
	
	Function GetCurID(ByVal eventData As MFPEventData)
		Dim id As String = eventData.Form.Fields.GetField("curID").Value
		Return id
	End Function
	
	Sub gobackBtn_OnChange(ByVal eventData As MFPEventData)
		eventData.Form.Fields.GetField("searchBtn").Value = False
		Call InitializeInterface(eventData)
	End Sub
	
	Function InitializeInterface(ByVal eventData As MFPEventData)
		eventData.Form.Fields.GetField("patientID").Value = ""
		eventData.Form.Fields.GetField("nom").Value = ""
		eventData.Form.Fields.GetField("prenom").Value ="" 
		eventData.Form.Fields.GetField("patientList").Value =""
		eventData.Form.Fields.GetField("dateofbirth").Value = ""
		eventData.Form.Fields.GetField("documentType").Value = ""
		eventData.Form.Fields.GetField("visitdate").Value = "" 
		
		Call SetCurID(eventData, "-1")
		ShowInterface(eventData, True)
		Dim recordCount As Integer = GetPatientList(eventData,"" , "", "")
		
		If recordCount > 0 Then
			Call ShowStatus(eventData, recordCount & " Records found. Please select the Patient List.")
		Else
			Call ShowStatus(eventData, "Connect error !")
		End If
	End Function
	
	Sub searchBtn_OnChange(ByVal eventData As MFPEventData) 'When Search Check box clicked, event occurs
		eventData.Form.Fields.GetField("searchBtn").Value = False
		
		Dim patientID As String = eventData.Form.Fields.GetField("patientID").Value
		Dim nom As String = eventData.Form.Fields.GetField("nom").Value
		Dim prenom As String = eventData.Form.Fields.GetField("prenom").Value
		
		Dim recordCount As Integer = GetPatientList(eventData, patientID, nom, prenom)
		If recordCount > 1 Then
			Call ShowStatus(eventData, recordCount & " Records found. Please select the Patient List.")
		Else If recordCount = 1 Then
			Call ShowInterface(eventData, False)
			Dim curID As String = GetCurID(eventData)
			If curID > 0 Then
				Call ShowStatus(eventData, "1 Record found. Please select the Patient List.")
				Call GetPaitentInfoByID(eventData, curID)
			End If
			
		Else
			Call ShowStatus(eventData, "Not found record.")
		End If
	End Sub
	Sub patientList_OnChange(ByVal eventData As MFPEventData) 'When patient List selected, event occurs
		'Dim patient
		Dim patientList As ListField = eventData.Form.Fields.GetField("patientList")
		Dim id As String = patientList.GetSelectedItem().Value
		Call ShowInterface(eventData, False)
		Call GetPaitentInfoByID(eventData, id)
	End Sub
	Function GetPatientList(ByVal eventData As MFPEventData, ByVal patientID As String, ByVal nom As String, ByVal prenom As String)
		
		Dim connString As String	
		
		connString = ODBC_CONNECTION_STRING
		Dim conn As OdbcConnection
		Dim dataReader As OdbcDataReader
		Dim command As OdbcCommand
		conn = New System.Data.Odbc.OdbcConnection(connString)
		
		Dim  sqlString As String = ""
		If patientID <> "" Then
			sqlString = sqlString & "AND patientID2 = '" & patientID & "'"
		End If
		If nom <> "" Then
			sqlString = sqlString & "AND nom = '" & nom & "'"
		End If
		If prenom <> "" Then
			sqlString = sqlString & "AND prenom = '" & prenom & "'"
		End If
		
		Dim cmdText As String = "SELECT * FROM patient WHERE 1=1 " & sqlString 
		Dim documentList As ListField = eventData.Form.Fields.GetField("patientList")

		
		documentList.FindMode = False
		documentList.Items.Clear()
		Dim listItem As ListItem 
		
		Dim index As Integer = 0
		Dim selected_id As String = "-1"
		Dim txtString As String
		Try
			command = New OdbcCommand(cmdText)
			command.Connection = conn
			conn.Open()
			dataReader = command.ExecuteReader()
			
			
			'rowCount = dataReader.FieldCount
			documentList.FindMode = False
			documentList.Items.Clear()
			'documentList.MaxSearchResults = rowCount + 1
			
			While dataReader.Read()
				
				Dim iDate As String = dataReader.getString(2)
				Dim oDate As DateTime = Convert.ToDateTime(iDate)	
				selected_id = dataReader.getString(3)
				txtString = dataReader.getstring(4) & " - " & dataReader.getString(0) & " " & dataReader.getString(1) & " - " & oDate.Month & "/" & oDate.Day & "/" & oDate.Year
				
				listItem = New ListItem(txtString, dataReader.getString(3))
				documentList.Items.Add(listItem)
				index += 1
			End While
		
			If index = 1 Then
				Call SetCurID(eventData, selected_id)
			Else
				Call SetCurID(eventData, "-1")
			End If
			
			If Not (dataReader Is Nothing) Then
				dataReader.Close()
			End If
			conn.Close()			
		Catch ex As Exception
			'	Console.WriteLine(ex.Message)
		End Try		
		Return index	
	End Function
	
	Function GetPaitentInfoByID(ByVal eventData As MFPEventData, ByVal id As String)
		Dim connString As String	
		
		connString = ODBC_CONNECTION_STRING
		Dim conn As OdbcConnection
		Dim dataReader As OdbcDataReader
		Dim command As OdbcCommand
		conn = New System.Data.Odbc.OdbcConnection(connString)
		
		
		Dim cmdText As String = "SELECT * FROM patient WHERE ID = " & id & ""
		Dim documentList As ListField = eventData.Form.Fields.GetField("patientList")

		
		documentList.FindMode = False
		documentList.Items.Clear()
		Dim listItem As ListItem 
		
		Dim index As Integer = 0
		Dim selected_id As String = "-1"
		Dim txtString As String
		Try
			command = New OdbcCommand(cmdText)
			command.Connection = conn
			conn.Open()
			dataReader = command.ExecuteReader()
			
			
			'rowCount = dataReader.FieldCount
			documentList.FindMode = False
			documentList.Items.Clear()
			'documentList.MaxSearchResults = rowCount + 1
			
			While dataReader.Read()
				Dim iDate As String = dataReader.getString(2)
				Dim oDate As DateTime = Convert.ToDateTime(iDate)
				txtString = dataReader.getstring(4) & " - " & dataReader.getString(0) & " " & dataReader.getString(1) & " - " & oDate.Month & "/" & oDate.Day & "/" & oDate.Year
				
				Call ShowStatus(eventData, txtString)
				'selected_id = dataReader.getString(3)
				'Call SetCurID(eventData, selected_id)			

				eventData.Form.Fields.GetField("dateofbirth").Value = oDate.Month & "/" & oDate.Day & "/" & oDate.Year
				eventData.Form.Fields.GetField("documentType").Value =  "Neurology Document"
				eventData.Form.Fields.GetField("visitdate").Value =  "8/17/2021"

				index += 1
			End While
			
			If Not (dataReader Is Nothing) Then
				dataReader.Close()
			End If
			conn.Close()			
		Catch ex As Exception
			'	Console.WriteLine(ex.Message)
		End Try	
	End Function
End Module
