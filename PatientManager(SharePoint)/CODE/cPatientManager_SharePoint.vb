'LoadAssembly:System.Web.Extensions.dll
'LoadAssembly:System.Data.dll
'LoadAssembly:System.Data.DataSetExtensions.dll
'LoadAssembly:System.Xml.dll
'LoadAssembly:System.Security.Cryptography.X509Certificates.dll
'LoadAssembly:System.Web.dll
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
Imports System.Security.Cryptography.X509Certificates
Imports System.Web.Hosting
Module Script
	'SECURITY_TOKEN is the Share Point Token Key
	'Const SECURITY_TOKEN As String = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Im5PbzNaRHJPRFhFSzFqS1doWHNsSFJfS1hFZyIsImtpZCI6Im5PbzNaRHJPRFhFSzFqS1doWHNsSFJfS1hFZyJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvc3Rhc2luZGV2LnNoYXJlcG9pbnQuY29tQDFiNDQ3NDdlLTQyM2ItNDE5Ny1iYTI4LTExNmJmYWE2YjU0ZiIsImlzcyI6IjAwMDAwMDAxLTAwMDAtMDAwMC1jMDAwLTAwMDAwMDAwMDAwMEAxYjQ0NzQ3ZS00MjNiLTQxOTctYmEyOC0xMTZiZmFhNmI1NGYiLCJpYXQiOjE2Mjk1MzU3NjYsIm5iZiI6MTYyOTUzNTc2NiwiZXhwIjoxNjI5NjIyNDY2LCJpZGVudGl0eXByb3ZpZGVyIjoiMDAwMDAwMDEtMDAwMC0wMDAwLWMwMDAtMDAwMDAwMDAwMDAwQDFiNDQ3NDdlLTQyM2ItNDE5Ny1iYTI4LTExNmJmYWE2YjU0ZiIsIm5hbWVpZCI6IjI5YzNjNWU5LTY2MTUtNDc1Yy04Nzc2LWE0ZjU2MjNmNjdiZkAxYjQ0NzQ3ZS00MjNiLTQxOTctYmEyOC0xMTZiZmFhNmI1NGYiLCJvaWQiOiIzMzJiODEzMC0wY2QxLTRjMTItOGU3YS1iNzVkOWFkYjZjYjgiLCJzdWIiOiIzMzJiODEzMC0wY2QxLTRjMTItOGU3YS1iNzVkOWFkYjZjYjgiLCJ0cnVzdGVkZm9yZGVsZWdhdGlvbiI6ImZhbHNlIn0.RTWq7JvJmcgACS5dTo0T8NOzsv_T8jddSgc6rEYW5Pevudn--_mVaKvy-rtJZ-YFtZ9BExGUL_ExWxboyePEbPnwbCdA_EH3RapM2eCgtGPpebhwKeIs32btmzY18ImsysGbArDnHAIV06qs174bAGZ0wqh8FuZNy7o10Fs_0XhYlr4lpVjbeFzY6vtJvUqD2O2i2A3leix-zF6ycGYGFfEWZoLx1dmjkQ9aMSIAwJdaIOGOxiuDAi5dz7YJ5Xgs9uPZszaWNlUkFDy8AReM9o40Ty-LPSsRbCyQoXSVOdQV5ZgPXHpYqizehOpUH4aJ74kKk3gHtpk3XLDz_aVO3w"
	Dim SECURITY_TOKEN As String = ""
	Const CLIENT_ID As String = "29c3c5e9-6615-475c-8776-a4f5623f67bf"
	Const CLIENT_SECRET As String = "EHQb9lfxdBX2wLfjP5Fgj0aNk9QQdfQ3v98s3rYjfDk="
	Const SITE_URL As String= "stasindev.sharepoint.com"
	Const TENANT_ID As String = "1b44747e-423b-4197-ba28-116bfaa6b54f"
	Const SITE_NAME As String = "PatientManage"
	Const SITE_LIST_NAME As String = "PatientManage"
	
		
	
	Sub Form_OnLoad(ByVal eventData As MFPEventData)
		SECURITY_TOKEN = GetTokenKey(eventData)
		Call InitializeInterface(eventData)
		
	End Sub
	'Fn : Show the app status
	'Param
	'- strMSG : message content
	'Ret : nothing
	Function ShowStatus(ByVal eventData As MFPEventData, ByVal strMSG As String)
		Dim status As TextField = eventData.Form.Fields.GetField("status")
		status.Value = strMSG	
	End Function
	'Fn : Show Interface field. 
	'Param
	'- bShowFlag : If bShowFlag Is True, show the search Interface, Or If False, show the result Interface
	'Ret : nothing
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
	'Fn : Set a id of the selected patient in a hidden labelfield "curID"
	'Param
	'- id : id of the selected patient
	'Ret : nothing
	Function SetCurID(ByVal eventData As MFPEventData, ByVal id As String)
		eventData.Form.Fields.GetField("curID").Value = id
	End Function
	'Fn : Get a id of the selected patient from a hidden labelfield "curID"
	'Param : nothing
	'Ret : id
	Function GetCurID(ByVal eventData As MFPEventData)
		Dim id As String = eventData.Form.Fields.GetField("curID").Value
		Return id
	End Function
	
	Sub gobackBtn_OnChange(ByVal eventData As MFPEventData)
		eventData.Form.Fields.GetField("searchBtn").Value = False
		'When goback button is clicked, initialize the field value of interface.
		Call InitializeInterface(eventData)
	End Sub
	'Fn : Initialize the field values of interface the empty string and show the Search Mode
	'Param : nothing
	'Ret : nothing
	Function InitializeInterface(ByVal eventData As MFPEventData)
		eventData.Form.Fields.GetField("patientID").Value = ""
		eventData.Form.Fields.GetField("nom").Value = ""
		eventData.Form.Fields.GetField("prenom").Value ="" 
		eventData.Form.Fields.GetField("patientList").Value =""
		eventData.Form.Fields.GetField("dateofbirth").Value = ""
		eventData.Form.Fields.GetField("documentType").Value = ""
		eventData.Form.Fields.GetField("visitdate").Value = "" 
		'set a selected id "-1"
		Call SetCurID(eventData, "-1")
		'show the search mode
		ShowInterface(eventData, True)
		'if Parameter is "", "", "", show the info of total patients
		Dim recordCount As Integer = GetPatientList(eventData, "", "", "")
		'if recordcount is greater than 0, successed
		'or is less than 0, that is -1, failed
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
		'Populate patient informations which has the following search conditions into the patient list.
		Dim recordCount As Integer = GetPatientList(eventData, patientID, nom, prenom)
		If recordCount > 1 Then 'if searching result is greater than 1, show the searching count in status.
			Call ShowStatus(eventData, recordCount & " Records found. Please select the Patient List.")
		Else If recordCount = 1 Then'if searching result is 1, change the screen the output mode and show the information of patient.
			Call ShowInterface(eventData, False)
			Dim curID As String = GetCurID(eventData)
			If curID > 0 Then
				Call ShowStatus(eventData, "1 Record found. Please select the Patient List.")
				Call GetPaitentInfoByID(eventData, curID)
			End If
			
		Else 'if searching is failed, show "Not found record" message on status
			Call ShowStatus(eventData, "Not found record.")
		End If
	End Sub
	Sub patientList_OnChange(ByVal eventData As MFPEventData) 'When patient List selected, event occurs
		'Dim patient
		Dim patientList As ListField = eventData.Form.Fields.GetField("patientList")
		Dim id As String = patientList.GetSelectedItem().Value
		'change the screen mode the result output mode
		Call ShowInterface(eventData, False)
		'get the patient info which the patient id is id, and show the result in interface.
		Call GetPaitentInfoByID(eventData, id)
	End Sub
	'Fn : By using a Share Point API, get the patient list
	'API URI : https://stasindev.sharepoint.com/sites/Patient/_api/Web/Lists/getbytitle('patient')/Items
	'Param 
	'-- patiendID : searching patiend id
	'-- nom : searching the last name of patient
	'-- prenom : the first name of patient
	'Ret : searching result count. if -1 returns, searching is failed
	Function GetPatientList(ByVal eventData As MFPEventData, ByVal patientID As String, ByVal nom As String, ByVal prenom As String)
		
		Dim url As String = "https://" & SITE_URL & "/sites/" & SITE_NAME & "/_api/Web/Lists/getbytitle('" & SITE_LIST_NAME & "')/Items"
		
		Dim address As Uri = New Uri(url)
	
		Dim request As HttpWebRequest = DirectCast(WebRequest.Create(address), HttpWebRequest)
		'set request mode "get" mode
		request.Method = "GET"
		'set request content-type "application/json;odata=verbose"
		request.ContentType = "application/json;odata=verbose"
		'set request accept "application/json;odata=verbose"
		request.Accept = "application/json;odata=verbose"
		'add token key to "Authorization" field of request header
		request.Headers.Add("Authorization", "Bearer " & SECURITY_TOKEN)
		
		Dim response As HttpWebResponse = DirectCast(request.GetResponse(), HttpWebResponse)
		
		Dim reader As StreamReader = New StreamReader(response.GetResponseStream())
		
		Dim json As String = reader.ReadToEnd()
		'response string is following
		' {"d": {"results":[{"nom" : "smith", "prenom":"anna"}, {"nom":"alex","prenom":"jhon"}...]...}}
		

		
		If Not response Is Nothing Then response.Close()
		'response sting has unnecessary string so that cut the string.
		'output result : [{"nom" : "smith", "prenom":"anna"}, {"nom":"alex","prenom":"jhon"}...]
		'reason : following engine does not parse the string.
		json = Mid(json, 17, json.Length - 18)
		
		Dim serializer As JavaScriptSerializer = New JavaScriptSerializer()
		'convert a string to arraylist
		Dim results As  System.Collections.ArrayList = serializer.Deserialize(Of System.Collections.ArrayList)(json)
		Dim results1 As  System.Collections.Generic.Dictionary(Of String, Object) 
		Dim table As DataTable = New DataTable("Patient")
		'For using sql instruction, convert arraylist to data table
		'set the field of tablea
		table.Columns.Add(New DataColumn("ID", GetType(String)))
		table.Columns.Add(New DataColumn("patientID", GetType(String)))		' "patientID" means "Title"
		table.Columns.Add(New DataColumn("nom", GetType(String)))
		table.Columns.Add(New DataColumn("prenom", GetType(String)))
		table.Columns.Add(New DataColumn("dateofbirth", GetType(String)))
		'populate the arraylist data in data table
		For index As Integer = 0 To results.Count - 1
			results1 = results.Item(index)	
			
			Dim iDate As String = results1.Item("dateofbirth")
			Dim oDate As DateTime = Convert.ToDateTime(iDate)
			table.Rows.Add(results1.Item("ID"), results1.Item("Title"), results1.Item("nom"), results1.Item("prenom"), oDate.Date)
						
		Next
		
		Dim rows() As System.Data.DataRow
		Dim row As System.Data.DataRow
		Dim  sqlString As String = ""
		'making sql
		If patientID <> "" Then
			sqlString = sqlString & "AND patientID = '" & patientID & "'"
		End If
		If nom <> "" Then
			sqlString = sqlString & "AND nom = '" & nom & "'"
		End If
		If prenom <> "" Then
			sqlString = sqlString & "AND prenom = '" & prenom & "'"
		End If
		'get the result which sql instruction runs
		rows = table.Select("1 = 1 " & sqlString)
		Dim documentList As ListField = eventData.Form.Fields.GetField("patientList")

		documentList.FindMode = False
		documentList.Items.Clear()
		documentList.MaxSearchResults = rows.Length + 1
		Dim listItem As ListItem 
		'populate the searching data on patient list
		For Each row In rows
			If rows.Length = 1 Then
				Call SetCurID(eventData, row("ID"))
			Else
				Call SetCurID(eventData, "-1")
			End If
			Dim iDate As String = row("dateofbirth")
			Dim oDate As DateTime = Convert.ToDateTime(iDate)
			
			listItem = New ListItem(row("patientID") & " - " & row("nom") & " " & row("prenom") & " - " & oDate.Month & "/" & oDate.Day & "/" & oDate.Year, row("ID"))
			documentList.Items.Add(listItem)
		Next
		
		
		'return searching result count
		Return rows.Length
		'GetPatientList = response.StatusCode
	End Function
	
	Function GetPaitentInfoByID(ByVal eventData As MFPEventData, ByVal id As String)
		Dim url As String = "https://" & SITE_URL & "/sites/" & SITE_NAME & "/_api/Web/Lists/getbytitle('" & SITE_LIST_NAME & "')/Items(" & id & ")"
		
		Dim address As Uri = New Uri(url)
	
		Dim request As HttpWebRequest = DirectCast(WebRequest.Create(address), HttpWebRequest)
	
		request.Method = "GET"
		request.ContentType = "application/json;odata=verbose"
		request.Accept = "application/json;odata=verbose"
		request.Headers.Add("Authorization", "Bearer " & SECURITY_TOKEN)
		
		Dim response As HttpWebResponse = DirectCast(request.GetResponse(), HttpWebResponse)
		
		Dim reader As StreamReader = New StreamReader(response.GetResponseStream())
		
		Dim json As String = reader.ReadToEnd()
		If Not response Is Nothing Then response.Close()
		
		json = Mid(json, 6, json.Length - 6)
		json = "[" & json & "]"
		
		Dim serializer As JavaScriptSerializer = New JavaScriptSerializer()
		
		Dim results As  System.Collections.ArrayList = serializer.Deserialize(Of System.Collections.ArrayList)(json)
		
		Dim results1 As  System.Collections.Generic.Dictionary(Of String, Object) 
		
		For index As Integer = 0 To results.Count - 1
			results1 = results.Item(index)	
			
			Dim iDate As String = results1.Item("dateofbirth")
			Dim oDate As DateTime = Convert.ToDateTime(iDate)
			Call ShowStatus(eventData, results1.Item("Title") & " - " & results1.Item("nom") & " " & results1.Item("prenom") & " - " & oDate.Month & "/" & oDate.Day & "/" & oDate.Year)
			
			Call SetCurID(eventData, results1.Item("ID"))			

			eventData.Form.Fields.GetField("dateofbirth").Value = oDate.Month & "/" & oDate.Day & "/" & oDate.Year
			eventData.Form.Fields.GetField("documentType").Value =  "Neurology Document"
			eventData.Form.Fields.GetField("visitdate").Value =  "8/17/2021"
		Next		
		
		GetPaitentInfoByID = response.StatusCode
	End Function
	
	Function GetTokenKey(ByVal eventData As MFPEventData) As String
		
		Dim url As String = "https://accounts.accesscontrol.windows.net/" & TENANT_ID & "/tokens/OAuth/2/"
		
		Dim address As Uri = New Uri(url)
	
		Dim request As HttpWebRequest = DirectCast(WebRequest.Create(address), HttpWebRequest)
	
		request.Method = "POST"
		'request.ContentType = "application/json;odata=verbose"
		request.Accept = "application/json;odata=verbose"
		
		
		
		'Dim json_data As String = "{""grant_type"":""client_credentials"", "
		'json_data = json_data & """client_id"": """ & CLIENT_ID & "@" &  TENANT_ID & ""","
		'json_data = json_data & """client_secret"": """ & CLIENT_SECRET & ""","
		'json_data = json_data & """resource"": ""00000003-0000-0ff1-ce00-000000000000/" & SITE_URL & "@" & TENANT_ID & """}"
		Dim json_data As String = "grant_type=client_credentials"
		json_data = json_data & "&client_id=" & CLIENT_ID & "@" &  TENANT_ID
		json_data = json_data & "&client_secret=" & CLIENT_SECRET
		json_data = json_data & "&resource=00000003-0000-0ff1-ce00-000000000000/" & SITE_URL & "@" & TENANT_ID
		
		'Call ShowStatus(eventData, json_data)
		
		Dim json_bytes() As Byte = System.Text.Encoding.ASCII.GetBytes(json_data)
		request.ContentLength = json_bytes.Length
		
		Dim stream As IO.Stream = request.GetRequestStream
		stream.Write(json_bytes, 0, json_bytes.Length)

		Dim response As HttpWebResponse = request.GetResponse
		Dim reader As StreamReader = New StreamReader(response.GetResponseStream())
		
		Dim json As String = reader.ReadToEnd()
		reader.Close()
		If Not response Is Nothing Then response.Close()
		
		json = "[" & json & "]"
		'Call ShowStatus(eventData, json)
		Dim serializer As JavaScriptSerializer = New JavaScriptSerializer()
		
		Dim results As  System.Collections.ArrayList = serializer.Deserialize(Of System.Collections.ArrayList)(json)
		
		Dim results1 As  System.Collections.Generic.Dictionary(Of String, Object) 
		
		Dim tokenKey As String
		For index As Integer = 0 To results.Count - 1
			results1 = results.Item(index)	
			tokenKey = results1("access_token")
			'Dim iDate As String = results1.Item("dateofbirth")
			
		Next
		
		
		Return tokenKey
	End Function
End Module
