Imports System.Text.RegularExpressions
Imports Newtonsoft.Json.Linq
Public Class NFunctions


	''' <summary>
	''' Checks if a string is valid email address.
	''' </summary>
	''' <param name="email_">String to check</param>
	''' <returns>True or False</returns>
	Public Shared Function IsEmail(email_ As String) As Boolean
		Dim isValid As Boolean = True
		If Not Regex.IsMatch(email_,
			"^([\w-]+\.)*?[\w-]+@[\w-]+\.([\w-]+\.)*?[\w]+$") Then
			isValid = False
		End If
		Return isValid
	End Function

#Region "J"
	''' <summary>
	''' Turns database field's string to JSON. Just bind the list it returns to the ListBox or ComboBox directly, or loop through to populate the ListBox, ComboBox or TextBox (using D.DatabaseToControlJSON()).
	''' </summary>
	''' <param name="val_">Database field's string</param>
	''' <returns>List you can bind to ListBox or ComboBox</returns>
	Public Shared Function DatabaseToListJSON(val_ As String) As List(Of String)
		Dim l As New List(Of String)
		Try
			l.Clear()
		Catch ex As Exception
		End Try
		If val_.Trim.Length < 1 Then
			Return l
			Exit Function
		End If

		Dim str_ As String = ""
		If Mid(val_.Trim, 1, 1) <> "{" Then
			str_ = "{" & val_.Trim & "}"
		Else
			str_ = val_.Trim
		End If

		Dim w As JObject = JObject.Parse(str_)

		Dim s() = w.PropertyValues.ToArray
		For i As Integer = 0 To s.Count - 1
			l.Add(s(i).ToString)
		Next
		Return l
	End Function
	''' <summary>
	''' DEPRACATED. Checks if string contains property with its value.
	''' </summary>
	''' <param name="val_">JSON string</param>
	''' <param name="property_">Property name</param>
	''' <param name="value_">The value</param>
	''' <returns>True if property_ with value_ is found, False otherwise</returns>
	Public Shared Function DatabaseContainsJSON(val_ As String, property_ As String, value_ As String) As Boolean
		Dim str_ As String = ""
		If Mid(val_.Trim, 1, 1) <> "{" Then
			str_ = "{" & val_.Trim & "}"
		Else
			str_ = val_.Trim
		End If
		Dim w As JObject = JObject.Parse(str_)
		Return w.TryGetValue(property_, value_)
	End Function
	''' <summary>
	''' Turns string to JSON. Before using anything JSON, you might want to use what is returned from this.
	''' </summary>
	''' <param name="val_">String to turn to JSON</param>
	''' <returns>JSON</returns>
	Public Shared Function ToJSON(val_) As String
		Dim str_ As String = ""
		If Mid(val_.Trim, 1, 1) <> "{" Then
			str_ = "{" & val_.Trim & "}"
		Else
			str_ = val_.Trim
		End If
		Return str_
	End Function

	''' <summary>
	''' Gets value from property.
	''' </summary>
	''' <param name="val_">Value sought.</param>
	''' <param name="property_">Property to search for.</param>
	''' <returns></returns>
	Public Shared Function GetValueJSON(val_ As String, property_ As String) As String
		Dim x As JObject = JObject.Parse(ToJSON(val_))
		Return x.GetValue(property_).ToString
	End Function

#End Region

#Region "Map"
	Public Function MapProvidersList() As List(Of String)
		Dim l_ As New List(Of String)

		l_.Add("Bing Hybrid")
		l_.Add("Bing")

		l_.Add("ArcGIS StreetMap World 2D")
		l_.Add("ArcGIS World Street")
		l_.Add("Google China Hybrid")
		l_.Add("Google China")
		l_.Add("Google Hybrid")
		l_.Add("Google")
		l_.Add("OpenCycle Landscape")
		l_.Add("OpenCycle Transport")
		l_.Add("OpenCycle")
		l_.Add("WikiMapia")
		l_.Add("Yandex Hybrid")
		l_.Add("Yandex")

		''not sure they're working
		'l_.Add("Near Hybrid")
		'l_.Add("Near")
		'l_.Add("Near Satellite")
		''l_.Add("OpenStreetMap Hybrid)")
		''l_.Add("OpenStreetMap")
		'l_.Add("CloudMade")
		'l_.Add("MapBender")
		'l_.Add("Yahoo Hybrid")
		'l_.Add("Yahoo")
		'l_.Add("Google Korea Hybrid")
		'l_.Add("Google Korea")
		'l_.Add("OpenSeaMap Hybrid")
		'l_.Add("OpenStreet4U")
		'l_.Add("OpenStreet MapQuest Hybrid")
		'l_.Add("OpenStreet MapQuest")
		'l_.Add("OpenStreet MapQuest Sattelite")
		'l_.Add("Ovi Hybrid")
		'l_.Add("Ovi")
		'l_.Add("Ovi Sattelite")
		'l_.Add("Ovi Terrain")
		'l_.Add("Yahoo Satellite")

		''maybe not directly relevant
		'l_.Add("Yandex Satellite")
		'l_.Add("ArcGIS World Physical")
		'l_.Add("ArcGIS World Shaded Relief")
		'l_.Add("ArcGIS World Terrain Base")
		'l_.Add("ArcGIS World Topo")
		'l_.Add("Bing Satellite")
		'l_.Add("Google China Satellite")
		'l_.Add("Google China Terrain")
		'l_.Add("Google Korea Satellite")
		'l_.Add("Google Satellite")
		'l_.Add("Google Terrain")

		''maybe too specific...
		'l_.Add("ArcGIS Topo US 2D")
		'l_.Add("Czech Geographic")
		'l_.Add("Czech History")
		'l_.Add("Czech History (Old)")
		'l_.Add("Czech Hybrid")
		'l_.Add("Czech Hybrid (Old)")
		'l_.Add("Czech")
		'l_.Add("Czech (Old)")
		'l_.Add("Czech Satellite")
		'l_.Add("Czech Satellite (Old)")
		'l_.Add("Czech Turist")
		'l_.Add("Czech Turist (Old)")
		'l_.Add("Czech Turist (Winter)")
		'l_.Add("Latvia")
		'l_.Add("Lithuania 3d")
		'l_.Add("Lithuania Hybrid")
		'l_.Add("Lithuania Hybrid (Old)")
		'l_.Add("Lithuania")
		'l_.Add("Lithuania OrtoFoto")
		'l_.Add("Lithuania OrtoFoto (Old)")
		'l_.Add("Lithuania Relief")
		'l_.Add("Lithuania TOP 50")
		'l_.Add("Spain")
		'l_.Add("Sweden")
		'l_.Add("Turkey")
		Return l_
	End Function

#End Region
End Class
