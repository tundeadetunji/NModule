Imports NModule.Methods
Public Class SJ
	Public Shared Function ConcatJSON(previous_json_string As String, new_json_string As String) As String
		Return previous_json_string & ", " & new_json_string
	End Function
	Public Shared Function ToDateJSON(date_)
		Dim d_ As DateTime = DateTime.Parse(date_)
		Return d_.Year & "-" & LeadingZero(d_.Month) & "-" & LeadingZero(d_.Day)
	End Function
	Public Shared Function DictionaryToDatabaseJSON(l_p As List(Of Object), l_v As List(Of Object)) As String
		Dim val_ As String = ""
		For i As Integer = 0 To l_p.Count - 1
			val_ &= "'" & l_p.Item(i) & "':'" & l_v.Item(i)
			If i < l_p.Count - 1 Then
				val_ &= ", "
			End If
		Next
		Return val_
	End Function
	Public Shared Function DictionaryToDatabaseJSON(l As List(Of Object)) As String
		Dim val_ As String = ""
		Dim ltemp As New List(Of Object)
		Try
			ltemp.Clear()
		Catch ex As Exception
		End Try
		ltemp = l
		If ltemp.Count > 0 Then
			With ltemp
				For k As Integer = 0 To .Count - 1 Step 2
					val_ &= "'" & .Item(k) & "':'" & .Item(k + 1) & "'"
					'
					If k <> .Count - 2 Then
						val_ &= ", "
					End If
				Next
			End With
		End If
		Return val_
	End Function

End Class
