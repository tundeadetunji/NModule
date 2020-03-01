Imports NModule.NFunctions
Imports Microsoft.Win32
Imports System.Security.AccessControl
Imports System.IO
Imports System.Windows.Forms
Public Class Methods

	Private suffx
	Private prefx
	Private countr
	Private TextHasSpace As Boolean
	Private strSource


	''' <summary>
	''' Changes text to title case. Called from _TextChanged. Multiline is not supported yet.
	''' </summary>
	''' <param name="strSource"></param>
	Public Shared Sub ToTitleCase(ByRef strSource As Control)
		Dim g As New Methods
		Try
			g.ConvertTextToTitleCase(strSource)
		Catch
		End Try
	End Sub

	''' <summary>
	''' Use ToTitleCase instead.
	''' </summary>
	''' <param name="strSource"></param>
	Public Sub ConvertTextToTitleCase(ByRef strSource As Control)

		'convert text to title case
		'called from TextChanged
		If TypeOf strSource Is TextBox Then
			Dim t As TextBox = strSource
			If t.Multiline = True Then Exit Sub
		End If

		On Error Resume Next
		If Len(strSource.Text) = 0 Then
			suffx = ""
			prefx = ""
			countr = Nothing
			Exit Sub
		End If

		If Len(strSource.Text) = 1 Then strSource.Text = UCase(strSource.Text) : System.Windows.Forms.SendKeys.Send("{End}")

		If Mid(strSource.Text, Len(strSource.Text), 1) = " " Or Mid(strSource.Text, Len(strSource.Text), 1) = "." Or Mid(strSource.Text, Len(strSource.Text), 1) = ChrW(13) Then
			TextHasSpace = True
			countr = Len(strSource.Text) + 1
			prefx = strSource.Text
		End If

		If prefx <> "" And Len(strSource.Text) = Val(countr) Then
			suffx = UCase(Mid(strSource.Text, Len(strSource.Text), 1))
			strSource.Text = prefx & suffx
			System.Windows.Forms.SendKeys.Send("{End}")
			'clear counters
			suffx = ""
			prefx = ""
			countr = Nothing
		End If
	End Sub
	Public Sub SetTextChange(c As Control, placeholder As String, placeholderLabel As Label, Mode As String, IsUsername As Boolean)

		If Trim(c.Text) = "" Then
			c.Text = placeholder
			placeholderLabel.Visible = False
		ElseIf Trim(c.Text) <> "" And c.Text <> placeholder Then
			placeholderLabel.Visible = True
		End If

		Select Case Mode.ToLower
			Case "email"
				If TypeOf c Is TextBox Then
					Dim t As TextBox = c
					If Trim(t.Text).Length > 0 And t.Multiline = False And IsUsername = False Then
						ConvertTextToLowerCase(c)
					End If
				ElseIf TypeOf c Is ComboBox Then
					If Trim(c.Text).Length > 0 Then ConvertTextToLowerCase(c)
				End If
			Case Else
				If TypeOf c Is TextBox Then
					Dim t As TextBox = c
					If Trim(t.Text).Length > 0 And t.Multiline = False And IsUsername = False Then
						ConvertTextToTitleCase(c)
					End If
				ElseIf TypeOf c Is ComboBox Then
					If Trim(c.Text).Length > 0 Then ConvertTextToTitleCase(c)
				End If
		End Select
	End Sub
	''' <summary>
	''' Sets control's placeholder.
	''' </summary>
	''' <param name="c"></param>
	''' <param name="placeholder"></param>
	Public Sub SetUsernameTextChange(c As Control, placeholder As String)
		If TypeOf (c) Is TextBox Or TypeOf (c) Is ComboBox Then
			If c.Text = "" Then c.Text = placeholder
		End If
	End Sub
	''' <summary>
	''' Sets password control's placeholder.
	''' </summary>
	''' <param name="c"></param>
	''' <param name="placeholder"></param>
	Public Sub SetPasswordChange(c As TextBox, Optional placeholder As String = "Password")
		If c.Text = "" Then c.Text = placeholder : c.PasswordChar = Nothing : Exit Sub
		If c.Text.ToLower <> placeholder Then c.PasswordChar = "*" : Exit Sub
	End Sub
	''' <summary>
	''' Sets username's control's placeholder.
	''' </summary>
	''' <param name="c"></param>
	''' <param name="placeholder"></param>
	Public Sub UsernameTextChange(c As Control, Optional placeholder As String = "Username")
		If TypeOf (c) Is TextBox Or TypeOf (c) Is ComboBox Then
			If c.Text = "" Then c.Text = placeholder
		End If
	End Sub
	''' <summary>
	''' Sets password control's placeholder. Same as SetPasswordChange.
	''' </summary>
	''' <param name="c"></param>
	''' <param name="placeholder"></param>
	Public Sub PasswordTextChange(c As TextBox, Optional placeholder As String = "Password")
		If c.Text = "" Then c.Text = placeholder : c.PasswordChar = Nothing : Exit Sub
		If c.Text.ToLower <> placeholder Then c.PasswordChar = "*" : Exit Sub
	End Sub

	''' <summary>
	''' Allows numbers and % only. Called from _KeyPress.
	''' </summary>
	''' <param name="e"></param>
	''' <param name="controlText"></param>
	Public Shared Sub AllowPercentage(ByRef e As System.Windows.Forms.KeyPressEventArgs, controlText As String)
		'        On Error Resume Next
		'allow numbers(48-57), backspace(8), enter(13), full stop(46), delete(127), percent(37) and minus(45) only
		If (e.KeyChar >= ChrW(48) And e.KeyChar <= ChrW(57)) Or e.KeyChar = ChrW(8) Or e.KeyChar = ChrW(13) Or e.KeyChar = ChrW(127) Or e.KeyChar = ChrW(46) Or e.KeyChar = ChrW(45) Or e.KeyChar = ChrW(37) Then
			Exit Sub
		Else
			e.KeyChar = ChrW(0)
		End If
	End Sub

	''' <summary>
	''' Allows positive and negative numbers and period only. Called from _KeyPress. Good for money.
	''' </summary>
	''' <param name="e"></param>
	Public Shared Sub AllowNumberOnly(ByRef e As System.Windows.Forms.KeyPressEventArgs)
		'        On Error Resume Next
		'allow numbers(48-57), backspace(8), enter(13), full stop(46), delete(127) and minus(45) only
		If (e.KeyChar >= ChrW(48) And e.KeyChar <= ChrW(57)) Or e.KeyChar = ChrW(8) Or e.KeyChar = ChrW(13) Or e.KeyChar = ChrW(127) Or e.KeyChar = ChrW(46) Or e.KeyChar = ChrW(45) Then
			Exit Sub
		Else
			e.KeyChar = ChrW(0)
		End If
	End Sub
	''' <summary>
	''' Discontinued. Do not use it.
	''' </summary>
	''' <param name="e"></param>
	Public Sub AllowPIN(ByRef e As System.Windows.Forms.KeyPressEventArgs)
		'        On Error Resume Next
		'allow numbers (48-57), backspace(8), enter(13), delete(127) only
		If (e.KeyChar >= ChrW(48) And e.KeyChar <= ChrW(57)) Or e.KeyChar = ChrW(8) Or e.KeyChar = ChrW(13) Or e.KeyChar = ChrW(127) Then
			Exit Sub
		Else
			e.KeyChar = ChrW(0)
		End If
	End Sub

	''' <summary>
	''' Allows numbers only. Called from _KeyPress. Good for PIN.
	''' </summary>
	''' <param name="e"></param>

	Public Shared Sub AllowDigitOnly(ByRef e As System.Windows.Forms.KeyPressEventArgs)
		'        On Error Resume Next
		'allow numbers(48-57), backspace(8), enter(13), full stop(46), delete(127) only
		If (e.KeyChar >= ChrW(48) And e.KeyChar <= ChrW(57)) Or e.KeyChar = ChrW(8) Or e.KeyChar = ChrW(13) Or e.KeyChar = ChrW(127) Or e.KeyChar = ChrW(46) Then
			Exit Sub
		Else
			e.KeyChar = ChrW(0)
		End If
	End Sub

	''' <summary>
	''' Allows nothing.
	''' </summary>
	''' <param name="e"></param>
	Public Shared Sub AllowNothing(ByRef e As System.Windows.Forms.KeyPressEventArgs)
		e.KeyChar = ChrW(0)
	End Sub
	''' <summary>
	''' Changes text to uppercase. Called from _TextChanged.
	''' </summary>
	''' <param name="strSource"></param>
	Public Shared Sub ToUpperCase(ByRef strSource As Control)
		'		On Error Resume Next
		'called from TextChanged
		If Len(strSource.Text) = 0 Then Exit Sub
		Try
			strSource.Text = UCase(strSource.Text)
		Catch
		End Try
	End Sub
	''' <summary>
	''' Changes text to lower case. Called from _TextChanged. Useful for email addresses and websites.
	''' </summary>
	''' <param name="strSource"></param>
	Public Shared Sub ToLowerCase(ByRef strSource As Control)
		Dim g As New Methods
		Try
			g.ConvertTextToLowerCase(strSource)
		Catch
		End Try
	End Sub

	''' <summary>
	''' Use ToLowerCase instead.
	''' </summary>
	''' <param name="strSource"></param>
	Public Sub ConvertTextToLowerCase(ByRef strSource As Control)
		On Error Resume Next
		'convert text to lower case (useful for email and websites addresses)
		'called from TextChanged
		If Len(strSource.Text) = 0 Then Exit Sub

		strSource.Text = LCase(strSource.Text)

	End Sub
	''' <summary>
	''' Allows nothing. Same as AllowNothing.
	''' </summary>
	''' <param name="sender"></param>
	''' <param name="e"></param>
	Public Sub NoKey(sender As Object, e As KeyPressEventArgs)
		AllowNothing(e)
	End Sub

	''' <summary>
	''' Formats number to 2 decimal places.
	''' </summary>
	''' <param name="val_"></param>
	''' <returns></returns>
	Public Shared Function ToCurrency(val_)
		Try
			'			If val_.ToString.Length > 0 Then
			Return RoundNumber(val_)
			'			End If
		Catch
		End Try
	End Function


#Region "Lists"
	''' <summary>
	''' Basic ListBox operations.
	''' </summary>
	Public Shared Sub IncludeItem(left_list As ListBox, right_list As ListBox)
		With left_list
			If .Items.Count < 1 Or .SelectedIndex < 0 Or .SelectedIndex > .Items.Count - 1 Then Exit Sub
			right_list.Items.Add(.SelectedItem.ToString)
			.Items.RemoveAt(.SelectedIndex)
		End With
	End Sub

	''' <summary>
	''' Basic ListBox operations.
	''' </summary>
	Public Shared Sub IncludeAllItems(left_list As ListBox, right_list As ListBox)
		With left_list
			If .Items.Count < 1 Then Exit Sub
			.SelectedIndex = 0
			For ia As Integer = 0 To .Items.Count - 1
				right_list.Items.Add(.Items.Item(ia).ToString)
				'				.Items.RemoveAt(ia)
			Next
			.Items.Clear()
		End With
	End Sub

	''' <summary>
	''' Basic ListBox operations.
	''' </summary>
	Public Shared Sub ExcludeItem(left_list As ListBox, right_list As ListBox)
		With right_list
			If .Items.Count < 1 Or .SelectedIndex < 0 Or .SelectedIndex > .Items.Count - 1 Then Exit Sub
			left_list.Items.Add(.SelectedItem.ToString)
			.Items.RemoveAt(.SelectedIndex)
		End With
	End Sub

	''' <summary>
	''' Basic ListBox operations.
	''' </summary>
	Public Shared Sub ExcludeAllItems(left_list As ListBox, right_list As ListBox)
		With right_list
			If .Items.Count < 1 Then Exit Sub
			.SelectedIndex = 0
			For r As Integer = 0 To .Items.Count - 1
				left_list.Items.Add(.Items.Item(r).ToString)
				'				.Items.RemoveAt(ia)
			Next
			.Items.Clear()
		End With
	End Sub
	''' <summary>
	''' Basic ListBox operations.
	''' </summary>
	Public Shared Sub AddItem_Multiline(list_ As ListBox, item_ As TextBox, Optional message_if_already_exists As String = "", Optional allow_duplicate As Boolean = False, Optional clear_after_add As Boolean = True)
		Dim left_list As ListBox = list_

		If item_.Text.Trim.Length < 1 Then Exit Sub

		If left_list.Items.Contains(item_.Text.Trim) Then
			If allow_duplicate Then
				If message_if_already_exists.Length > 0 Then
					If MsgBox(message_if_already_exists, MsgBoxStyle.YesNo, "") = MsgBoxResult.No Then
						Exit Sub
					End If
				Else
				End If
			Else
				Exit Sub
			End If
		End If

		Dim x As String = item_.Text.Trim.Replace(vbCr, "<New Line>")
		Dim y As String = x.Replace(vbCrLf, "<New Line>")
		Dim item__ As String = y

		With left_list
			.Items.Add(item__)
			item__ = ""
			x = "" : y = ""          ''
			If clear_after_add = True Then item_.Text = ""
			.SelectedIndex = .Items.Count - 1
		End With
	End Sub
	''' <summary>
	''' Basic ListBox operations.
	''' </summary>
	Public Shared Function AddItem(list_ As ListBox, item_ As TextBox, Optional message_if_already_exists As String = "", Optional allow_duplicate As Boolean = False, Optional clear_after_add As Boolean = True) As Boolean
		Dim left_list As ListBox = list_
		If item_.Text.Trim.Length > 0 Then
			If left_list.Items.Contains(item_.Text.Trim) Then
				If allow_duplicate Then
					Return AddTheItem(list_, item_, clear_after_add)
					'warn user
					'If message_if_already_exists.Length > 0 Then
					'	If MsgBox(message_if_already_exists, MsgBoxStyle.YesNo, "") = MsgBoxResult.No Then
					'		Exit Sub
					'	End If
					'Else
					'End If
				Else
					Return False
				End If
			Else
				Return AddTheItem(list_, item_, clear_after_add)
			End If
		End If
	End Function
	''' <summary>
	''' Basic ListBox operations.
	''' </summary>
	Private Shared Function AddTheItem(list_ As ListBox, item_ As TextBox, Optional clear_after_add As Boolean = True) As Boolean
		Dim item__ As String = item_.Text.Trim
		Dim left_list As ListBox = list_

		With left_list
			.Items.Add(item__)
			item__ = ""
			If clear_after_add = True Then item_.Text = ""
			.SelectedIndex = .Items.Count - 1
		End With
		Return True
	End Function
	''' <summary>
	''' Basic ListBox operations.
	''' </summary>
	''' <example>TheTextBox.Text = RetrieveItem(TheListBox)</example>
	Public Shared Function RetrieveItem(list_ As ListBox, Optional item_ As TextBox = Nothing)
		Dim LeftList As ListBox = list_
		With LeftList
			If .Items.Count = 0 Or .SelectedIndex < 0 Then
				Return ""
			Else
				Try
					If item_ IsNot Nothing Then
						item_.ReadOnly = True
						item_.Text = LeftList.SelectedItem.ToString.Replace("<New Line>", vbCrLf)
					End If
				Catch ex As Exception
				End Try
				Return LeftList.SelectedItem.ToString.Replace("<New Line>", vbCrLf)
			End If
		End With


		'		RetrieveItem(TheListBox, TheTextBox)
		'		Or
		'		TheTextBox.Text = RetrieveItem(TheListBox)

	End Function

	''' <summary>
	''' Basic ListBox operations.
	''' </summary>
	Public Shared Sub MoveItemUp(list_ As ListBox)
		Dim LeftList As ListBox = list_
		With LeftList
			If .Items.Count = 0 Or .SelectedIndex <= 0 Then
				Exit Sub
			End If
			.Hide()
		End With

		Dim selected_index As Integer
		Dim before_selected_index As Integer

		Dim selected_item
		Dim before_selected_item

		Dim right_list As New List(Of String)
		Try
			right_list.Clear()
		Catch ex As Exception
		End Try

		With LeftList
			before_selected_item = .Items.Item(.SelectedIndex - 1)
			selected_item = .Items.Item(.SelectedIndex)
			before_selected_index = .SelectedIndex - 1
			selected_index = .SelectedIndex
			.SelectedIndex = 0
			With .Items
				For i As Integer = 0 To .Count - 1
					right_list.Add(LeftList.Items.Item(LeftList.SelectedIndex))
					Try
						LeftList.SelectedIndex += 1
					Catch
					End Try
				Next
				.Clear()
			End With
			With right_list
				For j As Integer = 0 To .Count - 1
					If j = before_selected_index Then
						LeftList.Items.Add(selected_item)
					ElseIf j = selected_index Then
						LeftList.Items.Add(before_selected_item)
					Else
						LeftList.Items.Add(right_list(j))
					End If
				Next
			End With

			.SelectedIndex = before_selected_index
		End With

		Try
			right_list.Clear()
		Catch ex As Exception
		End Try

		LeftList.Show()
	End Sub
	''' <summary>
	''' Basic ListBox operations.
	''' </summary>
	''' <example>Me.Text = ShowSelected(TheListBox, TheLabel)</example>
	Public Shared Function ShowSelected(list_ As ListBox, Optional label_ As Control = Nothing) As String
		If list_.Items.Count = 0 Or list_.SelectedIndex < 0 Then
			If label_ IsNot Nothing Then
				label_.Text = ""
			End If
			Return ""
		Else
			If label_ IsNot Nothing Then
				Try
					label_.Text = (list_.SelectedIndex + 1).ToString
				Catch ex As Exception
				End Try
			End If
			Return (list_.SelectedIndex + 1).ToString
		End If



		'		Me.Text = ShowSelected(TheListBox)
		'		Or, for both
		'		Me.Text = ShowSelected(TheListBox, TheLabel)
		'		Or, for label alone
		'		ShowSelected(TheListBox, TheLabel)


	End Function

	''' <summary>
	''' Basic ListBox operations.
	''' </summary>
	Public Shared Sub RemoveItem(list_ As ListBox)
		Dim selected_ As Integer = list_.SelectedIndex
		If list_.Items.Count > 0 And list_.SelectedIndex >= 0 Then
			list_.Items.RemoveAt(list_.SelectedIndex)
			Try
				list_.SelectedIndex = selected_
			Catch ex As Exception
				Try
					list_.SelectedIndex = selected_ - 1
				Catch
					Try
						list_.SelectedIndex = list_.Items.Count - 1
					Catch
					End Try
				End Try
			End Try
		End If

	End Sub
#End Region

#Region "Reusable"

	''' <summary>
	''' Adds a random number to the list, keeping all the list's items unique.
	''' </summary>
	''' <param name="random_inclusive_min">Where to start (used by random)</param>
	''' <param name="random_exclusive_max">Where to stop (used by random)</param>
	''' <param name="already_">The list that to add to</param>
	''' <returns>The list with the new number</returns>

	Public Function RandomList(random_inclusive_min As Integer, random_exclusive_max As Integer, already_ As List(Of Integer)) As List(Of Integer)
		Dim r_val
2:
		r_val = Random_(random_inclusive_min, random_exclusive_max)
		If already_.Count > 0 And already_.Contains(r_val) Then
			GoTo 2
		Else
			already_.Add(r_val)
			Return already_
		End If
	End Function

	''' <summary>
	''' Creates a shuffled (integer) list.
	''' </summary>
	''' <param name="items_from"></param>
	''' <param name="items_to"></param>
	''' <param name="items_"></param>
	''' <returns></returns>
	Public Function CreateRandomList(items_from As Integer, items_to As Integer, items_ As List(Of Integer)) As List(Of Integer)
		For i As Integer = items_from To items_to
			RandomList(items_from, items_to + 1, items_)
		Next
		Return items_


		'		Dim temp_list As New List(Of Integer)
		'		g.CreateRandomList(1, 7, temp_list)

		'		'Then use the list, for example:
		'		If temp_list.Count < 1 Then Exit Sub
		'		With temp_list
		'			For j As Integer = 0 To .Count - 1
		'				u.Text &= .Item(j).ToString & vbCrLf
		'			Next
		'		End With

	End Function

	Public Function CycleThrough(c As Control, Optional reverse_ As Boolean = False) As Integer
		Dim l As ListBox, d As ComboBox
		If TypeOf c Is ListBox Then
			l = c
			With l
				If .SelectedIndex < 1 Then .SelectedIndex = 0
				Try
					If reverse_ = True Then
						.SelectedIndex -= 1
					Else
						.SelectedIndex += 1
					End If
				Catch ex As Exception
					.SelectedIndex = 0
				End Try
			End With
			Return l.SelectedIndex
		End If
		If TypeOf c Is ComboBox Then
			d = c
			With d
				If .SelectedIndex < 1 Then .SelectedIndex = 0
				Try
					If reverse_ = True Then
						.SelectedIndex -= 1
					Else
						.SelectedIndex += 1
					End If
				Catch ex As Exception
					.SelectedIndex = 0
				End Try
			End With
			Return d.SelectedIndex
		End If
	End Function
	''' <summary>
	''' Generate random number.
	''' </summary>
	''' <param name="inclusive_min"></param>
	''' <param name="exclusive_max"></param>
	''' <returns></returns>
	Public Shared Function Random_(inclusive_min As Integer, exclusive_max As Integer) As Integer
		Dim generator As New Random
		Dim randomValue As Integer
		randomValue = generator.Next(inclusive_min, exclusive_max)
		Return randomValue
	End Function
	''' <summary>
	''' Creates a sequence of specified number of given character.
	''' </summary>
	''' <param name="times_"></param>
	''' <param name="char_"></param>
	''' <returns></returns>
	Public Shared Function LineBreak(times_ As Integer, Optional char_ As String = "*") As String
		Dim s As String = ""
		For i As Integer = 1 To times_
			s &= char_
		Next
		Return s
	End Function

	''' <summary>
	''' Returns the opposite of the (boolean) value.
	''' </summary>
	''' <param name="boolean_val">True or False</param>
	''' <returns>The opposite of boolean_val</returns>
	Public Shared Function BooleanOpposite(boolean_val As Boolean) As Boolean
		Select Case boolean_val
			Case True
				Return False
			Case False
				Return True
		End Select
	End Function
	''' <summary>
	''' Removes CarriageReturn or CarriageReturn+LineFeed from string.
	''' </summary>
	''' <param name="str_"></param>
	''' <returns></returns>
	Public Shared Function Stripped(str_ As String) As String
		Dim s As String = str_
		s = s.Replace(vbCr, " ").Trim
		s = s.Replace(vbCrLf, " ").Trim
		Return s.Trim
	End Function
	''' <summary>
	''' Splits string into 2 parts without the separator.
	''' </summary>
	''' <param name="string_"></param>
	''' <param name="separator_"></param>
	''' <returns></returns>
	Public Shared Function SplitString(string_ As String, separator_ As String) As Array
		separator_ = separator_.Trim
		Dim return_() As String = {}

		Dim string_parts() As String
		Dim left_ As String
		Dim right_ As String
		string_parts = string_.Split(separator_.ToCharArray, 2)

		If string_parts.Length = 2 Then
			left_ = string_parts(0)
			right_ = string_parts(1)
			If left_.Length > 0 And right_.Length > 0 Then
				return_ = {left_, right_}
			Else
				return_ = {Val(string_)}
			End If
		Else
			return_ = {Val(string_)}
		End If

		Return return_
	End Function
	''' <summary>
	''' Generates an ID using part of GUID combined with specified prefix and/or suffix. You might use NewGUID instead.
	''' </summary>
	''' <param name="case_acct_or_date_time"></param>
	''' <param name="replace_guid"></param>
	''' <returns></returns>
	Public Shared Function NewID(Optional case_acct_or_date_time As String = "date_time", Optional replace_guid As String = "") As String
		Dim m As New Methods
		Dim raw_id As String = System.Guid.NewGuid().ToString, counter As Integer

		If replace_guid.Trim.Length > 0 Then
			raw_id = replace_guid.Trim
			GoTo 2
		End If

		For i% = 1 To raw_id.Length
			If Mid(raw_id, i, 1) = "-" Then
				counter += 1
				If counter = 2 Then
					raw_id = Mid(raw_id, 1, i - 1)
				ElseIf counter = 0 Then
					Exit For
				End If
			End If
		Next

2:
		Select Case case_acct_or_date_time.ToLower
			Case "acct"
				raw_id = raw_id & "-" & My.Computer.Clock.LocalTime.Date.ToShortDateString
			Case "date_time"
				raw_id = raw_id & "-" & Now.Year.ToString & "." & Now.Month.ToString & "." & Now.Day.ToString & "-" & Now.Hour.ToString & "." & m.LeadingZero(Now.Minute.ToString) ' My.Computer.Clock.LocalTime.Date.ToShortDateString
		End Select
		Return raw_id
	End Function
	''' <summary>
	''' Generates a new ID.
	''' </summary>
	''' <param name="prefx"></param>
	''' <param name="suffx"></param>
	''' <returns></returns>
	Public Shared Function NewGUID(Optional prefx As String = "", Optional suffx As Boolean = False) As String
		Dim return__ As String = ""
		If prefx.Length > 0 Then return__ = prefx & "-"
		If suffx = True Then return__ &= Now.ToShortDateString & "-" & Now.ToLongTimeString & "-"
		return__ = return__.Replace(" ", "")
		Dim r As String = return__.Replace("/", "-")
		Dim r_ As String = r.Replace(":", "-")
		'		Dim r2 As String = r_.Replace("--", "-")
		Dim r3 As String = r_.Replace("AM", "-am")
		Dim return_ As String = r3.Replace("PM", "-pm")

		If return_.Length > 0 Then
			'			Return return_ & "-" & System.Guid.NewGuid().ToString
			Return return_ & System.Guid.NewGuid().ToString
		Else
			Return System.Guid.NewGuid().ToString
		End If
	End Function
	''' <summary>
	''' Adds 0 to a number, e.g. from 1 to 01.
	''' </summary>
	''' <param name="number"></param>
	''' <returns></returns>
	Public Shared Function LeadingZero(number As String) As String
		If number.Length < 2 Then Return "0" & number Else : Return number
	End Function

	Public Shared Function GetTimeZone() As String
		Dim offset_ As String = TimeZone.CurrentTimeZone.GetUtcOffset(DateTime.Now).ToString
		Dim suffix_ As String = " (UTC"
		If InStr(offset_, "-") = False Then suffix_ = " (UTC+"
		Return TimeZone.CurrentTimeZone.StandardName & suffix_ & TimeZone.CurrentTimeZone.GetUtcOffset(DateTime.Now).ToString & ")"
	End Function

	Public Shared Function TimeZone_() As String
		Return GetTimeZone()
	End Function

	Public Shared Function FullTimeWithZone() As String
		Return Now.ToLongTimeString & "  " & GetTimeZone()
	End Function

	Public Shared Function FullDateWithZone() As String
		Return Now.ToShortDateString & "  " & GetTimeZone()
	End Function

	Public Shared Function FullDateAsLongWithZone() As String
		Return Now.ToLongDateString & "  " & GetTimeZone()
	End Function

	Public Shared Function Explicit(Optional time_ As String = "", Optional date_ As String = "") As String
		Return FullTimeAsLong(time_, date_) & ",  " & GetTimeZone()
	End Function

	Public Shared Function FullDate(Optional time_ As String = "", Optional date_ As String = "") As String
		'same as FullTime, but puts date first
		Dim val As String = ""
		If time_.Length < 1 Then time_ = Now.ToLongTimeString

		If date_.Length < 1 Then date_ = Now.ToShortDateString

		Return date_ & ",  " & time_
	End Function

	Public Shared Function FullTime(Optional time_ As String = "", Optional date_ As String = "") As String
		'same as FullDate, but puts time first
		Dim val As String = ""
		If time_.Length < 1 Then time_ = Now.ToLongTimeString

		If date_.Length < 1 Then date_ = Now.ToShortDateString

		Return time_ & ",  " & date_
	End Function

	''' <summary>
	''' Returns specified date and time (or current) in Long Date and Long Time, time first. Equivalent to FullTimeAsLong.
	''' </summary>
	''' <param name="time_">The time in any format</param>
	''' <param name="date_">The date in any format</param>
	''' <returns>Date and time</returns>
	Public Shared Function FullDateAsLong(Optional time_ As String = "", Optional date_ As String = "") As String
		Dim val As String = ""
		If time_.Length < 1 Then
			time_ = Now.ToLongTimeString
		Else
			time_ = DateTime.Parse(time_).ToLongTimeString
		End If

		If date_.Length < 1 Then
			date_ = Now.ToLongDateString
		Else
			date_ = DateTime.Parse(date_).ToLongDateString
		End If

		Return date_ & ",  " & time_
	End Function

	''' <summary>
	''' Returns specified date and time (or current) in Long Date and Long Time, time first. Equivalent to FullDateAsLong.
	''' </summary>
	''' <param name="time_">The time in any format</param>
	''' <param name="date_">The date in any format</param>
	''' <returns>Date and time</returns>
	Public Shared Function FullTimeAsLong(Optional time_ As String = "", Optional date_ As String = "") As String
		Dim val As String = ""
		If time_.Length < 1 Then
			time_ = Now.ToLongTimeString
		Else
			time_ = DateTime.Parse(time_).ToLongTimeString
		End If

		If date_.Length < 1 Then
			date_ = Now.ToLongDateString
		Else
			date_ = DateTime.Parse(date_).ToLongDateString
		End If

		Return time_ & ",  " & date_
	End Function
	Public Shared Function FloorValue(val_ As String) As String
		Dim s As Integer
		For i As Integer = 1 To val_.Length
			If Mid(val_, i, 1) = "." Then
				s = i - 1
				Exit For
			End If
		Next
		Return Mid(val_, 1, s)
	End Function
	''' <summary>
	''' Formats number to 2 decimal places.
	''' </summary>
	''' <param name="val_"></param>
	''' <returns></returns>
	Public Shared Function RoundNumber(val_)
		'		If Len(val_.ToString) > 0 Or Val(val_) <> 0 Then
		Try
			'			If Val(val_) <> 0 Then
			Return FormatNumber(Val(val_), 2, TriState.False, TriState.False, TriState.False)
			'			Else
			'			Return 0
			'			End If
		Catch
		End Try
	End Function
	Public Shared Function CalculateSince(date_ As Date, Optional interval_ As DateInterval = DateInterval.Day, Optional suffixed As Boolean = True) As String
		Dim d = DateDiff(interval_, date_, Date.Parse(Now))
		If suffixed Then
			Return ToPlural(d, "days") & " ago"
		Else
			Return d
		End If
	End Function

	''' <summary>
	''' Changes a string to a cutom pluralized string. You will likely need to manually update this function to add new strings.
	''' </summary>
	''' <param name="count_"></param>
	''' <param name="str_to_change"></param>
	''' <param name="rest_of_full_string"></param>
	''' <param name="prefixed"></param>
	''' <returns></returns>
	Public Shared Function ToPlural(count_, str_to_change, Optional rest_of_full_string = "", Optional prefixed = True) As String
		Dim val_ As String = ""

		If Val(count_) = 1 Then
			Select Case str_to_change.ToString.ToLower
				Case "message"
					If prefixed Then
						val_ = "a new message"
					Else
						val_ = "message"
					End If
				Case "messages"
					If prefixed Then
						val_ = "a new message"
					Else
						val_ = "message"
					End If
				Case "mark"
					If prefixed Then
						val_ = "1 mark "
					Else
						val_ = "mark"
					End If
				Case "marks"
					If prefixed Then
						val_ = "1 mark "
					Else
						val_ = "mark"
					End If
				Case "minute"
					If prefixed Then
						val_ = "1 minute "
					Else
						val_ = "minute"
					End If
				Case "minutes"
					If prefixed Then
						val_ = "1 minute "
					Else
						val_ = "minute"
					End If

				Case "day"
					If prefixed Then
						val_ = "1 day "
					Else
						val_ = "day"
					End If
				Case "days"
					If prefixed Then
						val_ = "1 day "
					Else
						val_ = "day"
					End If

				Case "student"
					If prefixed Then
						val_ = "1 student "
					Else
						val_ = "student"
					End If
				Case "students"
					If prefixed Then
						val_ = "1 student "
					Else
						val_ = "student"
					End If

				Case "month"
					If prefixed Then
						val_ = "1 month "
					Else
						val_ = "month"
					End If
				Case "months"
					If prefixed Then
						val_ = "1 month "
					Else
						val_ = "month"
					End If

				Case "year"
					If prefixed Then
						val_ = "1 year "
					Else
						val_ = "year"
					End If
				Case "years"
					If prefixed Then
						val_ = "1 year "
					Else
						val_ = "year"
					End If

				Case "hour"
					If prefixed Then
						val_ = "1 hour "
					Else
						val_ = "hour"
					End If
				Case "hours"
					If prefixed Then
						val_ = "1 hour "
					Else
						val_ = "hour"
					End If

				Case "second"
					If prefixed Then
						val_ = "1 second "
					Else
						val_ = "second"
					End If
				Case "seconds"
					If prefixed Then
						val_ = "1 second "
					Else
						val_ = "second"
					End If

				Case "account"
					If prefixed Then
						val_ = "1 account "
					Else
						val_ = "account"
					End If
				Case "accounts"
					If prefixed Then
						val_ = "1 account "
					Else
						val_ = "account"
					End If

				Case "male"
					If prefixed Then
						val_ = "1 male "
					Else
						val_ = "male"
					End If
				Case "males"
					If prefixed Then
						val_ = "1 male "
					Else
						val_ = "male"
					End If

				Case "female"
					If prefixed Then
						val_ = "1 female "
					Else
						val_ = "female"
					End If
				Case "females"
					If prefixed Then
						val_ = "1 female "
					Else
						val_ = "female"
					End If

				Case "user"
					If prefixed Then
						val_ = "1 user "
					Else
						val_ = "user"
					End If
				Case "users"
					If prefixed Then
						val_ = "1 user "
					Else
						val_ = "user"
					End If

				Case "file"
					If prefixed Then
						val_ = "1 file "
					Else
						val_ = "file"
					End If
				Case "files"
					If prefixed Then
						val_ = "1 file "
					Else
						val_ = "file"
					End If

				Case "record"
					If prefixed Then
						val_ = "1 record "
					Else
						val_ = "record"
					End If
				Case "records"
					If prefixed Then
						val_ = "1 record "
					Else
						val_ = "record"
					End If
			End Select
		Else
			Select Case str_to_change.ToString.ToLower
				Case "message"
					If prefixed Then
						val_ = count_ & " new messages"
					Else
						val_ = "messages"
					End If
				Case "messages"
					If prefixed Then
						val_ = count_ & " new messages"
					Else
						val_ = "messages"
					End If
				Case "mark"
					If prefixed Then
						val_ = count_ & " marks"
					Else
						val_ = "marks"
					End If
				Case "marks"
					If prefixed Then
						val_ = count_ & " marks"
					Else
						val_ = "marks"
					End If
				Case "minute"
					If prefixed Then
						val_ = count_ & " minutes"
					Else
						val_ = "minutes"
					End If
				Case "minutes"
					If prefixed Then
						val_ = count_ & " minutes"
					Else
						val_ = "minutes"
					End If

				Case "day"
					If prefixed Then
						val_ = count_ & " days"
					Else
						val_ = "days"
					End If
				Case "days"
					If prefixed Then
						val_ = count_ & " days"
					Else
						val_ = "days"
					End If

				Case "student"
					If prefixed Then
						val_ = count_ & " students"
					Else
						val_ = "students"
					End If
				Case "students"
					If prefixed Then
						val_ = count_ & " students"
					Else
						val_ = "students"
					End If

				Case "month"
					If prefixed Then
						val_ = count_ & " months"
					Else
						val_ = "months"
					End If
				Case "months"
					If prefixed Then
						val_ = count_ & " months"
					Else
						val_ = "months"
					End If
				Case "year"
					If prefixed Then
						val_ = count_ & " years"
					Else
						val_ = "years"
					End If
				Case "years"
					If prefixed Then
						val_ = count_ & " years"
					Else
						val_ = "years"
					End If
				Case "hour"
					If prefixed Then
						val_ = count_ & " hours"
					Else
						val_ = "hours"
					End If
				Case "hours"
					If prefixed Then
						val_ = count_ & " hours"
					Else
						val_ = "hours"
					End If
				Case "second"
					If prefixed Then
						val_ = count_ & " seconds"
					Else
						val_ = "seconds"
					End If
				Case "seconds"
					If prefixed Then
						val_ = count_ & " seconds"
					Else
						val_ = "seconds"
					End If
				Case "account"
					If prefixed Then
						val_ = count_ & " accounts"
					Else
						val_ = "accounts"
					End If
				Case "accounts"
					If prefixed Then
						val_ = count_ & " accounts"
					Else
						val_ = "accounts"
					End If

				Case "male"
					If prefixed Then
						val_ = count_ & " males"
					Else
						val_ = "males"
					End If
				Case "males"
					If prefixed Then
						val_ = count_ & " males"
					Else
						val_ = "males"
					End If
				Case "female"
					If prefixed Then
						val_ = count_ & " females"
					Else
						val_ = "females"
					End If
				Case "females"
					If prefixed Then
						val_ = count_ & " females"
					Else
						val_ = "females"
					End If
				Case "user"
					If prefixed Then
						val_ = count_ & " users"
					Else
						val_ = "users"
					End If
				Case "users"
					If prefixed Then
						val_ = count_ & " users"
					Else
						val_ = "users"
					End If
				Case "file"
					If prefixed Then
						val_ = count_ & " files"
					Else
						val_ = "files"
					End If
				Case "files"
					If prefixed Then
						val_ = count_ & " files"
					Else
						val_ = "files"
					End If
				Case "record"
					If prefixed Then
						val_ = count_ & " records"
					Else
						val_ = "records"
					End If
				Case "records"
					If prefixed Then
						val_ = count_ & " records"
					Else
						val_ = "records"
					End If
			End Select
		End If

		If rest_of_full_string.Length > 0 Then
			val_ &= " " & rest_of_full_string
		End If

		Return val_
	End Function

	''' <summary>
	''' Determines if val_ is within 0 and 100.
	''' </summary>
	''' <param name="val_"></param>
	''' <returns></returns>
	Public Shared Function InPercent(val_ As String) As Boolean
		val_ = val_.Trim
		If Val(val_) < 0 Or Val(val_) > 100 Or Val(val_) = 0 Then
			Return False
		Else
			Return True
		End If
	End Function
	''' <summary>
	''' Checks if a number is even.
	''' </summary>
	''' <param name="val"></param>
	''' <returns></returns>
	Public Shared Function IsEven(val As Long) As Boolean
		IsEven = (val Mod 2 = 0)
	End Function
	''' <summary>
	''' Boolean to String.
	''' </summary>
	''' <param name="val"></param>
	''' <returns>Yes if True, No if False.</returns>
	Public Shared Function ToBString(val) As String
		Select Case val.ToString
			Case True
				Return "Yes"
			Case False
				Return "No"
		End Select
	End Function


#End Region

#Region "Machine"
	''' <summary>
	''' Creates, adds to, or overwrites the content of file (format is text).
	''' </summary>
	''' <param name="file_">The path to the file.</param>
	''' <param name="txt_">Intended string content of the file.</param>
	''' <param name="append_">Should it add to the content of the file (if it has) or overwrite everything?</param>
	''' <param name="dont_trim">Should trailing spaces be ignored?</param>
	''' <returns>True if the operation is successful, false if there's an exception.</returns>
	Public Shared Function WriteText(file_ As String, txt_ As String, Optional append_ As Boolean = False, Optional dont_trim As Boolean = False) As Boolean
		Try
			If dont_trim Then
				FileIO.FileSystem.WriteAllText(file_, txt_, append_)
			Else
				FileIO.FileSystem.WriteAllText(file_, txt_.Trim, append_)
			End If
			Return True
		Catch ex As Exception
			Return False
		End Try
	End Function
	''' <summary>
	''' Retrieves the content of a file (format is text).
	''' </summary>
	''' <param name="file_">The path to the file.</param>
	''' <param name="dont_trim">Should trailing spaces be ignored?</param>
	''' <returns>The (text) content of the file.</returns>
	Public Shared Function ReadText(file_ As String, Optional dont_trim As Boolean = False) As String
		Dim docName As String = Path.GetFileName(file_)
		Dim docPath As String = Path.GetDirectoryName(file_)
		Dim stream As New FileStream(file_, FileMode.Open)
		Dim reader As New StreamReader(stream)
		Try
			If dont_trim = True Then
				Return reader.ReadToEnd()
			Else
				Return reader.ReadToEnd().Trim
			End If
		Finally
			reader.Dispose()
			stream.Dispose()
		End Try

	End Function
	Public Shared Function ToPath(str_ As String) As String
		Try
			If str_.Substring(str_.Length - 1, 1) <> "\" Then
				Return str_ & "\"
			Else
				Return str_
			End If
		Catch ex As Exception

		End Try
	End Function

#End Region

#Region "Machine"
	''' <summary>
	''' Set permission for file. The exe calling this sub likely need to have administrator privileges.
	''' </summary>
	''' <param name="file_"></param>
	''' <param name="temp_file"></param>
	''' <param name="user_"></param>
	Public Shared Sub PermitFile(file_ As List(Of String), Optional temp_file As String = "", Optional user_ As String = "")
		On Error Resume Next
		'		Dim f_w As New FormatWindow
		Dim t_file As String
		If temp_file.Length < 1 Then temp_file = My.Application.Info.DirectoryPath & "\FilePermissions.bat"
		If user_.Length < 1 Then user_ = Environment.UserName

		Dim str_ As String = "" ' "icacls """ & file_ & """ /grant " & Environment.UserName & ":(F)"
		For i As Integer = 0 To file_.Count - 1
			str_ &= "icacls """ & file_(i) & """ /grant " & user_ & ":(F)" & vbCrLf
		Next
		WriteText(temp_file, str_)
		StartFile(temp_file)

	End Sub

	''' <summary>
	''' Checks if extension of file fits that of an image.
	''' </summary>
	''' <param name="filename_"></param>
	''' <returns></returns>
	Public Shared Function IsImage(filename_ As String) As Boolean
		Select Case System.IO.Path.GetExtension(filename_).ToLower
			Case ".jpg"
				Return True
			Case ".jpeg"
				Return True
			Case ".png"
				Return True
			Case ".gif"
				Return True
			Case ".bmp"
				Return True
			Case ".tif"
				Return True
			Case ".ico"
				Return True
			Case Else
				Return False
		End Select
	End Function

	Public Sub Delete(file_ As String, Optional recycle_ As Boolean = False, Optional showUI_ As Boolean = False)
		Try
			My.Computer.FileSystem.DeleteFile(file_, showUI_, recycle_)
		Catch ex As Exception
		End Try
	End Sub

	Public Function SetTextToClipboard(str_ As String) As Boolean
		Try
			Clipboard.SetText(str_)
			Return True
		Catch ex As Exception
			Return False
		End Try
	End Function
	Public Function GetTextFromClipboard(str_ As String) As String
		Try
			If Clipboard.ContainsText Then
				Return Clipboard.GetText
			End If
		Catch ex As Exception
		End Try
	End Function
	''' <summary>
	''' Set permission for file. The exe calling this sub likely need to have administrator privileges. Same as PermitFile.
	''' </summary>
	''' <param name="file_"></param>
	''' <param name="perm_"></param>
	''' <param name="remove_existing"></param>
	Public Sub PermissionForFile(file_ As String, Optional perm_ As FileSystemRights = FileSystemRights.FullControl, Optional remove_existing As Boolean = False)
		If file_.Length < 1 Then Exit Sub
		If perm_.ToString.Length < 1 Then perm_ = FileSystemRights.FullControl

		Dim user_ As String = Environment.UserName
		'If user_BLANK_FOR_CURRENT.Length < 1 Then
		'	user_ = user__
		'Else
		'	user__ = user_BLANK_FOR_CURRENT.Trim
		'End If

		Dim FilePath As String = file_
		Dim UserAccount As String = user_ ' "Everyone"
		Dim FileInfo As IO.FileInfo = New IO.FileInfo(FilePath)
		Dim FileAcl As New FileSecurity
		FileAcl.AddAccessRule(New FileSystemAccessRule(UserAccount, perm_, AccessControlType.Deny))
		'FolderAcl.SetAccessRuleProtection(True, False) 'uncomment to remove existing permissions
		''        FileInfo.SetAccessControl(FileAcl)
	End Sub
	''' <summary>
	''' Sets folder acccess. Windows related.
	''' </summary>
	''' <param name="folder_">Folder to set permission on</param>
	''' <param name="perm_">Type of permission</param>
	''' <param name="remove_existing">Should existing permissions for this user be overwritten if found?</param>
	Public Sub PermissionForFolder(folder_ As String, Optional perm_ As FileSystemRights = FileSystemRights.FullControl, Optional remove_existing As Boolean = False)
		If folder_.Length < 1 Then Exit Sub
		If perm_.ToString.Length < 1 Then perm_ = FileSystemRights.FullControl

		Dim user_ As String = Environment.UserName
		'If user_BLANK_FOR_CURRENT.Length < 1 Then
		'	user_ = user__
		'Else
		'	user__ = user_BLANK_FOR_CURRENT.Trim
		'End If

		Dim FolderPath As String = folder_ 'Specify the folder here
		Dim UserAccount As String = user_ 'Specify the user here

		Dim FolderInfo As IO.DirectoryInfo = New IO.DirectoryInfo(FolderPath)
		Dim FolderAcl As New DirectorySecurity
		'		'		FolderAcl.AddAccessRule(New FileSystemAccessRule(UserAccount, perm_, InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow))
		FolderAcl.AddAccessRule(New FileSystemAccessRule(UserAccount, perm_, InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow))
		If remove_existing = True Then FolderAcl.SetAccessRuleProtection(True, False) 'uncomment to remove existing permissions
		FolderInfo.SetAccessControl(FolderAcl)
	End Sub

	''' <summary>
	''' Sets folder acccess. Windows related. Same as PermissionForFolder.
	''' </summary>
	''' <param name="folder_">Folder to set permission on</param>
	''' <param name="perm_">Type of permission</param>
	''' <param name="remove_existing">Should existing permissions for this user be overwritten if found?</param>

	Public Shared Sub PermitFolder(folder_ As String, Optional perm_ As FileSystemRights = FileSystemRights.FullControl, Optional remove_existing As Boolean = False) ', Optional user_BLANK_FOR_CURRENT As String = "")
		'		Dim f_w As New FormatWindow
		If folder_.Length < 1 Then Exit Sub
		If perm_.ToString.Length < 1 Then perm_ = FileSystemRights.FullControl

		Dim user_ As String = Environment.UserName
		'If user_BLANK_FOR_CURRENT.Length < 1 Then
		'	user_ = f_w.user__
		'Else
		'	f_w.user__ = user_BLANK_FOR_CURRENT.Trim
		'End If

		Dim FolderPath As String = folder_ 'Specify the folder here
		Dim UserAccount As String = user_ 'Specify the user here

		Dim FolderInfo As IO.DirectoryInfo = New IO.DirectoryInfo(FolderPath)
		Dim FolderAcl As New DirectorySecurity
		'		'		FolderAcl.AddAccessRule(New FileSystemAccessRule(UserAccount, perm_, InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow))
		FolderAcl.AddAccessRule(New FileSystemAccessRule(UserAccount, perm_, InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow))
		If remove_existing = True Then FolderAcl.SetAccessRuleProtection(True, False) 'uncomment to remove existing permissions
		FolderInfo.SetAccessControl(FolderAcl)
	End Sub
	''' <summary>
	''' Creates registry entry for a path, to start with PC logs on.
	''' </summary>
	''' <param name="file_"></param>
	''' <param name="key_"></param>
	Public Sub ToStartup(file_ As String, key_ As String)
		If file_.Length < 1 Or key_.Length < 1 Then Exit Sub
		Try
			My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run", key_, file_)
		Catch x As Exception
		End Try
	End Sub
	''' <summary>
	''' Removes registry entry for a path initially set to start with PC logs on. Opposite of ToStartup.
	''' </summary>
	''' <param name="file_"></param>
	''' <param name="key_"></param>
	Public Sub RemoveFromStartup(file_ As String, key_ As String)
		If file_.Length < 1 Or key_.Length < 1 Then Exit Sub

		Using key As RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey("Software")
			key.DeleteSubKey("Software\Microsoft\Windows\CurrentVersion\Run\" & key_)
		End Using

		Dim exists As Boolean = False
		Try
			If My.Computer.Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Run\" & key_) IsNot Nothing Then
				'				exists = True
			End If
		Finally
			'			My.Computer.Registry.CurrentUser.Close()
		End Try

	End Sub

	''' <summary>
	''' Sets file/folder attribute.
	''' </summary>
	''' <param name="file_OR_directory"></param>
	''' <param name="show_"></param>
	''' <param name="remove_"></param>
	Public Sub SetAttribute(file_OR_directory As String, Optional show_ As Boolean = False, Optional remove_ As Boolean = True)
		If show_ = False And remove_ = False Then
			Try
				SetAttr(file_OR_directory, FileAttribute.Hidden)
				Exit Sub
			Catch ex As Exception
			End Try
		End If

		If show_ = True Then
			Try
				SetAttr(file_OR_directory, FileAttribute.Normal)
				Exit Sub
			Catch ex As Exception
			End Try
		Else
		End If

		If remove_ = True Then
			Try
				SetAttr(file_OR_directory, FileAttribute.Hidden + FileAttribute.System)
				Exit Sub
			Catch ex As Exception
			End Try
		Else
		End If

	End Sub

	''' <summary>
	''' Creates shortcut to a path.
	''' </summary>
	''' <param name="target_"></param>
	''' <param name="folder_"></param>
	''' <param name="filename_without_extension"></param>
	''' <param name="icon_file"></param>
	Public Sub CreateShortcut(target_ As String, folder_ As String, filename_without_extension As String, Optional icon_file As String = "")
		Dim wsh As Object = CreateObject("WScript.Shell")

		wsh = CreateObject("WScript.Shell")

		Dim MyShortcut, DesktopPath

		DesktopPath = folder_

		MyShortcut = wsh.CreateShortcut(DesktopPath & "\" & filename_without_extension & ".lnk")

		MyShortcut.TargetPath = wsh.ExpandEnvironmentStrings(target_)

		MyShortcut.WorkingDirectory = wsh.ExpandEnvironmentStrings(folder_)

		MyShortcut.WindowStyle = 4

		If icon_file IsNot Nothing And icon_file.Length > 1 Then
			MyShortcut.IconLocation = wsh.ExpandEnvironmentStrings(icon_file)
		End If

		MyShortcut.Save()

	End Sub
	''' <summary>
	''' Checks if VLC Media Player is running.
	''' </summary>
	''' <returns></returns>
	Public Function PlayerIsOn() As Boolean
		Dim p() As Process = Process.GetProcessesByName("vlc")
		PlayerIsOn = p.Count > 0

	End Function

	''' <summary>
	''' Checks if Windows Media Player is running.
	''' </summary>
	''' <returns></returns>
	Public Function MusicIsOn() As Boolean
		Dim p() As Process = Process.GetProcessesByName("wmplayer")
		MusicIsOn = p.Count > 0

	End Function
	''' <summary>
	''' Checks if an application is running.
	''' </summary>
	''' <returns></returns>
	Public Shared Function AppIsOn(name_without_extension As String) As Boolean
		Dim p() As Process = Process.GetProcessesByName(name_without_extension)
		AppIsOn = p.Count > 0
	End Function
	''' <summary>
	''' Checks if an application is not running. Opposite of AppIsOn.
	''' </summary>
	''' <returns></returns>
	Public Shared Function AppNotOn(name_without_extension As String) As Boolean
		Dim p() As Process = Process.GetProcessesByName(name_without_extension)
		AppNotOn = p.Count < 1
	End Function

	''' <summary>
	''' Stops an application from running.
	''' </summary>
	Public Shared Sub KillProcess(process_name_without_extension As String)
		For Each proc As Process In Process.GetProcessesByName(process_name_without_extension)
			proc.Kill()
		Next
	End Sub

	''' <summary>
	''' Opens apps from JSON string.
	''' </summary>
	''' <param name="files_list_semiJSON">JSON string</param>
	''' <param name="style_">For StartApp()</param>
	''' <param name="style__">For StartApp()</param>
	''' <param name="checkFirst">For StartApp()</param>
	Public Shared Sub StartApps(files_list_semiJSON As String, Optional style_ As ProcessWindowStyle = ProcessWindowStyle.Normal, Optional style__ As Boolean = False, Optional checkFirst As Boolean = True)
		'open apps if not already open
		Dim v As String = files_list_semiJSON.Trim
		If v.Length < 1 Then Return
		v = v.Replace("\", "\\")
		Dim s As New List(Of String)
		Try
			s.Clear()
		Catch ex As Exception
		End Try
		s = DatabaseToListJSON(v)
		For i As Integer = 0 To s.Count - 1
			StartFile(s(i), style_, style__, checkFirst)
		Next

	End Sub


	''' <summary>
	''' Depracated. Use StartFile() or StartApps instead.
	''' </summary>
	''' <param name="file_">The path to the file.</param>
	''' <param name="style_">ProcessWindowStyle (normal, hidden etc)</param>
	''' <param name="style__">Ignore ProcessWindowStyle, completely hide it.</param>
	''' <param name="checkFirst">Check if an instance of the file is already running (in Tasks).</param>
	''' <returns>True if file exists and is opened, false if not.</returns>

	Public Function StartApp(file_ As String, Optional style_ As ProcessWindowStyle = ProcessWindowStyle.Normal, Optional style__ As Boolean = False, Optional checkFirst As Boolean = True) As Boolean
		If checkFirst = True Then
			Try
				If AppIsOn(Path.GetFileNameWithoutExtension(file_)) Then
					'					Return False
					Exit Function
				End If
			Catch
			End Try
		End If

		Dim s As New ProcessStartInfo
		If style__ = True Then
			s.WindowStyle = ProcessWindowStyle.Hidden
		Else
			s.WindowStyle = style_
		End If
		s.FileName = file_
		Try
			Process.Start(s)
			Return True
		Catch ex As Exception
			Return False
		End Try

2:
	End Function

	''' <summary>
	''' Opens a file. If checkFirst is True, then the program file won't run if it is already running; this is ignored if it's not a program file.
	''' </summary>
	''' <param name="file_">The path to the file.</param>
	''' <param name="style_">ProcessWindowStyle (normal, hidden etc)</param>
	''' <param name="style__">Ignore ProcessWindowStyle, completely hide it.</param>
	''' <param name="checkFirst">Check if an instance of the file is already running (in Tasks).</param>
	''' <returns>True if file exists and is opened, false if not.</returns>
	Public Shared Function StartFile(file_ As String, Optional style_ As ProcessWindowStyle = ProcessWindowStyle.Normal, Optional style__ As Boolean = False, Optional checkFirst As Boolean = False) As Boolean
		If checkFirst = True Then
			Try
				If AppIsOn(Path.GetFileNameWithoutExtension(file_)) Then
					'					Return False
					Exit Function
				End If
			Catch
			End Try
		End If

		Dim s As New ProcessStartInfo
		If style__ = True Then
			s.WindowStyle = ProcessWindowStyle.Hidden
		Else
			s.WindowStyle = style_
		End If
		s.FileName = file_
		Try
			Process.Start(s)
			Return True
		Catch ex As Exception
			Return False
		End Try

2:
	End Function

	''' <summary>
	''' Discontinued. Use StartFile instead.
	''' </summary>
	''' <param name="file_"></param>
	''' <param name="checkFirst"></param>
	''' <param name="style_"></param>
	''' <param name="style__"></param>
	''' <returns></returns>
	Public Shared Function OpenFile(file_ As String, Optional checkFirst As Boolean = True, Optional style_ As ProcessWindowStyle = ProcessWindowStyle.Normal, Optional style__ As Boolean = False) As Boolean
		If checkFirst = True Then
			Try
				If AppIsOn(Path.GetFileNameWithoutExtension(file_)) Then
					Exit Function
				End If
			Catch
			End Try
		End If

		Dim s As New ProcessStartInfo
		If style__ = True Then
			s.WindowStyle = ProcessWindowStyle.Hidden
		Else
			s.WindowStyle = style_
		End If
		s.FileName = file_
		Try
			Process.Start(s)
			Return True
		Catch ex As Exception
			Return False
		End Try

2:
	End Function
	''' <summary>
	''' Starts a program, typically with an argument. Equivalent to StartFile.
	''' </summary>
	''' <param name="file_"></param>
	''' <param name="arg_"></param>
	''' <param name="checkFirst"></param>
	''' <returns></returns>
	Public Shared Function StartFileWithArgument(file_ As String, Optional arg_ As String = Nothing, Optional checkFirst As Boolean = False) As Boolean
		If checkFirst = True Then
			Try
				If AppIsOn(Path.GetFileNameWithoutExtension(file_)) Then
					Exit Function
				End If
			Catch
			End Try
		End If

		Try
			Process.Start(file_, arg_)
			Return True
		Catch ex As Exception
			Return False
		End Try
	End Function

	Public Shared Function SaveFile(_filename, _filter) As Array
		Dim f_ As New SaveFileDialog
		Dim return_() As String

		With f_
			.FileName = _filename
			.Filter = _filter
			.ShowDialog()

			If .FileName.Trim <> "" Then
				return_ = {True, .FileName}
			ElseIf .FileName.Trim = "" Then
				return_ = {False, .FileName}
			End If
		End With
		Return return_
	End Function

	Public Shared Function DiskLocation(default_string As String, Optional description_ As String = "", Optional root_folder As String = "", Optional show_new_folder_button As Boolean = True) As Array
		Dim f_ As New FolderBrowserDialog
		Dim return_() As String

		With f_
			.Description = description_
			If root_folder.Length > 0 Then .RootFolder = root_folder
			.ShowNewFolderButton = show_new_folder_button

			.ShowDialog()

			If .SelectedPath.Trim <> "" Then
				return_ = {True, .SelectedPath}
			ElseIf .SelectedPath.Trim = "" Then
				return_ = {False, default_string}
			End If
		End With
		Return return_
	End Function


	Public Shared Sub Hibernate()
		Try
			Application.SetSuspendState(PowerState.Hibernate, True, True)
		Catch
		End Try
	End Sub

	Public Shared Sub LogOff()
		Try
			Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.System) & "\shutdown.exe", "-l")
		Catch
		End Try
	End Sub
	Public Shared Sub Restart()
		Try
			Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.System) & "\shutdown.exe", "-r -t 0")
		Catch
		End Try
	End Sub
	Public Shared Sub Shutdown()
		Try
			Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.System) & "\shutdown.exe", "-s -t 0")
		Catch
		End Try
	End Sub
	''' <summary>
	''' Stops the current application from running, or another application if process_name_without_extension is specified.
	''' </summary>
	''' <param name="process_name_without_extension"></param>
	Public Shared Sub ExitApp(Optional process_name_without_extension As String = "")
		Select Case process_name_without_extension.Length
			Case < 1
				Environment.Exit(0)
			Case Else
				If AppIsOn(process_name_without_extension) Then KillProcess(process_name_without_extension)
		End Select
	End Sub
	''' <summary>
	''' Plays a sound file.
	''' </summary>
	''' <param name="file_"></param>
	''' <param name="loop_"></param>

	Public Shared Sub PlayAudio(file_ As String, Optional loop_ As Boolean = False)
		If file_.Length < 1 Or file_ Is Nothing Then Exit Sub

		Dim mode_ As AudioPlayMode
		If loop_ = True Then
			mode_ = AudioPlayMode.BackgroundLoop
		Else
			mode_ = AudioPlayMode.Background
		End If

		Try
			My.Computer.Audio.Play(file_, mode_)
		Catch ex As Exception
		End Try
	End Sub
	''' <summary>
	''' Stops playing a sound file.
	''' </summary>

	Public Shared Sub StopAudio()
		Try
			My.Computer.Audio.Stop()
		Catch ex As Exception
		End Try
	End Sub
	''' <summary>
	''' Checks if a file exists.
	''' </summary>
	''' <param name="file_"></param>
	''' <returns></returns>

	Public Shared Function Exists(file_ As String) As Boolean
		If file_.Length < 1 Then
			Return False
			Exit Function
		End If
		Exists = My.Computer.FileSystem.FileExists(file_) Or My.Computer.FileSystem.DirectoryExists(file_)
	End Function

#End Region

End Class
