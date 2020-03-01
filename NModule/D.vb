Imports NModule.Methods
Imports System.Drawing
Imports System.Windows.Forms
Imports NModule.NFunctions
Public Class D
	''' <summary>
	''' Binds Question Types to Control. You can adjust the list in Web_Module.Methods.QuestionTypeList.
	''' </summary>
	''' <param name="d_">ComboBox.</param>
	''' <param name="FirstIsEmpty">Should the first item appear empty? Default is False.</param>
	''' <param name="sort_">Should the items of the Combobox be sorted after population?</param>
	Public Shared Sub QuestionTypeDrop(d_ As ComboBox, Optional FirstIsEmpty As Boolean = False, Optional sort_ As Boolean = False)
		Dim web_methods_ As New Web_Module.Methods

		If d_.Items.Count > 0 Then Exit Sub

		Clear(d_)

		If FirstIsEmpty Then d_.Items.Add("")

		Dim l As List(Of String) = web_methods_.QuestionTypeList()
		For i As Integer = 0 To l.Count - 1
			d_.Items.Add(l(i).ToString)
		Next
		d_.Sorted = sort_
	End Sub
	''' <summary>
	''' Binds list of Map Providers supported by GMap control to Control. You can adjust the list in NFunctions.
	''' </summary>
	''' <param name="d_">ComboBox or ListBox.</param>
	''' <param name="sorted_">Should the items of d_ (if Combobox) be sorted after population?</param>
	Public Shared Sub MapDrop(d_ As Control, Optional sorted_ As Boolean = False)
		If TypeOf d_ IsNot ComboBox And TypeOf d_ IsNot ListBox Then Exit Sub
		Clear(d_)
		Dim c_ As ComboBox
		Dim l As ListBox
		Dim n As New NFunctions
		Dim l_ As New List(Of String)
		l_ = n.MapProvidersList

		If TypeOf d_ Is ComboBox Then
			c_ = d_
			c_.Sorted = sorted_

			For i As Integer = 0 To l_.Count - 1
				c_.Items.Add(l_(i).ToString)
			Next

			c_.SelectedIndex = 0
			Exit Sub
		End If

		If TypeOf d_ Is ListBox Then
			l = d_

			For i As Integer = 0 To l_.Count - 1
				l.Items.Add(l_(i).ToString)
			Next

			l.SelectedIndex = 0
			Exit Sub
		End If
	End Sub
	''' <summary>
	''' Binds the types of tasks in Bookkeeping to ComboBox. You can adjust the list in Web_Module.Methods.TaskTypeList.
	''' </summary>
	''' <param name="d_">The ComboBox to bind to.</param>
	''' <param name="FirstIsEmpty">Should the first item appear empty? Default is False.</param>
	''' <param name="sort_">Should the items of the Combobox be sorted after population?</param>
	Public Shared Sub TaskTypeDrop(d_ As ComboBox, Optional FirstIsEmpty As Boolean = False, Optional sort_ As Boolean = False)
		Dim web_methods_ As New Web_Module.Methods

		If d_.Items.Count > 0 Then Exit Sub

		Clear(d_)

		If FirstIsEmpty Then d_.Items.Add("")

		Dim l As List(Of String) = web_methods_.TaskTypeList()
		For i As Integer = 0 To l.Count - 1
			d_.Items.Add(l(i).ToString)
		Next
		d_.Sorted = sort_
	End Sub

	''' <summary>
	''' Binds Feedback types to ComboBox. You can adjust the list in Web_Module.Methods.ReminderTypeList.
	''' </summary>
	''' <param name="d_">The ComboBox to bind to.</param>
	''' <param name="FirstIsEmpty">Should the first item appear empty? Default is False.</param>
	''' <param name="sort_">Should the items of the Combobox be sorted after population?</param>
	Public Shared Sub ReminderTypeDrop(d_ As ComboBox, Optional FirstIsEmpty As Boolean = False, Optional sort_ As Boolean = True)
		Dim web_methods_ As New Web_Module.Methods
		'		Dim d As New Web_Module.DataConnectionDesktop

		If d_.Items.Count > 0 Then Exit Sub

		Clear(d_)

		If FirstIsEmpty Then d_.Items.Add("")

		Dim l As List(Of String) = web_methods_.ReminderTypeList()
		For i As Integer = 0 To l.Count - 1
			d_.Items.Add(l(i).ToString)
		Next
		d_.Sorted = sort_
	End Sub

	''' <summary>
	''' Binds Reminder intervals to ComboBox. You can adjust the list in Web_Module.Methods.ReminderList.
	''' </summary>
	''' <param name="d_">The ComboBox to bind to.</param>
	''' <param name="FirstIsEmpty">Should the first item appear empty? Default is False.</param>
	''' <param name="sort_">Should the items of the Combobox be sorted after population?</param>
	Public Shared Sub ReminderDrop(d_ As ComboBox, Optional FirstIsEmpty As Boolean = False, Optional sort_ As Boolean = True)
		Dim web_methods_ As New Web_Module.Methods

		If d_.Items.Count > 0 Then Exit Sub

		Clear(d_)

		If FirstIsEmpty Then d_.Items.Add("")

		Dim l As List(Of String) = web_methods_.ReminderList()
		For i As Integer = 0 To l.Count - 1
			d_.Items.Add(l(i).ToString)
		Next
		d_.Sorted = sort_
	End Sub

	''' <summary>
	''' Binds Reminder recurrence intervals to ComboBox. You can adjust the list in Web_Module.Methods.RecurrenceList.
	''' </summary>
	''' <param name="d_">The ComboBox to bind to.</param>
	''' <param name="FirstIsEmpty">Should the first item appear empty? Default is False.</param>
	''' <param name="sort_">Should the items of the Combobox be sorted after population?</param>

	Public Shared Sub RecurrenceDrop(d_ As ComboBox, Optional FirstIsEmpty As Boolean = False, Optional sort_ As Boolean = True)
		Dim web_methods_ As New Web_Module.Methods
		If d_.Items.Count > 0 Then Exit Sub

		Clear(d_)

		If FirstIsEmpty Then d_.Items.Add("")

		Dim l As List(Of String) = web_methods_.RecurrenceList()
		For i As Integer = 0 To l.Count - 1
			d_.Items.Add(l(i).ToString)
		Next
		d_.Sorted = sort_
	End Sub
	''' <summary>
	''' Binds Tasks status to ComboBox. You can adjust the list in Web_Module.Methods.StatusList.
	''' </summary>
	''' <param name="d_">The ComboBox to bind to.</param>
	''' <param name="IsUpdate">Set this to true to not add started to the list.</param>
	''' <param name="FirstIsEmpty">Should the first item appear empty? Default is False.</param>
	''' <param name="sort_">Should the items of the Combobox be sorted after population?</param>
	Public Shared Sub StatusDrop(d_ As ComboBox, Optional IsUpdate As Boolean = False, Optional FirstIsEmpty As Boolean = False, Optional sort_ As Boolean = True)
		Dim web_methods_ As New Web_Module.Methods

		If d_.Items.Count > 0 Then Exit Sub

		Clear(d_)

		If FirstIsEmpty Then d_.Items.Add("")

		Dim l As List(Of String) = web_methods_.StatusList(IsUpdate)
		For i As Integer = 0 To l.Count - 1
			d_.Items.Add(l(i).ToString)
		Next
		d_.Sorted = sort_
	End Sub
	''' <summary>
	''' Enables or disables controls.
	''' </summary>
	''' <param name="control_">Controls in an array.</param>
	''' <param name="state_">True (Enable) or False (Disable). Default is True.</param>
	''' <example>
	''' EnableControls({Button1, ComboBox1, PictureBox1}, False)
	''' </example>
	Public Shared Sub EnableControls(control_ As Array, Optional state_ As Boolean = True)
		On Error Resume Next
		If control_.Length < 1 Then Exit Sub
		Dim control__ As Control
		For i As Integer = 0 To control_.Length - 1
			control__ = control_(i)
			control__.Enabled = state_
		Next

	End Sub
	''' <summary>
	''' Enables or Disables a single control.
	''' </summary>
	''' <param name="control">Control to enable/disable.</param>
	''' <param name="state_">True (Enable) or False (Disable). Default is True.</param>
	''' <example>
	''' For Each b As Control In Me.Controls
	''' If TypeOf b Is Button Then EnableControl(b)
	''' Next
	''' </example>
	Public Shared Sub EnableControl(control As Control, Optional state_ As Boolean = True)
		On Error Resume Next
		control.Enabled = state_

		'to enable
		'		For Each b As Control In Me.Controls
		'			If TypeOf b Is Button Then g.EnableControl(b)
		'		Next
		'to disable
		'		For Each b As Control In Me.Controls
		'			If TypeOf b Is Button Then g.EnableControl(b, False)
		'		Next

	End Sub
	''' <summary>
	''' Disables a single control.
	''' </summary>
	''' <param name="control">Control to disable.</param>
	''' <example>
	''' For Each b As Control In Me.Controls
	''' If TypeOf b Is Button Then DisableControl(b)
	''' Next
	''' </example>

	Public Shared Sub DisableControl(control As Control)
		EnableControl(control, False)

		'		For Each b As Control In Me.Controls
		'			If TypeOf b Is Button Then g.DisableControl(b)
		'		Next

	End Sub


	''' <summary>
	''' Populates ComboBox with numbers.
	''' </summary>
	''' <param name="d_">ComboBox to fill</param>
	''' <param name="end_">Ending Number</param>
	''' <param name="start_">Beginning Number</param>
	''' <param name="firstItemIsEmpty"></param>
	Public Shared Sub NumberDrop(d_ As ComboBox, end_ As Integer, Optional start_ As Integer = 0, Optional firstItemIsEmpty As Boolean = False)
		If d_.Items.Count > 0 Then Exit Sub

		With d_.Items
			Try
				If firstItemIsEmpty = True And d_.DataSource Is Nothing Then .Add("")
			Catch
			End Try
			For i As Integer = start_ To end_
				.Add(i)
			Next
		End With
	End Sub
	''' <summary>
	''' Determines if all controls and/or file(s) have text.
	''' </summary>
	''' <param name="controls_">Controls or File paths or both</param>
	''' <returns>True or False</returns>

	Public Shared Function IsEmpty(controls_ As Array) As Boolean
		Dim counter_ As Integer = 0
		With controls_
			For i As Integer = 0 To .Length - 1
				If IsEmpty(controls_(i), True, "", "", Nothing) Then
					counter_ += 1
				End If
			Next
		End With
		Return Val(counter_) = controls_.Length
	End Function

	''' <summary>
	''' Determines if control or file has text.
	''' </summary>
	''' <param name="c_">File path or ComboBox or TextBox or PictureBox or NumericUpDown</param>
	''' <param name="use_trim">Should content be trimmed before check?</param>
	''' <param name="control_to_focus">Focus on this when IsEmpty is True</param>
	''' <param name="string_feedback">What to say to user</param>
	''' <returns>True or False</returns>
	Public Shared Function IsEmpty(c_ As Object, Optional use_trim As Boolean = True, Optional string_feedback As String = "", Optional app_ As String = "", Optional control_to_focus As Control = Nothing) As Boolean
		Dim n__ As NumericUpDown
		Dim p__ As PictureBox
		Dim d__ As ComboBox
		Dim t__ As TextBox
		Dim return_val As Boolean = False

		If TypeOf c_ Is PictureBox Then
			p__ = c_
			If p__.BackgroundImage Is Nothing And p__.Image Is Nothing Then
				return_val = True
			End If
		ElseIf TypeOf c_ Is NumericUpDown Then
			n__ = c_
			If (n__.Value) = 0 Then return_val = True
		ElseIf TypeOf c_ Is ComboBox Then
			d__ = c_
			If use_trim = True Then
				If d__.Text.Trim.Length < 1 Then
					return_val = True
				End If
			Else
				If d__.Text.Length < 1 Then
					return_val = True
				End If
			End If
		ElseIf TypeOf c_ Is TextBox Then
			t__ = c_
			If use_trim = True Then
				If t__.Text.Trim.Length < 1 Then
					return_val = True
				End If
			Else
				If t__.Text.Length < 1 Then
					return_val = True
				End If
			End If
		Else
			Try
				If ReadText(c_).Length < 1 Then
					return_val = True
				End If
			Catch ex As Exception
			End Try
		End If
		If return_val = True And string_feedback.Length > 0 And app_.Length > 0 Then
			Try
				Dim feedback_ As New Feedback.Feedback()
				feedback_.fFeedback(string_feedback)
			Catch
			End Try
		End If
		If control_to_focus IsNot Nothing Then
			Try
				control_to_focus.Focus()
			Catch ex As Exception
			End Try
		ElseIf control_to_focus Is Nothing Then
			Try
				c_.Focus()
			Catch ex As Exception
			End Try
		End If
		Return return_val
	End Function

	''' <summary>
	''' Binds Titles Of Courtesy to ComboBox.
	''' </summary>
	''' <param name="c_">ComboBox to bind to.</param>
	''' <param name="first_item_is_empty">Should the first item appear empty? Default is False.</param>
	Public Shared Sub TitleDrop(c_ As ComboBox, Optional first_item_is_empty As Boolean = False)
		If c_.Items.Count > 0 Then Exit Sub

		'Dim d As New Web_Module.DataConnectionDesktop
		Clear(c_)
		With c_
			With .Items
				If first_item_is_empty Then .Add("")
				.Add("Mr.")
				.Add("Mrs.")
				.Add("Ms.")
			End With
			.Sorted = True
			.SelectedIndex = -1
			.Text = ""
		End With

	End Sub
	''' <summary>
	''' Determines if control does not have text. Opposite of IsEmpty.
	''' </summary>
	''' <param name="control_">Control that supports Text property.</param>
	''' <returns>True or False</returns>

	Public Shared Function NotEmpty(control_ As Control) As Boolean
		If TypeOf control_ Is TextBox Or TypeOf control_ Is ComboBox Then
			If control_.Text.Trim.Length > 0 Then
				Return True
			Else
				Return False
			End If
		End If
	End Function

	''' <summary>
	''' Returns the text inside a control that supports Text property or a file.
	''' </summary>
	''' <param name="control_">Control or file path.</param>
	''' <returns></returns>
	Public Shared Function Content(control_ As Object)
		Dim nud As NumericUpDown
		Dim p As PictureBox
		Dim h As CheckBox
		Dim d As DateTimePicker
		Try
			'If trim_ Then
			If TypeOf control_ Is NumericUpDown Then
				nud = control_
				Return nud.Value
			ElseIf TypeOf control_ Is CheckBox Then
				h = control_
				Return h.Checked
			ElseIf TypeOf control_ Is DateTimePicker Then
				d = control_
				Return d.Value
			ElseIf TypeOf control_ Is String Then
				Return ReadText(control_)
				'			ElseIf TypeOf control_ Is picturebox Then
				'				p = control_
				'				If p.BackgroundImage IsNot Nothing Then
				'				Return p.BackgroundImage
				'			Else
				'				Return p.Image
				'			End If
			Else
				Try
					Return control_.Text.Trim
				Catch ex As Exception
				End Try
			End If
		Catch
		End Try
	End Function
	''' <summary>
	''' Determines if control has text or item.
	''' </summary>
	''' <param name="c_">Control to check (TextBox or ComboBox or ListBox or DataGridView or PictureBox or NumericUpDown)</param>
	''' <param name="use_trim_">Should the content be trimmed or not?</param>
	''' <returns>True if it's empty or without content, False otherwise.</returns>
	Public Shared Function HasData(c_ As Control, Optional use_trim_ As Boolean = True) As Boolean
		Dim p__ As PictureBox
		Dim n__ As NumericUpDown

		Dim t_ As TextBox
		Dim d_ As ComboBox
		Dim l_ As ListBox
		Dim g__ As DataGridView

		If TypeOf c_ Is PictureBox Then
			p__ = c_
			Return (p__.BackgroundImage Is Nothing And p__.Image Is Nothing)
		ElseIf TypeOf c_ Is TextBox Then
			t_ = c_
			If use_trim_ = True Then
				Return t_.Text.Trim.Length > 0
			Else
				Return t_.Text.Length > 0
			End If
		ElseIf TypeOf c_ Is ComboBox Then
			d_ = c_
			Return d_.Items.Count > 0
		ElseIf TypeOf c_ Is NumericUpDown Then
			n__ = c_
			Return (n__.Value) <> 0
		ElseIf TypeOf c_ Is ListBox Then
			l_ = c_
			Return l_.Items.Count > 0
		ElseIf TypeOf c_ Is DataGridView Then
			g__ = c_
			Return g__.Rows.Count > 0
		End If
	End Function

	''' <summary>
	''' Attempts to clear control of text, image, items and DataSource.
	''' </summary>
	''' <param name="c_">Control to clear.</param>
	''' <param name="initial_string">String to set as Text.</param>
	Public Shared Sub Clear(c_ As Control, Optional initial_string As String = "")

		Dim c As ComboBox
		Dim l As ListBox
		Dim t As TextBox
		Dim p As PictureBox
		Dim h As CheckBox
		If TypeOf c_ Is CheckBox Then
			h = c_
			h.Checked = False
		End If
		If TypeOf c_ Is ComboBox Then
			c = c_
			Try
				c.DataSource = Nothing
			Catch ex As Exception
			End Try
			c.Items.Clear()
			c.Text = initial_string
		End If
		If TypeOf c_ Is ListBox Then
			l = c_
			Try
				l.DataSource = Nothing
			Catch ex As Exception
			End Try
			l.Items.Clear()
		End If
		If TypeOf c_ Is TextBox Then
			t = c_
			t.Text = initial_string
		End If
		If TypeOf c_ Is PictureBox Then
			p = c_
			Try
				p.Image = Nothing
			Catch ex As Exception
			End Try
			Try
				p.BackgroundImage = Nothing
			Catch ex As Exception
			End Try
		End If
	End Sub

	''' <summary>
	''' Binds Gender to ComboBox. You can adjust the list in Web_Module.Methods.GenderList.
	''' </summary>
	''' <param name="d_">ComboBox to bind to.</param>
	''' <param name="FirstIsEmpty">Should the first item appear empty? Default is False.</param>
	Public Shared Sub GenderDrop(d_ As ComboBox, Optional FirstIsEmpty As Boolean = False)
		Dim web_methods_ As New Web_Module.Methods

		If d_.Items.Count > 0 Then Exit Sub

		Clear(d_)

		If FirstIsEmpty Then d_.Items.Add("")

		Dim l As List(Of String) = web_methods_.GenderList()
		For i As Integer = 0 To l.Count - 1
			d_.Items.Add(l(i).ToString)
		Next
	End Sub

	''' <summary>
	''' Converts boolean values to user-friendly pattern e.g. yes, never.
	''' To convert to boolean, use DropTextToBoolean(str_ As String).
	''' </summary>
	''' <param name="boolean_val">True or False</param>
	''' <param name="pattern_">always/never (default) OR a/n OR a, yes/no OR y/n OR y, on/off OR o/f OR o, 1/0, true/false OR t/f OR t, if/not OR i/n OR i</param>
	''' <returns>Always/Never (default), Yes/No, On/Off, 1/0, True/Talse, If possible/Not at all</returns>
	Public Shared Function BooleanToDropText(boolean_val As Boolean, Optional pattern_ As String = "always/never") As String

		Select Case Convert.ToBoolean(boolean_val)
			Case True
				If pattern_.ToLower = "" Then
					Return "Always"
				End If
				If pattern_.ToLower = "always/never" Then
					Return "Always"
				End If
				If pattern_.ToLower = "a/n" Then
					Return "Always"
				End If
				If pattern_.ToLower = "a" Then
					Return "Always"
				End If
				If pattern_.ToLower = "yes/no" Then
					Return "Yes"
				End If
				If pattern_.ToLower = "y/n" Then
					Return "Yes"
				End If
				If pattern_.ToLower = "y" Then
					Return "Yes"
				End If
				If pattern_.ToLower = "on/off" Then
					Return "On"
				End If
				If pattern_.ToLower = "o/f" Then
					Return "On"
				End If
				If pattern_.ToLower = "o" Then
					Return "On"
				End If
				If pattern_.ToLower = "1/0" Then
					Return "1"
				End If
				If pattern_.ToLower = "true/false" Then
					Return "True"
				End If
				If pattern_.ToLower = "t/f" Then
					Return "True"
				End If
				If pattern_.ToLower = "t" Then
					Return "True"
				End If
				If pattern_.ToLower = "if/not" Then
					Return "If possible"
				End If
				If pattern_.ToLower = "i/n" Then
					Return "If possible"
				End If
				If pattern_.ToLower = "i" Then
					Return "If possible"
				End If
			Case False
				If pattern_.ToLower = "" Then
					Return "Never"
				End If
				If pattern_.ToLower = "always/never" Then
					Return "Never"
				End If
				If pattern_.ToLower = "a/n" Then
					Return "Never"
				End If
				If pattern_.ToLower = "a" Then
					Return "Never"
				End If
				If pattern_.ToLower = "yes/no" Then
					Return "No"
				End If
				If pattern_.ToLower = "y/n" Then
					Return "No"
				End If
				If pattern_.ToLower = "y" Then
					Return "No"
				End If
				If pattern_.ToLower = "on/off" Then
					Return "Off"
				End If
				If pattern_.ToLower = "o/f" Then
					Return "Off"
				End If
				If pattern_.ToLower = "o" Then
					Return "Off"
				End If
				If pattern_.ToLower = "1/0" Then
					Return "0"
				End If
				If pattern_.ToLower = "true/false" Then
					Return "False"
				End If
				If pattern_.ToLower = "t/f" Then
					Return "False"
				End If
				If pattern_.ToLower = "t" Then
					Return "False"
				End If
				If pattern_.ToLower = "if/not" Then
					Return "False"
				End If
				If pattern_.ToLower = "i/n" Then
					Return "False"
				End If
				If pattern_.ToLower = "i" Then
					Return "False"
				End If
		End Select
	End Function

	''' <summary>
	''' Converts user-friendly word like yes or always to boolean i.e. true or false.
	''' To convert back to user-friendly word, use BooleanToDropText(boolean_val As Boolean, Optional pattern_ As String = "always/never")
	''' </summary>
	''' <example>
	''' <code>
	''' Dim userDecision as string = DropTextToBoolean(dropText.text)
	''' </code>
	''' </example>
	''' <param name="str_">always/never (default), yes/no, on/off, 1/0, true/false, if/not</param>
	''' <returns>True or False</returns>
	Public Shared Function DropTextToBoolean(str_ As String) As Boolean

		Select Case str_.ToString.Trim.ToLower
			Case ""
				Return False

			Case "yes"
				Return True
			Case "no"
				Return False


			Case "always"
				Return True
			Case "never"
				Return False


			Case "on"
				Return True
			Case "off"
				Return False


			Case "1"
				Return True
			Case "0"
				Return False


			Case "true"
				Return True
			Case "false"
				Return False


			Case "if"
				Return True
			Case "not"
				Return False
		End Select
	End Function

	''' <summary>
	''' Populates drop with drop-text version of boolean, depending on the pattern. Same as PopulateBooleanDrop.
	''' </summary>
	''' <param name="d_">ComboBox to fill</param>
	''' <param name="firstItemIsEmpty">Should 's first item be empty?</param>
	''' <param name="pattern_">always/never (default) OR a/n OR a, yes/no OR y/n OR y, on/off OR o/f OR o, 1/0, true/false OR t/f OR t, if/not OR i/n OR i</param>
	Public Shared Sub BooleanDrop(d_ As ComboBox, Optional pattern_ As String = "always/never", Optional firstItemIsEmpty As Boolean = False)
		Dim f_ As New D
		f_.PopulateBooleanDrop(d_, pattern_, firstItemIsEmpty)

		'		Dim w As New FormatWindowWeb
		'		With d_
		'			If .Items.Count > 0 Then Exit Sub
		'			With .Items
		'			If firstItemIsEmpty = True Then .Add("")
		'			.Add(w.BooleanToDropText(True, pattern_))
		'			.Add(w.BooleanToDropText(False, pattern_))
		'		End With
		'		End With
	End Sub

	''' <summary>
	''' Populates drop with drop-text version of boolean, depending on the pattern.
	''' </summary>
	''' <param name="d_">ComboBox to fill</param>
	''' <param name="firstItemIsEmpty">Should 's first item be empty?</param>
	''' <param name="pattern_">always/never (default) OR a/n OR a, yes/no OR y/n OR y, on/off OR o/f OR o, 1/0, true/false OR t/f OR t, if/not OR i/n OR i</param>
	Public Sub PopulateBooleanDrop(d_ As ComboBox, Optional pattern_ As String = "always/never", Optional firstItemIsEmpty As Boolean = False)
		Dim w As New NModule.W
		With d_
			If .Items.Count > 0 Then Exit Sub
			With .Items
				If firstItemIsEmpty = True Then .Add("")
				.Add(W.BooleanToDropText(True, pattern_))
				.Add(W.BooleanToDropText(False, pattern_))
			End With
			Try
				.Text = ""
				.SelectedIndex = -1
			Catch ex As Exception

			End Try
		End With
	End Sub

	Public Sub PopulateBEDrop(d_ As ComboBox, Optional firstItemIsEmpty As Boolean = False)
		If d_.Items.Count > 0 Then Exit Sub

		With d_.Items
			Try
				If firstItemIsEmpty = True And d_.DataSource Is Nothing Then .Add("")
			Catch
			End Try
			.Add("Begin")
			.Add("End")
		End With
	End Sub

	Public Function Toboolean(str_)
		Convert.ToBoolean(Convert.ToInt32(str_))
	End Function

#Region "JSON"
	''' <summary>
	''' Converts the items of ListBox to JSON. You can save the output directly to the database. Not intended to be used as dictionary. Opposite of DatabaseToControlJSON.
	''' </summary>
	''' <param name="l">ListBox with items</param>
	''' <returns>String to save to database</returns>
	Public Shared Function ControlToDatabaseJSON(l As ListBox) As String
		Dim val_ As String = ""
		Dim ltemp As ListBox = l
		If ltemp.Items.Count > 0 Then
			With ltemp
				.SelectedIndex = 0
				For k As Integer = 0 To .Items.Count - 1
					val_ &= "'" & .SelectedItem & "':'" & .SelectedItem & "'"
					'
					If k <> .Items.Count - 1 Then
						val_ &= ", "
						.SelectedIndex += 1
					End If
				Next
			End With
		End If
		Return val_
	End Function

	''' <summary>
	''' Takes database JSON string and populates control (ListBox, ComboBox or Textbox) with it. Opposite of ControlToDatabaseJSON. You can use NFunctions.DatabaseToListJSON() instead if you want it as a list.
	''' </summary>
	''' <param name="val_">Database JSON string</param>
	''' <param name="control_">Control to bind to</param>
	''' <param name="step_">Should the values be treated strictly as dictionary?</param>

	Public Shared Sub DatabaseToControlJSON(val_ As String, Optional control_ As Control = Nothing, Optional step_ As Integer = 1)
		If val_.Length < 1 Then Exit Sub

		Clear(control_)

		Dim s As New List(Of String)
		Try
			s.Clear()
		Catch ex As Exception
		End Try
		s = DatabaseToListJSON(val_.Trim)

		Dim c As ComboBox
		Dim l_ As ListBox
		Dim t As TextBox

		If TypeOf control_ Is ComboBox Then
			c = control_
			c.DataSource = s
			Exit Sub
		ElseIf TypeOf control_ Is ListBox Then
			l_ = control_
			l_.DataSource = s
			Exit Sub
		End If

		If TypeOf control_ Is TextBox Then
			t = control_
			t.Text = ""
			For i As Integer = 0 To s.Count - 1 Step step_
				t.Text &= s(i).ToString & vbCrLf
			Next
		End If
	End Sub

#End Region
End Class
