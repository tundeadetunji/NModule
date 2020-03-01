Imports System.Web
Imports System.Web.UI.WebControls
Public Class W
	''' <summary>
	''' Binds Tasks status to DropDownList. You can adjust the list in Web_Module.Methods.StatusList.
	''' </summary>
	''' <param name="d_">The DropDownList to bind to.</param>
	''' <param name="FirstIsEmpty">Should the first item appear empty? Default is True.</param>

	Public Shared Sub StatusDrop(d_ As DropDownList, Optional FirstIsEmpty As Boolean = True)
		If d_.Items.Count > 0 Then Exit Sub

		Dim web_methods_ As New Web_Module.Methods
		'		Dim d_w As New Web_Module.DataConnectionWeb

		Clear(d_)

		If FirstIsEmpty Then d_.Items.Add("")

		Dim l As List(Of String) = web_methods_.StatusList()
		For i As Integer = 0 To l.Count - 1
			d_.Items.Add(l(i).ToString)
		Next
	End Sub
	''' <summary>
	''' Binds Reminder recurrence intervals to DropDownList. You can adjust the list in Web_Module.Methods.RecurrenceList.
	''' </summary>
	''' <param name="d_">The DropDownList to bind to.</param>
	''' <param name="FirstIsEmpty">Should the first item appear empty? Default is True.</param>

	Public Shared Sub RecurrenceDrop(d_ As DropDownList, Optional FirstIsEmpty As Boolean = True)
		If d_.Items.Count > 0 Then Exit Sub

		Dim web_methods_ As New Web_Module.Methods
		'		Dim d_w As New Web_Module.DataConnectionWeb

		Clear(d_)

		If FirstIsEmpty Then d_.Items.Add("")

		Dim l As List(Of String) = web_methods_.RecurrenceList()
		For i As Integer = 0 To l.Count - 1
			d_.Items.Add(l(i).ToString)
		Next
	End Sub
	''' <summary>
	''' Binds Reminder intervals to DropDownList. You can adjust the list in Web_Module.Methods.ReminderList.
	''' </summary>
	''' <param name="d_">The DropDownList to bind to.</param>
	''' <param name="FirstIsEmpty">Should the first item appear empty? Default is True.</param>

	Public Shared Sub ReminderDrop(d_ As DropDownList, Optional FirstIsEmpty As Boolean = True)
		If d_.Items.Count > 0 Then Exit Sub

		Dim web_methods_ As New Web_Module.Methods
		'		Dim d_w As New Web_Module.DataConnectionWeb

		Clear(d_)

		If FirstIsEmpty Then d_.Items.Add("")

		Dim l As List(Of String) = web_methods_.ReminderList()
		For i As Integer = 0 To l.Count - 1
			d_.Items.Add(l(i).ToString)
		Next
	End Sub

	Public Shared Sub ReminderTypeDrop(d_ As DropDownList, Optional FirstIsEmpty As Boolean = True)
		If d_.Items.Count > 0 Then Exit Sub

		Dim web_methods_ As New Web_Module.Methods
		'		Dim d_w As New Web_Module.DataConnectionWeb

		Clear(d_)

		If FirstIsEmpty Then d_.Items.Add("")

		Dim l As List(Of String) = web_methods_.ReminderTypeList()
		For i As Integer = 0 To l.Count - 1
			d_.Items.Add(l(i).ToString)
		Next
	End Sub

	Public Shared Sub EnableControl(div_ As System.Web.UI.HtmlControls.HtmlGenericControl, Optional state_ As Boolean = True)
		div_.Visible = state_
	End Sub


	''' <summary>
	''' Populates DropDownList with numbers.
	''' </summary>
	''' <param name="d_">DropDownList to fill</param>
	''' <param name="end_">Ending Number</param>
	''' <param name="start_">Beginning Number</param>
	''' <param name="firstItemIsEmpty"></param>
	Public Shared Sub NumberDrop(d_ As DropDownList, end_ As Integer, Optional start_ As Integer = 0, Optional firstItemIsEmpty As Boolean = False)
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
	''' Attempts to clear control of text, image, items and DataSource.
	''' </summary>
	''' <param name="c_">Control to clear.</param>
	Public Shared Sub Clear(c_ As WebControl)

		Dim c As DropDownList
		Dim l As ListBox
		Dim t As TextBox
		Dim p As System.Web.UI.WebControls.Image
		Dim h As CheckBox
		If TypeOf c_ Is CheckBox Then
			h = c_
			h.Checked = False
		End If
		If TypeOf c_ Is DropDownList Then
			c = c_
			Try
				c.DataSource = Nothing
			Catch ex As Exception
			End Try
			c.Items.Clear()
			c.Text = ""
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
			t.Text = ""
		End If
		If TypeOf c_ Is System.Web.UI.WebControls.Image Then
			p = c_
			Try
				p.ImageUrl = Nothing
			Catch ex As Exception
			End Try
		End If
	End Sub
	''' <summary>
	''' Returns the text inside a control.
	''' </summary>
	''' <param name="control_">TextBox or DropDownList.</param>
	''' <returns></returns>

	Public Shared Function Content(control_ As WebControl) As String
		Try
			'If trim_ Then
			If TypeOf control_ Is TextBox Then
				Return CType(control_, TextBox).Text.Trim
			ElseIf TypeOf control_ Is DropDownList Then
				Return CType(control_, DropDownList).Text.Trim
			End If
			'Else
			'Return control_.Text
			'End If
		Catch
		End Try
	End Function

	''' <summary>
	''' Indicates if supported web control has content
	''' </summary>
	''' <param name="g_OR_d_OR_l">The web control (GridView or DropDownList or ListBox</param>
	''' <returns>True if control has content, False if not.</returns>
	Public Shared Function HasData(g_OR_d_OR_l) As Boolean

		Dim g_ As GridView
		Dim d_ As DropDownList
		Dim l_ As ListBox

		If g_OR_d_OR_l IsNot Nothing And TypeOf (g_OR_d_OR_l) Is GridView Then
			g_ = g_OR_d_OR_l
			Return g_.Rows.Count > 0
		ElseIf g_OR_d_OR_l IsNot Nothing And TypeOf (g_OR_d_OR_l) Is DropDownList Then
			d_ = g_OR_d_OR_l
			Return d_.Items.Count > 0
		ElseIf g_OR_d_OR_l IsNot Nothing And TypeOf (g_OR_d_OR_l) Is ListBox Then
			l_ = g_OR_d_OR_l
			Return l_.Items.Count > 0
		End If
	End Function
	''' <summary>
	''' Binds Titles Of Courtesy to DropDownList.
	''' </summary>
	''' <param name="c_">DropDownList to bind to.</param>
	''' <param name="first_item_is_empty">Should the first item appear empty? Default is True.</param>

	Public Shared Sub TitleDrop(c_ As DropDownList, Optional first_item_is_empty As Boolean = True)
		If c_.Items.Count > 0 Then Exit Sub

		Clear(c_)
		With c_
			With .Items
				If first_item_is_empty Then .Add("")
				.Add("Mr.")
				.Add("Mrs.")
				.Add("Ms.")
			End With
			.SelectedIndex = -1
			.Text = ""
		End With

	End Sub
	''' <summary>
	''' Binds Gender to DropDownList. You can adjust the list in Web_Module.Methods.GenderList.
	''' </summary>
	''' <param name="d_">DropDownList to bind to.</param>
	''' <param name="FirstIsEmpty">Should the first item appear empty? Default is True.</param>

	Public Shared Sub GenderDrop(d_ As DropDownList, Optional FirstIsEmpty As Boolean = True)
		If d_.Items.Count > 0 Then Exit Sub

		Dim web_methods_ As New Web_Module.Methods

		Clear(d_)

		If FirstIsEmpty Then d_.Items.Add("")

		Dim l As List(Of String) = web_methods_.GenderList()
		For i As Integer = 0 To l.Count - 1
			d_.Items.Add(l(i).ToString)
		Next
	End Sub
	''' <summary>
	''' Determines if all controls have text.
	''' </summary>
	''' <param name="controls_">Controls in array.</param>
	''' <returns>True or False</returns>

	Public Shared Function IsEmpty(controls_ As Array) As Boolean
		Dim counter_ As Integer = 0
		With controls_
			For i As Integer = 0 To .Length - 1
				If IsEmpty(controls_(i)) Then
					counter_ += 1
				End If
			Next
		End With
		Return Val(counter_) = controls_.Length
	End Function

	''' <summary>
	''' Determines if control has text.
	''' </summary>
	''' <param name="c_">DropDownList or Textbox</param>
	''' <param name="use_trim">Should content be trimmed before check?</param>
	''' <returns></returns>
	Public Shared Function IsEmpty(c_ As WebControl, Optional use_trim As Boolean = True) As Boolean
		Dim d__ As DropDownList
		Dim t__ As TextBox

		If TypeOf c_ Is DropDownList Then
			d__ = c_
			If use_trim = True Then
				Return d__.Text.Trim.Length < 1
			Else
				Return d__.Text.Length < 1
			End If
		End If

		If TypeOf c_ Is TextBox Then
			t__ = c_
			If use_trim = True Then
				Return t__.Text.Trim.Length < 1
			Else
				Return t__.Text.Length < 1
			End If
		End If

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
	''' Populates drop with drop-text version of boolean, depending on the pattern.
	''' </summary>
	''' <param name="d_">DropDownList</param>
	''' <param name="firstItemIsEmpty">Should DropDownList's first item be empty?</param>
	''' <param name="pattern_">always/never (default) OR a/n OR a, yes/no OR y/n OR y, on/off OR o/f OR o, 1/0, true/false OR t/f OR t, if/not OR i/n OR i</param>
	Public Shared Sub BooleanDrop(d_ As DropDownList, Optional pattern_ As String = "always/never", Optional firstItemIsEmpty As Boolean = True)
		With d_
			If .Items.Count > 0 Then Exit Sub
			With .Items
				If firstItemIsEmpty = True Then .Add("")
				.Add(BooleanToDropText(True, pattern_))
				.Add(BooleanToDropText(False, pattern_))
			End With
		End With
	End Sub


	''' <summary>
	''' Sets the title of the page.
	''' </summary>
	''' <param name="regular_str">The title, if condition_ exists</param>
	''' <param name="alternate_str">The title, if condition_ does not exist</param>
	''' <param name="condition_">Determines if alternate_ should be used in place of regular_</param>
	''' <returns></returns>
	Public Shared Function PageTitle(regular_str As String, alternate_str As String, condition_ As String) As String
		If condition_ IsNot Nothing Then
			Return regular_str
		Else
			Return alternate_str
		End If
	End Function

End Class
