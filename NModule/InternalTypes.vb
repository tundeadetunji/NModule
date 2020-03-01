Public Class InternalTypes
#Region "General_Module"
	Enum ActionItems
		DoNothing
		Hibernate
		LogOff
		OpenFile
		OpenApp
		Shutdown
		Sleep
	End Enum
#End Region
#Region "Web_Module"

	Enum BooleanToDropText
		AlwaysNever
		YesNo
		OnOff
		OneZero
		TrueFalse
		IfNot
	End Enum

	Enum Alert As Integer
		alert
		danger
		success
		warning
	End Enum

	Enum Gender
		Male
		Female
		Trans
	End Enum
	Enum Title
		Mr
		Mrs
		Ms
	End Enum

#End Region

End Class
