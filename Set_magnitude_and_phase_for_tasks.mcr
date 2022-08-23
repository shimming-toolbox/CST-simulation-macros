' Macro used to set magnitude and phase for all tasks.'
' The magnitudes and the phases need to be calculated beforehand'

Sub Main ()
	Dim Magnitudes As Variant
	Dim Phases As Variant
	Dim nb_coils As Integer
	Dim Port_offset As Integer

	'########################################INSERT PARAMETERS HERE#################################################'
	nb_coils = 1
	Port_start = 1
	Magnitudes = Array(1)
	Phases = Array(1)
	'###########################################################################################################'

	'#########Select the first Coil location######'
	SimulationTask.StartTaskNameIteration
	Dim Name As String
	Name = SimulationTask.GetNextTaskName
	Name = SimulationTask.GetNextTaskName

	'######### Iterate through every coil to set magnitude and phase ######'
	For Port = Port_start To Port_start + nb_coils - 1
		Name = SimulationTask.GetNextTaskName
		With SimulationTask
			.name(Name)
			.SetComplexPortExcitation(Port, Magnitudes(Port-Port_start), Phases(Port-Port_start))
		End With
	Next Port

	'######### Set CPmode values ######'
	Name = SimulationTask.GetNextTaskName
	For Port = Port_start To Port_start + nb_coils - 1
		With SimulationTask
			.name(Name)
			.SetComplexPortExcitation(Port, Magnitudes(Port-Port_start), Phases(Port-Port_start))
		End With
	Next Port

	'######### Set negCPmode values ######'
	Name = SimulationTask.GetNextTaskName
	For Port = Port_start To Port_start + nb_coils - 1
		With SimulationTask
			.name(Name)
			.SetComplexPortExcitation(Port, Magnitudes(Port-Port_start), -Phases(Port-Port_start))
		End With
	Next Port

	'######### Set zeroPhase values ######'
	Name = SimulationTask.GetNextTaskName
	For Port = Port_start To Port_start + nb_coils - 1
		With SimulationTask
			.name(Name)
			.SetComplexPortExcitation(Port, Magnitudes(Port-Port_start), 0)
		End With
	Next Port

	'######### Iterate through every excitation to set magnitude and phase ######'
	Name = SimulationTask.GetNextTaskName
	For Excitation = 1 To nb_coils * nb_coils
		Name = SimulationTask.GetNextTaskName
		Dim Port1 As Integer
		Dim Port2 As Integer
		Dim index As Integer

		For i = 1 To nb_coils
			index = FindPosition(Name, "_" & CStr(i))
			If index > 0 Then
				Port1 = Port_start + i - 1
				Port2 = Port_start + CInt(Mid(Name, index+2, 1)) - 1

				With SimulationTask
					.name(Name)
					.SetComplexPortExcitation(Port1, Magnitudes(Port1-Port_start), Phases(Port1-Port_start))
					If FindPosition(Name, "prime") > 0 Then
						.SetComplexPortExcitation(Port2, Magnitudes(Port2-Port_start), 90-Phases(Port2-Port_start))
					Else
						.SetComplexPortExcitation(Port2, Magnitudes(Port2-Port_start), Phases(Port2-Port_start))
					End If
				End With
				Exit For
			End If
		Next i
	Next Excitation

End Sub

'Function to find a substring in a string'
' Inputs'
'	Ref: Reference string to look into'
'	find: substring to find in the main string'
'Returns:
'	the index of the start of the substring'
Function FindPosition(Ref As String, find As String) As Integer
	Dim Position As Integer
	Position = InStr(Ref, find)
	FindPosition = Position
End Function
