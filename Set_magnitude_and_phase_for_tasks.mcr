' Test_macro

Sub Main ()
	Dim Magnitudes As Variant
	Dim nb_coils As Integer
	Dim Port_offset As Integer

	'########################################INSERT PARAMETERS HERE#################################################'
	nb_coils = 8
	Port_start = 24
	Magnitude = Array(1.94862190712210,	1.10669122909554, 1.26596992010020,	1.32685228907074,	0.971669026661701,	0.304070905426335,	1.21250404173537,	2.53804993344609)
	Phase = Array(-179.994495794148,	-170.202367386334,	58.6623521594078,	50.8973684142317,	43.4973150663814,	265.364891892856,	-97.1287619423240,	25.1440897700117)
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
			.SetComplexPortExcitation(Port, Magnitude(Port-Port_start), Phase(Port-Port_start))
		End With
	Next Port

	'######### Set CPmode values ######'
	Name = SimulationTask.GetNextTaskName
	For Port = Port_start To Port_start + nb_coils - 1
		With SimulationTask
			.name(Name)
			.SetComplexPortExcitation(Port, Magnitude(Port-Port_start), Phase(Port-Port_start))
		End With
	Next Port

	'######### Set negCPmode values ######'
	Name = SimulationTask.GetNextTaskName
	For Port = Port_start To Port_start + nb_coils - 1
		With SimulationTask
			.name(Name)
			.SetComplexPortExcitation(Port, Magnitude(Port-Port_start), -Phase(Port-Port_start))
		End With
	Next Port

	'######### Set zeroPhase values ######'
	Name = SimulationTask.GetNextTaskName
	For Port = Port_start To Port_start + nb_coils - 1
		With SimulationTask
			.name(Name)
			.SetComplexPortExcitation(Port, Magnitude(Port-Port_start), 0)
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
					.SetComplexPortExcitation(Port1, Magnitude(Port1-Port_start), Phase(Port1-Port_start))
					If FindPosition(Name, "prime") > 0 Then
						.SetComplexPortExcitation(Port2, Magnitude(Port2-Port_start), 90-Phase(Port2-Port_start))
					Else
						.SetComplexPortExcitation(Port2, Magnitude(Port2-Port_start), Phase(Port2-Port_start))
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
