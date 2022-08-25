' Macro used to set magnitude and phase for all tasks.'
' The magnitudes and the phases need to be calculated beforehand'

'Custom parameters
'nb_coils: Number of coils
'port_start: Index of the first port for your coils (Note: your ports should be indexed sequentially)
'Magnitudes: List of magnitudes for every coil
'Phases: List of phases for every coil
'Extra_Tasks: Extra tasks other than coils you want to add. Make sure to include zero or neg if you want the phases to be altered subsequently

Sub Main ()
	Dim Magnitudes As Variant
	Dim Phases As Variant
	Dim nb_coils As Integer
	Dim Port_offset As Integer

	'########################################INSERT PARAMETERS HERE#################################################'
	nb_coils = 8
	port_start = 24
	Magnitudes = Array(1, 2, 3, 4, 5, 6, 7, 8)
	Phases = Array(10, -20, 30, -40, 50, -60, 70, -80)
	Extra_Tasks = Array("CPmode", "negCPMode", "zeroPhase")
	'###########################################################################################################'

	'#########Loop through every coil######'
	For Coil = 1 To nb_coils
		With SimulationTask
			.name("Coil_Combinations\Coil_" & Coil)
			.SetComplexPortExcitation(port_start - 1 + Coil, Magnitudes(Coil-1), Phases(Coil-1))
		End With
			For Coil2 = Coil To nb_coils
				With SimulationTask
					.name("SAR_excitations\Exc_" & Coil & "_" & Coil2)
					.SetComplexPortExcitation(port_start - 1 + Coil, Magnitudes(Coil-1), Phases(Coil-1))
					.SetComplexPortExcitation(port_start - 1 + Coil2, Magnitudes(Coil2-1), Phases(Coil2-1))
				End With

				If Coil <> Coil2 Then
					With SimulationTask
						.name("SAR_excitations\Exc_" & Coil & "_" & Coil2 & "_prime")
						.SetComplexPortExcitation(port_start - 1 + Coil, Magnitudes(Coil-1), Phases(Coil-1))
						.SetComplexPortExcitation(port_start - 1 + Coil2, Magnitudes(Coil2-1), 90 + Phases(Coil2-1))
					End With
				End If
		Next Coil2
	Next Coil

	For Each task In Extra_Tasks
		Dim phase_sign As Integer
		phase_sign = 1

		If FindPosition(task, "zero") > 0 Then
			phase_sign = 0
		End If

		If FindPosition(task, "neg") > 0 Then
			phase_sign = -1
		End If

		With SimulationTask
			.name("Coil_Combinations\" & task)
			For Port = port_start To port_start + nb_coils -1
				.SetComplexPortExcitation(Port, Magnitudes(Port - port_start), phase_sign * Phases(Port - port_start))
			Next Port
		End With
	Next task
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
