' Create_SAR_Excitations
' Creates all the excitations AC tasks necessary for simulation
' WARNING: Magnitude and phase for the coils are set to 1 and 0, you can change those values but, ultimately,
' we recommend running the Set_phase_and_magnitude macro

'Custom parameters
'nb_coils: Number of coils
'port_start: Index of the first port for your coils (Note: your ports should be indexed sequentially)
'default_phase: Phase inputted as default when creating the excitation tasks
'default_magnitude: Magnitude inputted as default when creating the excitations tasks

Sub Main ()
	Dim nb_coils As Integer
	Dim port_start As Integer
	Dim default_magnitude As Double
	Dim default_phase As Double

	'########################################INSERT PARAMETERS HERE#################################################'
	nb_coils = 8
	port_start = 24
	default_phase = 0
	default_magnitude = 1
	'###########################################################################################################'

	With SimulationTask
		.name("SAR_excitations")
		.type("sequence")
		.create

		.SetProperty ("enabled", "True")
	End With

	For i = 1 To nb_coils
		For j = i To nb_coils
			With SimulationTask
				.name("SAR_excitations\Exc_" & i & "_" & j)
				.type("ac")
				.create

				.SetProperty ("maximum frequency range", "True")
				.SetProperty ("enabled", "True")
				.EnableResult ("block", True )
				.SetComplexPortExcitation(port_start - 1 + i, default_magnitude, default_phase)
				.SetComplexPortExcitation(port_start - 1 + j, default_magnitude, default_phase)
				.SetPortSourceType (port_start - 1 + i, "signal")
				.SetPortSourceType (port_start - 1 + j, "signal")
			End With

			If i <> j Then
				With SimulationTask
					.name("SAR_excitations\Exc_" & i & "_" & j & "_prime")
					.type("ac")
					.create

					.SetProperty ("maximum frequency range", "True")
					.SetProperty ("enabled", "True")
					.EnableResult ("block", True )
					.SetComplexPortExcitation(port_start - 1 + i, default_magnitude, default_phase)
					.SetComplexPortExcitation(port_start - 1 + j, default_magnitude, default_phase)
					.SetPortSourceType (port_start - 1 + i, "signal" )
					.SetPortSourceType (port_start - 1 + j, "signal" )
				End With
			End If
		Next j
	Next i

End Sub
