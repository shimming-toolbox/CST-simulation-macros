' Create_CoilCombinations
' This Macro creates all your AC tasks for your coils under the Coil_Combinations sequence
' It will set your magnitude and phase to what you chose down below, 1 magnitude and 0 phase being the default settings
' You can also add extra tasks like CP mode with the Extra_Tasks parameters

Sub Main ()
	Dim nb_coils As Integer
	Dim port_start As Integer
	Dim phantom_magnitude As Double
	Dim phantom_phase As Double

	'########################################INSERT PARAMETERS HERE#################################################'
	nb_coils = 8
	port_start = 24
	Extra_Tasks = Array("test_CPmode", "test_negCPMode", "test_zeroPhase")

	'Default value for magnitude and phase. Used to simulate on the phantom'
	phantom_magnitude = 1
	phantom_phase = 0
	'###########################################################################################################'

	With SimulationTask
		.name("Coil_Combinations")
		.type("sequence")
		.SetProperty ("enabled", "False")
		.create
	End With

	For i = 1 To nb_coils
		With SimulationTask
			.name("Coil_Combinations\Coil_" & i)
			.type("ac")
			.create

			.SetProperty ("maximum frequency range", "True")
			.SetProperty ("enabled", "False")
			.EnableResult ("block", True )
			.SetComplexPortExcitation(port_start - 1 + i, phantom_magnitude, phantom_phase)
			.SetPortSourceType (port_start - 1 + i, "signal")
		End With
	Next i

	For Each task In Extra_Tasks
		With SimulationTask
			.name("Coil_Combinations\" & task)
			.type("ac")
			.create

			.SetProperty ("enabled", "False")
			.EnableResult ("block", True )
			For Port = port_start To port_start + nb_coils - 1
				.SetComplexPortExcitation(Port, phantom_magnitude, phantom_phase)
				.SetPortSourceType (Port , "signal" )
			Next Port
		End With
	Next task

End Sub
