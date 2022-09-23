' Postprocessing_CalculateSAR
' This Macro calculates SAR for all your SAR_excitations.
' You can also add extra tasks like CP mode with the Extra_Tasks parameters

'Custom parameters
'nb_coils: Number of coils
'Extra_Tasks: Extra tasks other than coils you want to add

Sub Main ()
	Dim nb_coils As Integer


	'####################INSERT PARAMETERS HERE###########################'
	nb_coils = 1
	Extra_Tasks = Array("CPmode", "negCPMode", "zeroPhase")

	'#####################################################################'


	For i = 1 To nb_coils
		For j = i To nb_coils
			With SAR
				.Reset
				.PowerlossMonitor ("loss (f=297.2) [SAR_excitations@Exc_" & i & "_" & j & "]")
				.AverageWeight (10)
				.SetOption ("no subvolume")
				.SetOption ("rescale 1")
				.SetOption ("scale stimulated")
				.SetOption ("CST C95.3")
				.Create
			End With

		If i <> j Then
			With SAR
				.Reset
				.PowerlossMonitor ("loss (f=297.2) [SAR_excitations@Exc_" & i & "_" & j & "_prime]")
				.AverageWeight (10)
				.SetOption ("no subvolume")
				.SetOption ("rescale 1")
				.SetOption ("scale stimulated")
				.SetOption ("CST C95.3")
				.Create
				End With
		End If
		Next j
	Next i

	For Each task In Extra_Tasks
		With SAR
				.Reset
				.PowerlossMonitor ("loss (f=297.2) [CoilCombinations@" & task & "]")
				.AverageWeight (10)
				.SetOption ("no subvolume")
				.SetOption ("rescale 1")
				.SetOption ("scale stimulated")
				.SetOption ("CST C95.3")
				.Create
		End With
	Next task

End Sub

