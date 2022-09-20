' Postprocessing_ExportSAR
' This Macro calculates exports SAR for all your SAR_excitations

'Custom parameters
'nb_coils: Number of coils
'xmin, xmax, ymin, ymax, zmin, zmax: subvolume for SAR export
'stepwidth: step width for data export (in units of project)
'export_directory: directory for exported SAR files (directory must already exist)

Sub Main ()
	Dim nb_coils As Integer
	Dim xmin As Double
	Dim xmax As Double
	Dim ymin As Double
	Dim ymax As Double
	Dim zmin As Double
	Dim zmax As Double
	Dim stepwidth As Double
	Dim export_directory As String


	'####################INSERT PARAMETERS HERE###########################'
	nb_coils = 1

	xmin = -9.5
	xmax = 9.5
	ymin = -4.5
	ymax = 6.5
	zmin = -11.5
	zmax = 4

	stepwidth = 0.1

	export_directory = "E:\20220906_VisualCoil_Gustav_mask"

	'#####################################################################'

	For i = 1 To nb_coils
		For j = i To nb_coils
		SelectTreeItem ("2D/3D Results\SAR\SAR (f=297.2) [SAR_excitations@Exc_" & i & "_" & j & "] (10g)")
			With ASCIIExport
    			.Reset
    			.FileName (export_directory & "\SAR (f=297.2) [SAR_excitations@Exc_" & i & "_" & j & "] (10g).txt")
    			.Mode("FixedWidth")
    			.Step (stepwidth)
    			.SetSubvolume(xmin, xmax, ymin, ymax, zmin, zmax)
    			.UseSubvolume("True")
   				.Execute
			End With

		If i <> j Then
				SelectTreeItem ("2D/3D Results\SAR\SAR (f=297.2) [SAR_excitations@Exc_" & i & "_" & j & "_prime] (10g)")
			With ASCIIExport
    			.Reset
    			.FileName (export_directory & "\SAR (f=297.2) [SAR_excitations@Exc_" & i & "_" & j & "_prime] (10g).txt")
    			.Mode("FixedWidth")
    			.Step (stepwidth)
    			.SetSubvolume(xmin, xmax, ymin, ymax, zmin, zmax)
   				.Execute
   			End With

		End If
		Next j
	Next i

End Sub

