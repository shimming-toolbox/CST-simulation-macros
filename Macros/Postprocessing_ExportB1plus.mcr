' Postprocessing_ExportSAR
' This Macro calculates exports B1 plus for all your Coil Combinations and Extra Tasks

'Custom parameters
'nb_coils: Number of coils
'Extra_Tasks: Extra tasks other than coils you want to export
'xmin, xmax, ymin, ymax, zmin, zmax: subvolume for B1 plus export
'stepwidth: step width for data export (in units of project)
'export_directory: directory for exported B1-plus files (directory must already exist)

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
	Extra_Tasks = Array("CPmode", "negCPMode", "zeroPhase")

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
		SelectTreeItem ("2D/3D Results\B1+ and B1-\B1+ (f=297.2) [CoilCombinations@Coil_1] 0.5 W stim")
			With ASCIIExport
    			.Reset
    			.FileName (export_directory & "\B1+ (f=297.2) [CoilCombinations@Coil_" & i & "] 0.5 W stim.txt")
    			.Mode("FixedWidth")
    			.Step (stepwidth)
    			.SetSubvolume(xmin, xmax, ymin, ymax, zmin, zmax)
    			.UseSubvolume("True")
   				.Execute
			End With
	Next i

	For Each task In Extra_Tasks
			SelectTreeItem ("2D/3D Results\B1+ and B1-\B1+ (f=297.2) [CoilCombinations@" & task & "] 0.5 W stim.txt")
		With ASCIIExport
    		.Reset
    		.FileName (export_directory & "\B1+ (f=297.2) [CoilCombinations@" & task & "] 0.5 W stim.txt")
    		.Mode("FixedWidth")
    		.Step (stepwidth)
    		.SetSubvolume(xmin, xmax, ymin, ymax, zmin, zmax)
   			.Execute
   		End With
	Next task

End Sub

