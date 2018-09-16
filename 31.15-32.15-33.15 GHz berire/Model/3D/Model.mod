'# MWS Version: Version 2014.0 - Feb 24 2014 - ACIS 23.0.0 -

'# length = mm
'# frequency = GHz
'# time = ns
'# frequency range: fmin = 20 fmax = 40
'# created = '[VERSION]2014.0|23.0.0|20140224[/VERSION]


'@ use template: Antenna - Planar_19

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
'set the units
With Units
    .Geometry "mm"
    .Frequency "GHz"
    .Voltage "V"
    .Resistance "Ohm"
    .Inductance "NanoH"
    .TemperatureUnit  "Kelvin"
    .Time "ns"
    .Current "A"
    .Conductance "Siemens"
    .Capacitance "PikoF"
End With
'----------------------------------------------------------------------------
Plot.DrawBox True
With Background
     .Type "Normal"
     .Epsilon "1.0"
     .Mue "1.0"
     .XminSpace "0.0"
     .XmaxSpace "0.0"
     .YminSpace "0.0"
     .YmaxSpace "0.0"
     .ZminSpace "0.0"
     .ZmaxSpace "0.0"
End With
With Boundary
     .Xmin "expanded open"
     .Xmax "expanded open"
     .Ymin "expanded open"
     .Ymax "expanded open"
     .Zmin "expanded open"
     .Zmax "expanded open"
     .Xsymmetry "none"
     .Ysymmetry "none"
     .Zsymmetry "none"
End With
' optimize mesh settings for planar structures
With Mesh
     .MergeThinPECLayerFixpoints "True"
     .RatioLimit "20"
     .AutomeshRefineAtPecLines "True", "6"
     .FPBAAvoidNonRegUnite "True"
     .ConsiderSpaceForLowerMeshLimit "False"
     .MinimumStepNumber "5"
     .AnisotropicCurvatureRefinement "True"
     .AnisotropicCurvatureRefinementFSM "True"
End With
With MeshSettings
     .SetMeshType "Hex"
     .Set "RatioLimitGeometry", "20"
     .Set "EdgeRefinementOn", "1"
     .Set "EdgeRefinementRatio", "6"
End With
With MeshSettings
     .SetMeshType "HexTLM"
     .Set "RatioLimitGeometry", "20"
End With
With MeshSettings
     .SetMeshType "Tet"
     .Set "VolMeshGradation", "1.5"
     .Set "SrfMeshGradation", "1.5"
End With
' change mesh adaption scheme to energy
' 		(planar structures tend to store high energy
'     	 locally at edges rather than globally in volume)
MeshAdaption3D.SetAdaptionStrategy "Energy"
' switch on FD-TET setting for accurate farfields
FDSolver.ExtrudeOpenBC "True"
PostProcess1D.ActivateOperation "vswr", "true"
PostProcess1D.ActivateOperation "yz-matrices", "true"
'----------------------------------------------------------------------------
'set the frequency range
Solver.FrequencyRange "20", "40"
Dim sDefineAt As String
sDefineAt = "20;30;40"
Dim sDefineAtName As String
sDefineAtName = "20;30;40"
Dim sDefineAtToken As String
sDefineAtToken = "f="
Dim aFreq() As String
aFreq = Split(sDefineAt, ";")
Dim aNames() As String
aNames = Split(sDefineAtName, ";")
Dim nIndex As Integer
For nIndex = LBound(aFreq) To UBound(aFreq)
Dim zz_val As String
zz_val = aFreq (nIndex)
Dim zz_name As String
zz_name = sDefineAtToken & aNames (nIndex)
' Define E-Field Monitors
With Monitor
    .Reset
    .Name "e-field ("& zz_name &")"
    .Dimension "Volume"
    .Domain "Frequency"
    .FieldType "Efield"
    .Frequency zz_val
    .Create
End With
' Define H-Field Monitors
With Monitor
    .Reset
    .Name "h-field ("& zz_name &")"
    .Dimension "Volume"
    .Domain "Frequency"
    .FieldType "Hfield"
    .Frequency zz_val
    .Create
End With
' Define Farfield Monitors
With Monitor
    .Reset
    .Name "farfield ("& zz_name &")"
    .Domain "Frequency"
    .FieldType "Farfield"
    .Frequency zz_val
    .ExportFarfieldSource "False"
    .Create
End With
Next
'----------------------------------------------------------------------------
With MeshSettings
     .SetMeshType "Hex"
     .Set "Version", 1%
End With
With Mesh
     .MeshType "PBA"
End With
'set the solver type
ChangeSolverType("HF Time Domain")

'@ define material: Rogers RT5880 (loss free)

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Material
     .Reset
     .Name "Rogers RT5880 (loss free)"
     .Folder ""
.FrqType "all" 
.Type "Normal" 
.SetMaterialUnit "GHz", "mm"
.Epsilon "2.2" 
.Mue "1.0" 
.Kappa "0.0" 
.TanD "0.0" 
.TanDFreq "0.0" 
.TanDGiven "False" 
.TanDModel "ConstTanD" 
.KappaM "0.0" 
.TanDM "0.0" 
.TanDMFreq "0.0" 
.TanDMGiven "False" 
.TanDMModel "ConstKappa" 
.DispModelEps "None" 
.DispModelMue "None" 
.DispersiveFittingSchemeEps "General 1st" 
.DispersiveFittingSchemeMue "General 1st" 
.UseGeneralDispersionEps "False" 
.UseGeneralDispersionMue "False" 
.Rho "0.0" 
.ThermalType "Normal" 
.ThermalConductivity "0.20" 
.SetActiveMaterial "all" 
.Colour "0.75", "0.95", "0.85" 
.Wireframe "False" 
.Transparency "0" 
.Create
End With

'@ new component: component1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Component.New "component1"

'@ define brick: component1:solid1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Brick
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "Rogers RT5880 (loss free)" 
     .Xrange "-W/2", "W/2" 
     .Yrange "-L/2", "-L/2" 
     .Zrange "0", "thickness" 
     .Create
End With

'@ delete shape: component1:solid1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Solid.Delete "component1:solid1"

'@ define brick: component1:solid1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Brick
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "Rogers RT5880 (loss free)" 
     .Xrange "-W/2", "W/2" 
     .Yrange "0", "L" 
     .Zrange "-thickness", "0" 
     .Create
End With

'@ define material: Copper (annealed)

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Material
     .Reset
     .Name "Copper (annealed)"
     .Folder ""
.FrqType "static" 
.Type "Normal" 
.SetMaterialUnit "Hz", "mm" 
.Epsilon "1" 
.Mue "1.0" 
.Kappa "5.8e+007" 
.TanD "0.0" 
.TanDFreq "0.0" 
.TanDGiven "False" 
.TanDModel "ConstTanD" 
.KappaM "0" 
.TanDM "0.0" 
.TanDMFreq "0.0" 
.TanDMGiven "False" 
.TanDMModel "ConstTanD" 
.DispModelEps "None" 
.DispModelMue "None" 
.DispersiveFittingSchemeEps "1st Order" 
.DispersiveFittingSchemeMue "1st Order" 
.UseGeneralDispersionEps "False" 
.UseGeneralDispersionMue "False" 
.FrqType "all" 
.Type "Lossy metal" 
.SetMaterialUnit "GHz", "mm" 
.Mue "1.0" 
.Kappa "5.8e+007" 
.Rho "8930.0" 
.ThermalType "Normal" 
.ThermalConductivity "401.0" 
.HeatCapacity "0.39" 
.MetabolicRate "0" 
.BloodFlow "0" 
.VoxelConvection "0" 
.MechanicsType "Isotropic" 
.YoungsModulus "120" 
.PoissonsRatio "0.33" 
.ThermalExpansionRate "17" 
.Colour "1", "1", "0" 
.Wireframe "False" 
.Reflection "False" 
.Allowoutline "True" 
.Transparentoutline "False" 
.Transparency "0" 
.Create
End With

'@ define brick: component1:solid2

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "-W/2", "W/2" 
     .Yrange "0", "L" 
     .Zrange "-(thicknes+thickness)", "-thickness" 
     .Create
End With

'@ rename block: component1:solid1 to: component1:solid Rogers

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Solid.Rename "component1:solid1", "solid Rogers"

'@ rename block: component1:solid2 to: component1:solid Plaka

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Solid.Rename "component1:solid2", "solid Plaka"

'@ define brick: component1:solid1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Brick
     .Reset 
     .Name "solid1" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "-W1/2", "W1/2" 
     .Yrange "0", "L1" 
     .Zrange "0", "thicknes" 
     .Create
End With

'@ define brick: component1:solid2

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Brick
     .Reset 
     .Name "solid2" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "-W2/2", "W2/2" 
     .Yrange "L1", "L1+L2" 
     .Zrange "0", "thicknes" 
     .Create
End With

'@ define brick: component1:solid3

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Brick
     .Reset 
     .Name "solid3" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "-W3/2", "W3/2" 
     .Yrange "L1+L2", "L1+L2+L3" 
     .Zrange "0", "thicknes" 
     .Create
End With

'@ define brick: component1:solid4

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Brick
     .Reset 
     .Name "solid4" 
     .Component "component1" 
     .Material "Copper (annealed)" 
     .Xrange "-W4/2", "W4/2" 
     .Yrange "L1+L2+L3", "L1+L2+L3+L4" 
     .Zrange "0", "thicknes" 
     .Create
End With

'@ define material: Aluminum

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Material
     .Reset
     .Name "Aluminum"
     .Folder ""
.FrqType "static" 
.Type "Normal" 
.SetMaterialUnit "Hz", "mm" 
.Epsilon "1" 
.Mue "1.0" 
.Kappa "3.56e+007" 
.TanD "0.0" 
.TanDFreq "0.0" 
.TanDGiven "False" 
.TanDModel "ConstTanD" 
.KappaM "0" 
.TanDM "0.0" 
.TanDMFreq "0.0" 
.TanDMGiven "False" 
.TanDMModel "ConstTanD" 
.DispModelEps "None" 
.DispModelMue "None" 
.DispersiveFittingSchemeEps "General 1st" 
.DispersiveFittingSchemeMue "General 1st" 
.UseGeneralDispersionEps "False" 
.UseGeneralDispersionMue "False" 
.FrqType "all" 
.Type "Lossy metal" 
.SetMaterialUnit "GHz", "mm" 
.Mue "1.0" 
.Kappa "3.56e+007" 
.Rho "2700.0" 
.ThermalType "Normal" 
.ThermalConductivity "237.0" 
.HeatCapacity "0.9"
.MechanicsType "Isotropic"
.YoungsModulus "69"
.PoissonsRatio "0.33"
.ThermalExpansionRate "23"
.Colour "1", "1", "0" 
.Wireframe "False" 
.Transparency "0" 
.Create
End With

'@ define brick: component1:solid5

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Brick
     .Reset 
     .Name "solid5" 
     .Component "component1" 
     .Material "Aluminum" 
     .Xrange "-W5/2", "W5/2" 
     .Yrange "L1+L2+L3+L51", "L1+L2+L3+L52" 
     .Zrange "0", "thicknes" 
     .Create
End With

'@ boolean subtract shapes: component1:solid4, component1:solid5

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Solid 
     .Version 9
     .Subtract "component1:solid4", "component1:solid5" 
     .Version 1
End With

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.762*5.22", "0.762*5.22"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.762", "0.762*5.22"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ define time domain solver parameters

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.SetCreator "High Frequency" 
With Solver 
     .Method "Hexahedral"
     .CalculationType "TD-S"
     .StimulationPort "All"
     .StimulationMode "All"
     .SteadyStateLimit "-30.0"
     .MeshAdaption "False"
     .AutoNormImpedance "False"
     .NormingImpedance "50"
     .CalculateModesOnly "False"
     .SParaSymmetry "False"
     .StoreTDResultsInCache  "False"
     .FullDeembedding "False"
     .SuperimposePLWExcitation "False"
     .UseSensitivityAnalysis "False"
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ define monitor: e-field (f=27.5)

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Monitor 
     .Reset 
     .Name "e-field (f=27.5)" 
     .Dimension "Volume" 
     .Domain "Frequency" 
     .FieldType "Efield" 
     .Frequency "27.5" 
     .UseSubvolume "False" 
     .SetSubvolume  "-7.2259104833333",  "7.2259104833333",  "-2.4982704833333",  "16.498270483333",  "-3.6412704833333",  "6.8569104833333" 
     .Create 
End With

'@ define monitor: h-field (f=27.5)

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Monitor 
     .Reset 
     .Name "h-field (f=27.5)" 
     .Dimension "Volume" 
     .Domain "Frequency" 
     .FieldType "Hfield" 
     .Frequency "27.5" 
     .UseSubvolume "False" 
     .SetSubvolume  "-7.2259104833333",  "7.2259104833333",  "-2.4982704833333",  "16.498270483333",  "-3.6412704833333",  "6.8569104833333" 
     .Create 
End With

'@ define farfield monitor: farfield (f=27.5)

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Monitor 
     .Reset 
     .Name "farfield (f=27.5)" 
     .Domain "Frequency" 
     .FieldType "Farfield" 
     .Frequency "27.5" 
     .UseSubvolume "False" 
     .ExportFarfieldSource "False" 
     .SetSubvolume  "-7.2259104833333",  "7.2259104833333",  "-2.4982704833333",  "16.498270483333",  "-3.6412704833333",  "6.8569104833333" 
     .Create 
End With

'@ define farfield monitor: farfield (f=28)

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Monitor 
     .Reset 
     .Name "farfield (f=28)" 
     .Domain "Frequency" 
     .FieldType "Farfield" 
     .Frequency "28" 
     .UseSubvolume "False" 
     .ExportFarfieldSource "False" 
     .SetSubvolume  "-7.2259104833333",  "7.2259104833333",  "-2.4982704833333",  "16.498270483333",  "-3.6412704833333",  "6.8569104833333" 
     .Create 
End With

'@ define monitor: e-field (f=28)

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Monitor 
     .Reset 
     .Name "e-field (f=28)" 
     .Dimension "Volume" 
     .Domain "Frequency" 
     .FieldType "Efield" 
     .Frequency "28" 
     .UseSubvolume "False" 
     .SetSubvolume  "-7.2259104833333",  "7.2259104833333",  "-2.4982704833333",  "16.498270483333",  "-3.6412704833333",  "6.8569104833333" 
     .Create 
End With

'@ define monitor: h-field (f=28)

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Monitor 
     .Reset 
     .Name "h-field (f=28)" 
     .Dimension "Volume" 
     .Domain "Frequency" 
     .FieldType "Hfield" 
     .Frequency "28" 
     .UseSubvolume "False" 
     .SetSubvolume  "-7.2259104833333",  "7.2259104833333",  "-2.4982704833333",  "16.498270483333",  "-3.6412704833333",  "6.8569104833333" 
     .Create 
End With

'@ define monitor: h-field (f=29.5)

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Monitor 
     .Reset 
     .Name "h-field (f=29.5)" 
     .Dimension "Volume" 
     .Domain "Frequency" 
     .FieldType "Hfield" 
     .Frequency "29.5" 
     .UseSubvolume "False" 
     .SetSubvolume  "-7.2259104833333",  "7.2259104833333",  "-2.4982704833333",  "16.498270483333",  "-3.6412704833333",  "6.8569104833333" 
     .Create 
End With

'@ define monitor: e-field (f=29.5)

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Monitor 
     .Reset 
     .Name "e-field (f=29.5)" 
     .Dimension "Volume" 
     .Domain "Frequency" 
     .FieldType "Efield" 
     .Frequency "29.5" 
     .UseSubvolume "False" 
     .SetSubvolume  "-7.2259104833333",  "7.2259104833333",  "-2.4982704833333",  "16.498270483333",  "-3.6412704833333",  "6.8569104833333" 
     .Create 
End With

'@ define farfield monitor: farfield (f=29.5)

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Monitor 
     .Reset 
     .Name "farfield (f=29.5)" 
     .Domain "Frequency" 
     .FieldType "Farfield" 
     .Frequency "29.5" 
     .UseSubvolume "False" 
     .ExportFarfieldSource "False" 
     .SetSubvolume  "-7.2259104833333",  "7.2259104833333",  "-2.4982704833333",  "16.498270483333",  "-3.6412704833333",  "6.8569104833333" 
     .Create 
End With

'@ define time domain solver parameters

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.SetCreator "High Frequency" 
With Solver 
     .Method "Hexahedral"
     .CalculationType "TD-S"
     .StimulationPort "All"
     .StimulationMode "All"
     .SteadyStateLimit "-40"
     .MeshAdaption "False"
     .AutoNormImpedance "False"
     .NormingImpedance "50"
     .CalculateModesOnly "False"
     .SParaSymmetry "False"
     .StoreTDResultsInCache  "False"
     .FullDeembedding "False"
     .SuperimposePLWExcitation "False"
     .UseSensitivityAnalysis "False"
End With

'@ farfield plot options

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With FarfieldPlot 
     .Plottype "Polar" 
     .Vary "angle1" 
     .Theta "90" 
     .Phi "90" 
     .Step "1" 
     .Step2 "1" 
     .SetLockSteps "True" 
     .SetPlotRangeOnly "False" 
     .SetThetaStart "0" 
     .SetThetaEnd "180" 
     .SetPhiStart "0" 
     .SetPhiEnd "360" 
     .SetTheta360 "False" 
     .SymmetricRange "False" 
     .SetTimeDomainFF "False" 
     .SetFrequency "-1" 
     .SetTime "0" 
     .SetColorByValue "True" 
     .DrawStepLines "False" 
     .DrawIsoLongitudeLatitudeLines "False" 
     .ShowStructure "False" 
     .SetStructureTransparent "False" 
     .SetFarfieldTransparent "False" 
     .SetSpecials "enablepolarextralines" 
     .SetPlotMode "Directivity" 
     .Distance "1" 
     .UseFarfieldApproximation "True" 
     .SetScaleLinear "False" 
     .SetLogRange "40" 
     .SetLogNorm "0" 
     .DBUnit "0" 
     .EnableFixPlotMaximum "False" 
     .SetFixPlotMaximumValue "1" 
     .SetInverseAxialRatio "False" 
     .SetAxesType "user" 
     .SetAnntenaType "unknown" 
     .Phistart "1.000000e+000", "0.000000e+000", "0.000000e+000" 
     .Thetastart "0.000000e+000", "0.000000e+000", "1.000000e+000" 
     .PolarizationVector "0.000000e+000", "1.000000e+000", "0.000000e+000" 
     .SetCoordinateSystemType "spherical" 
     .SetAutomaticCoordinateSystem "True" 
     .SetPolarizationType "Linear" 
     .SlantAngle 0.000000e+000 
     .Origin "bbox" 
     .Userorigin "0.000000e+000", "0.000000e+000", "0.000000e+000" 
     .SetUserDecouplingPlane "False" 
     .UseDecouplingPlane "False" 
     .DecouplingPlaneAxis "X" 
     .DecouplingPlanePosition "0.000000e+000" 
     .LossyGround "False" 
     .GroundEpsilon "1" 
     .GroundKappa "0" 
     .EnablePhaseCenterCalculation "False" 
     .SetPhaseCenterAngularLimit "3.000000e+001" 
     .SetPhaseCenterComponent "boresight" 
     .SetPhaseCenterPlane "both" 
     .ShowPhaseCenter "True" 
     .StoreSettings
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:2

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "2"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.762*5.22", "0.762*5.22"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.762", "0.762*5.22"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ delete port: port2

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "2"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.762*5.22", "0.762*5.22"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.762", "0.762*5.22"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ farfield plot options

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With FarfieldPlot 
     .Plottype "3D" 
     .Vary "angle1" 
     .Theta "90" 
     .Phi "90" 
     .Step "5" 
     .Step2 "5" 
     .SetLockSteps "True" 
     .SetPlotRangeOnly "False" 
     .SetThetaStart "0" 
     .SetThetaEnd "180" 
     .SetPhiStart "0" 
     .SetPhiEnd "360" 
     .SetTheta360 "False" 
     .SymmetricRange "False" 
     .SetTimeDomainFF "False" 
     .SetFrequency "-1" 
     .SetTime "0" 
     .SetColorByValue "True" 
     .DrawStepLines "False" 
     .DrawIsoLongitudeLatitudeLines "False" 
     .ShowStructure "False" 
     .SetStructureTransparent "False" 
     .SetFarfieldTransparent "False" 
     .SetSpecials "enablepolarextralines" 
     .SetPlotMode "Directivity" 
     .Distance "1" 
     .UseFarfieldApproximation "True" 
     .SetScaleLinear "False" 
     .SetLogRange "40" 
     .SetLogNorm "0" 
     .DBUnit "0" 
     .EnableFixPlotMaximum "False" 
     .SetFixPlotMaximumValue "1" 
     .SetInverseAxialRatio "False" 
     .SetAxesType "user" 
     .SetAnntenaType "unknown" 
     .Phistart "1.000000e+000", "0.000000e+000", "0.000000e+000" 
     .Thetastart "0.000000e+000", "0.000000e+000", "1.000000e+000" 
     .PolarizationVector "0.000000e+000", "1.000000e+000", "0.000000e+000" 
     .SetCoordinateSystemType "spherical" 
     .SetAutomaticCoordinateSystem "True" 
     .SetPolarizationType "Linear" 
     .SlantAngle 0.000000e+000 
     .Origin "bbox" 
     .Userorigin "0.000000e+000", "0.000000e+000", "0.000000e+000" 
     .SetUserDecouplingPlane "False" 
     .UseDecouplingPlane "False" 
     .DecouplingPlaneAxis "X" 
     .DecouplingPlanePosition "0.000000e+000" 
     .LossyGround "False" 
     .GroundEpsilon "1" 
     .GroundKappa "0" 
     .EnablePhaseCenterCalculation "False" 
     .SetPhaseCenterAngularLimit "3.000000e+001" 
     .SetPhaseCenterComponent "boresight" 
     .SetPhaseCenterPlane "both" 
     .ShowPhaseCenter "True" 
     .StoreSettings
End With

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.381*4.34", "0.381*4.34"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.381", "0.381*4.34"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.762*5.22", "0.762*5.22"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.762", "0.762*5.22"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.762*5.22", "0.762*5.22"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.762", "0.762*5.22"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.762*5.22", "0.762*5.22"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.762", "0.762*5.22"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ farfield plot options

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With FarfieldPlot 
     .Plottype "Polar" 
     .Vary "angle1" 
     .Theta "90" 
     .Phi "90" 
     .Step "1" 
     .Step2 "1" 
     .SetLockSteps "True" 
     .SetPlotRangeOnly "False" 
     .SetThetaStart "0" 
     .SetThetaEnd "180" 
     .SetPhiStart "0" 
     .SetPhiEnd "360" 
     .SetTheta360 "False" 
     .SymmetricRange "False" 
     .SetTimeDomainFF "False" 
     .SetFrequency "-1" 
     .SetTime "0" 
     .SetColorByValue "True" 
     .DrawStepLines "False" 
     .DrawIsoLongitudeLatitudeLines "False" 
     .ShowStructure "False" 
     .SetStructureTransparent "False" 
     .SetFarfieldTransparent "False" 
     .SetSpecials "enablepolarextralines" 
     .SetPlotMode "Directivity" 
     .Distance "1" 
     .UseFarfieldApproximation "True" 
     .SetScaleLinear "False" 
     .SetLogRange "40" 
     .SetLogNorm "0" 
     .DBUnit "0" 
     .EnableFixPlotMaximum "False" 
     .SetFixPlotMaximumValue "1" 
     .SetInverseAxialRatio "False" 
     .SetAxesType "user" 
     .SetAnntenaType "unknown" 
     .Phistart "1.000000e+000", "0.000000e+000", "0.000000e+000" 
     .Thetastart "0.000000e+000", "0.000000e+000", "1.000000e+000" 
     .PolarizationVector "0.000000e+000", "1.000000e+000", "0.000000e+000" 
     .SetCoordinateSystemType "spherical" 
     .SetAutomaticCoordinateSystem "True" 
     .SetPolarizationType "Linear" 
     .SlantAngle 0.000000e+000 
     .Origin "bbox" 
     .Userorigin "0.000000e+000", "0.000000e+000", "0.000000e+000" 
     .SetUserDecouplingPlane "False" 
     .UseDecouplingPlane "False" 
     .DecouplingPlaneAxis "X" 
     .DecouplingPlanePosition "0.000000e+000" 
     .LossyGround "False" 
     .GroundEpsilon "1" 
     .GroundKappa "0" 
     .EnablePhaseCenterCalculation "False" 
     .SetPhaseCenterAngularLimit "3.000000e+001" 
     .SetPhaseCenterComponent "boresight" 
     .SetPhaseCenterPlane "both" 
     .ShowPhaseCenter "True" 
     .StoreSettings
End With

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.762*5.22", "0.762*5.22"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.762", "0.762*5.22"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.762*5.22", "0.762*5.22"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.762", "0.762*5.22"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.762*5.22", "0.762*5.22"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.762", "0.762*5.22"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.762*5.22", "0.762*5.22"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.762", "0.762*5.22"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.762*5.22", "0.762*5.22"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.762", "0.762*5.22"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ farfield plot options

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With FarfieldPlot 
     .Plottype "3D" 
     .Vary "angle1" 
     .Theta "90" 
     .Phi "90" 
     .Step "5" 
     .Step2 "5" 
     .SetLockSteps "True" 
     .SetPlotRangeOnly "False" 
     .SetThetaStart "0" 
     .SetThetaEnd "180" 
     .SetPhiStart "0" 
     .SetPhiEnd "360" 
     .SetTheta360 "False" 
     .SymmetricRange "False" 
     .SetTimeDomainFF "False" 
     .SetFrequency "-1" 
     .SetTime "0" 
     .SetColorByValue "True" 
     .DrawStepLines "False" 
     .DrawIsoLongitudeLatitudeLines "False" 
     .ShowStructure "False" 
     .SetStructureTransparent "False" 
     .SetFarfieldTransparent "False" 
     .SetSpecials "enablepolarextralines" 
     .SetPlotMode "Directivity" 
     .Distance "1" 
     .UseFarfieldApproximation "True" 
     .SetScaleLinear "False" 
     .SetLogRange "40" 
     .SetLogNorm "0" 
     .DBUnit "0" 
     .EnableFixPlotMaximum "False" 
     .SetFixPlotMaximumValue "1" 
     .SetInverseAxialRatio "False" 
     .SetAxesType "user" 
     .SetAnntenaType "unknown" 
     .Phistart "1.000000e+000", "0.000000e+000", "0.000000e+000" 
     .Thetastart "0.000000e+000", "0.000000e+000", "1.000000e+000" 
     .PolarizationVector "0.000000e+000", "1.000000e+000", "0.000000e+000" 
     .SetCoordinateSystemType "spherical" 
     .SetAutomaticCoordinateSystem "True" 
     .SetPolarizationType "Linear" 
     .SlantAngle 0.000000e+000 
     .Origin "bbox" 
     .Userorigin "0.000000e+000", "0.000000e+000", "0.000000e+000" 
     .SetUserDecouplingPlane "False" 
     .UseDecouplingPlane "False" 
     .DecouplingPlaneAxis "X" 
     .DecouplingPlanePosition "0.000000e+000" 
     .LossyGround "False" 
     .GroundEpsilon "1" 
     .GroundKappa "0" 
     .EnablePhaseCenterCalculation "False" 
     .SetPhaseCenterAngularLimit "3.000000e+001" 
     .SetPhaseCenterComponent "boresight" 
     .SetPhaseCenterPlane "both" 
     .ShowPhaseCenter "True" 
     .StoreSettings
End With

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.762*5.22", "0.762*5.22"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.762", "0.762*5.22"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete monitors

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Monitor.Delete "e-field (f=20)" 
Monitor.Delete "farfield (f=20)" 
Monitor.Delete "h-field (f=20)"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.762*5.22", "0.762*5.22"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.762", "0.762*5.22"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.762*5.22", "0.762*5.22"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.762", "0.762*5.22"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.762*5.22", "0.762*5.22"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.762", "0.762*5.22"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.762*5.22", "0.762*5.22"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.762", "0.762*5.22"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.762*5.22", "0.762*5.22"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.762", "0.762*5.22"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.762*5.22", "0.762*5.22"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.762", "0.762*5.22"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.762*5.22", "0.762*5.22"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.762", "0.762*5.22"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.762*5.22", "0.762*5.22"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.762", "0.762*5.22"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.762*5.22", "0.762*5.22"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.762", "0.762*5.22"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.762*5.22", "0.762*5.22"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.762", "0.762*5.22"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.762*5.22", "0.762*5.22"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.762", "0.762*5.22"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.762*5.22", "0.762*5.22"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.762", "0.762*5.22"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.762*5.22", "0.762*5.22"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.762", "0.762*5.22"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.762*5.22", "0.762*5.22"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.762", "0.762*5.22"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.762*5.22", "0.762*5.22"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.762", "0.762*5.22"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.762*5.22", "0.762*5.22"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.762", "0.762*5.22"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.762*5.22", "0.762*5.22"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.762", "0.762*5.22"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.762*5.22", "0.762*5.22"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.762", "0.762*5.22"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.762*5.22", "0.762*5.22"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.762", "0.762*5.22"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.762*5.22", "0.762*5.22"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.762", "0.762*5.22"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.762*5.22", "0.762*5.22"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.762", "0.762*5.22"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.762*5.22", "0.762*5.22"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.762", "0.762*5.22"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.762*5.22", "0.762*5.22"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.762", "0.762*5.22"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.762*5.22", "0.762*5.22"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.762", "0.762*5.22"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.762*5.22", "0.762*5.22"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.762", "0.762*5.22"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.508*4.68", "0.508*4.68"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.508", "0.508*4.68"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.508*4.68", "0.508*4.68"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.508", "0.508*4.68"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.508*4.68", "0.508*4.68"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.508", "0.508*4.68"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ clear picks

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.ClearAllPicks

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.508*4.68", "0.508*4.68"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.508", "0.508*4.68"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.508*4.68", "0.508*4.68"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.508", "0.508*4.68"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete monitors

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Monitor.Delete "e-field (f=27.5)" 
Monitor.Delete "e-field (f=28)" 
Monitor.Delete "e-field (f=29.5)" 
Monitor.Delete "e-field (f=30)" 
Monitor.Delete "e-field (f=40)" 
Monitor.Delete "farfield (f=27.5)" 
Monitor.Delete "farfield (f=28)" 
Monitor.Delete "farfield (f=29.5)" 
Monitor.Delete "farfield (f=30)" 
Monitor.Delete "farfield (f=40)" 
Monitor.Delete "h-field (f=27.5)" 
Monitor.Delete "h-field (f=28)" 
Monitor.Delete "h-field (f=29.5)" 
Monitor.Delete "h-field (f=30)"

'@ define monitor: e-field (f=31)

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Monitor 
     .Reset 
     .Name "e-field (f=31)" 
     .Dimension "Volume" 
     .Domain "Frequency" 
     .FieldType "Efield" 
     .Frequency "31" 
     .UseSubvolume "False" 
     .SetSubvolume  "-5.6257104833333",  "5.6257104833333",  "-2.4982704833333",  "15.248270483333",  "-3.0412704833333",  "4.9107104833333" 
     .Create 
End With

'@ define monitor: h-field (f=31)

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Monitor 
     .Reset 
     .Name "h-field (f=31)" 
     .Dimension "Volume" 
     .Domain "Frequency" 
     .FieldType "Hfield" 
     .Frequency "31" 
     .UseSubvolume "False" 
     .SetSubvolume  "-5.6257104833333",  "5.6257104833333",  "-2.4982704833333",  "15.248270483333",  "-3.0412704833333",  "4.9107104833333" 
     .Create 
End With

'@ define farfield monitor: farfield (f=31)

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Monitor 
     .Reset 
     .Name "farfield (f=31)" 
     .Domain "Frequency" 
     .FieldType "Farfield" 
     .Frequency "31" 
     .UseSubvolume "False" 
     .ExportFarfieldSource "False" 
     .SetSubvolume  "-5.6257104833333",  "5.6257104833333",  "-2.4982704833333",  "15.248270483333",  "-3.0412704833333",  "4.9107104833333" 
     .Create 
End With

'@ define farfield monitor: farfield (f=32)

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Monitor 
     .Reset 
     .Name "farfield (f=32)" 
     .Domain "Frequency" 
     .FieldType "Farfield" 
     .Frequency "32" 
     .UseSubvolume "False" 
     .ExportFarfieldSource "False" 
     .SetSubvolume  "-5.6257104833333",  "5.6257104833333",  "-2.4982704833333",  "15.248270483333",  "-3.0412704833333",  "4.9107104833333" 
     .Create 
End With

'@ define monitor: h-field (f=32)

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Monitor 
     .Reset 
     .Name "h-field (f=32)" 
     .Dimension "Volume" 
     .Domain "Frequency" 
     .FieldType "Hfield" 
     .Frequency "32" 
     .UseSubvolume "False" 
     .SetSubvolume  "-5.6257104833333",  "5.6257104833333",  "-2.4982704833333",  "15.248270483333",  "-3.0412704833333",  "4.9107104833333" 
     .Create 
End With

'@ define monitor: e-field (f=32)

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Monitor 
     .Reset 
     .Name "e-field (f=32)" 
     .Dimension "Volume" 
     .Domain "Frequency" 
     .FieldType "Efield" 
     .Frequency "32" 
     .UseSubvolume "False" 
     .SetSubvolume  "-5.6257104833333",  "5.6257104833333",  "-2.4982704833333",  "15.248270483333",  "-3.0412704833333",  "4.9107104833333" 
     .Create 
End With

'@ define monitor: e-field (f=33)

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Monitor 
     .Reset 
     .Name "e-field (f=33)" 
     .Dimension "Volume" 
     .Domain "Frequency" 
     .FieldType "Efield" 
     .Frequency "33" 
     .UseSubvolume "False" 
     .SetSubvolume  "-5.6257104833333",  "5.6257104833333",  "-2.4982704833333",  "15.248270483333",  "-3.0412704833333",  "4.9107104833333" 
     .Create 
End With

'@ define monitor: h-field (f=33)

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Monitor 
     .Reset 
     .Name "h-field (f=33)" 
     .Dimension "Volume" 
     .Domain "Frequency" 
     .FieldType "Hfield" 
     .Frequency "33" 
     .UseSubvolume "False" 
     .SetSubvolume  "-5.6257104833333",  "5.6257104833333",  "-2.4982704833333",  "15.248270483333",  "-3.0412704833333",  "4.9107104833333" 
     .Create 
End With

'@ define farfield monitor: farfield (f=33)

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Monitor 
     .Reset 
     .Name "farfield (f=33)" 
     .Domain "Frequency" 
     .FieldType "Farfield" 
     .Frequency "33" 
     .UseSubvolume "False" 
     .ExportFarfieldSource "False" 
     .SetSubvolume  "-5.6257104833333",  "5.6257104833333",  "-2.4982704833333",  "15.248270483333",  "-3.0412704833333",  "4.9107104833333" 
     .Create 
End With

'@ delete monitor: h-field (f=40)

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Monitor.Delete "h-field (f=40)"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:2

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "2"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.508*4.68", "0.508*4.68"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.508", "0.508*4.68"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1"

'@ delete port: port2

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "2"

'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3"

'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.508*4.68", "0.508*4.68"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.508", "0.508*4.68"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With

'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ define time domain solver parameters

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.SetCreator "High Frequency" 

With Solver 
     .Method "Hexahedral"
     .CalculationType "TD-S"
     .StimulationPort "All"
     .StimulationMode "All"
     .SteadyStateLimit "-40"
     .MeshAdaption "False"
     .AutoNormImpedance "True"
     .NormingImpedance "50"
     .CalculateModesOnly "False"
     .SParaSymmetry "False"
     .StoreTDResultsInCache  "False"
     .FullDeembedding "False"
     .SuperimposePLWExcitation "False"
     .UseSensitivityAnalysis "False"
End With


'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ define monitor: e-field (f=31.15)

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Monitor 
     .Reset 
     .Name "e-field (f=31.15)" 
     .Dimension "Volume" 
     .Domain "Frequency" 
     .FieldType "Efield" 
     .Frequency "31.15" 
     .UseSubvolume "False" 
     .SetSubvolume  "-5.6257104833333",  "5.6257104833333",  "-2.4982704833333",  "15.248270483333",  "-3.0412704833333",  "4.9107104833333" 
     .Create 
End With 


'@ define monitor: h-field (f=31.15)

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Monitor 
     .Reset 
     .Name "h-field (f=31.15)" 
     .Dimension "Volume" 
     .Domain "Frequency" 
     .FieldType "Hfield" 
     .Frequency "31.15" 
     .UseSubvolume "False" 
     .SetSubvolume  "-5.6257104833333",  "5.6257104833333",  "-2.4982704833333",  "15.248270483333",  "-3.0412704833333",  "4.9107104833333" 
     .Create 
End With 


'@ define farfield monitor: farfield (f=31.15)

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Monitor 
     .Reset 
     .Name "farfield (f=31.15)" 
     .Domain "Frequency" 
     .FieldType "Farfield" 
     .Frequency "31.15" 
     .UseSubvolume "False" 
     .ExportFarfieldSource "False" 
     .SetSubvolume  "-5.6257104833333",  "5.6257104833333",  "-2.4982704833333",  "15.248270483333",  "-3.0412704833333",  "4.9107104833333" 
     .Create 
End With 


'@ define farfield monitor: farfield (f=32.15)

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Monitor 
     .Reset 
     .Name "farfield (f=32.15)" 
     .Domain "Frequency" 
     .FieldType "Farfield" 
     .Frequency "32.15" 
     .UseSubvolume "False" 
     .ExportFarfieldSource "False" 
     .SetSubvolume  "-5.6257104833333",  "5.6257104833333",  "-2.4982704833333",  "15.248270483333",  "-3.0412704833333",  "4.9107104833333" 
     .Create 
End With 


'@ define monitor: h-field (f=32.15)

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Monitor 
     .Reset 
     .Name "h-field (f=32.15)" 
     .Dimension "Volume" 
     .Domain "Frequency" 
     .FieldType "Hfield" 
     .Frequency "32.15" 
     .UseSubvolume "False" 
     .SetSubvolume  "-5.6257104833333",  "5.6257104833333",  "-2.4982704833333",  "15.248270483333",  "-3.0412704833333",  "4.9107104833333" 
     .Create 
End With 


'@ define monitor: e-field (f=32.15)

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Monitor 
     .Reset 
     .Name "e-field (f=32.15)" 
     .Dimension "Volume" 
     .Domain "Frequency" 
     .FieldType "Efield" 
     .Frequency "32.15" 
     .UseSubvolume "False" 
     .SetSubvolume  "-5.6257104833333",  "5.6257104833333",  "-2.4982704833333",  "15.248270483333",  "-3.0412704833333",  "4.9107104833333" 
     .Create 
End With 


'@ define monitor: e-field (f=33.15)

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Monitor 
     .Reset 
     .Name "e-field (f=33.15)" 
     .Dimension "Volume" 
     .Domain "Frequency" 
     .FieldType "Efield" 
     .Frequency "33.15" 
     .UseSubvolume "False" 
     .SetSubvolume  "-5.6257104833333",  "5.6257104833333",  "-2.4982704833333",  "15.248270483333",  "-3.0412704833333",  "4.9107104833333" 
     .Create 
End With 


'@ define monitor: h-field (f=33.15)

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Monitor 
     .Reset 
     .Name "h-field (f=33.15)" 
     .Dimension "Volume" 
     .Domain "Frequency" 
     .FieldType "Hfield" 
     .Frequency "33.15" 
     .UseSubvolume "False" 
     .SetSubvolume  "-5.6257104833333",  "5.6257104833333",  "-2.4982704833333",  "15.248270483333",  "-3.0412704833333",  "4.9107104833333" 
     .Create 
End With 


'@ define farfield monitor: farfield (f=33.15)

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With Monitor 
     .Reset 
     .Name "farfield (f=33.15)" 
     .Domain "Frequency" 
     .FieldType "Farfield" 
     .Frequency "33.15" 
     .UseSubvolume "False" 
     .ExportFarfieldSource "False" 
     .SetSubvolume  "-5.6257104833333",  "5.6257104833333",  "-2.4982704833333",  "15.248270483333",  "-3.0412704833333",  "4.9107104833333" 
     .Create 
End With 


'@ delete monitors

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Monitor.Delete "e-field (f=31)" 
Monitor.Delete "e-field (f=32)" 
Monitor.Delete "e-field (f=33)" 
Monitor.Delete "farfield (f=31)" 
Monitor.Delete "farfield (f=32)" 
Monitor.Delete "farfield (f=33)" 
Monitor.Delete "h-field (f=31)" 
Monitor.Delete "h-field (f=32)" 
Monitor.Delete "h-field (f=33)" 


'@ delete port: port1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Port.Delete "1" 


'@ pick face

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Pick.PickFaceFromId "component1:solid1", "3" 


'@ define port:1

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
' Port constructed by macro Calculate -> Calculate port extension coefficient


With Port
  .Reset
  .PortNumber "1"
  .NumberOfModes "1"
  .AdjustPolarization False
  .PolarizationAngle "0.0"
  .ReferencePlaneDistance "0"
  .TextSize "50"
  .Coordinates "Picks"
  .Orientation "Positive"
  .PortOnBound "True"
  .ClipPickedPortToBound "False"
  .XrangeAdd "0.508*4.68", "0.508*4.68"
  .YrangeAdd "0", "0"
  .ZrangeAdd "0.508", "0.508*4.68"
  .Shield "PEC"
  .SingleEnded "False"
  .Create
End With



'@ set pba mesh type

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
Mesh.MeshType "PBA"

'@ farfield plot options

'[VERSION]2014.0|23.0.0|20140224[/VERSION]
With FarfieldPlot 
     .Plottype "3D" 
     .Vary "angle1" 
     .Theta "90" 
     .Phi "90" 
     .Step "5" 
     .Step2 "5" 
     .SetLockSteps "True" 
     .SetPlotRangeOnly "False" 
     .SetThetaStart "0" 
     .SetThetaEnd "180" 
     .SetPhiStart "0" 
     .SetPhiEnd "360" 
     .SetTheta360 "False" 
     .SymmetricRange "False" 
     .SetTimeDomainFF "False" 
     .SetFrequency "-1" 
     .SetTime "0" 
     .SetColorByValue "True" 
     .DrawStepLines "False" 
     .DrawIsoLongitudeLatitudeLines "False" 
     .ShowStructure "False" 
     .SetStructureTransparent "False" 
     .SetFarfieldTransparent "False" 
     .SetSpecials "enablepolarextralines" 
     .SetPlotMode "Directivity" 
     .Distance "1" 
     .UseFarfieldApproximation "True" 
     .SetScaleLinear "False" 
     .SetLogRange "40" 
     .SetLogNorm "0" 
     .DBUnit "0" 
     .EnableFixPlotMaximum "False" 
     .SetFixPlotMaximumValue "1" 
     .SetInverseAxialRatio "False" 
     .SetAxesType "user" 
     .SetAnntenaType "directional_circular" 
     .Phistart "1.000000e+000", "0.000000e+000", "0.000000e+000" 
     .Thetastart "0.000000e+000", "0.000000e+000", "1.000000e+000" 
     .PolarizationVector "0.000000e+000", "1.000000e+000", "0.000000e+000" 
     .SetCoordinateSystemType "ludwig3" 
     .SetAutomaticCoordinateSystem "True" 
     .SetPolarizationType "Circular" 
     .SlantAngle 0.000000e+000 
     .Origin "bbox" 
     .Userorigin "0.000000e+000", "0.000000e+000", "0.000000e+000" 
     .SetUserDecouplingPlane "False" 
     .UseDecouplingPlane "False" 
     .DecouplingPlaneAxis "X" 
     .DecouplingPlanePosition "0.000000e+000" 
     .LossyGround "False" 
     .GroundEpsilon "1" 
     .GroundKappa "0" 
     .EnablePhaseCenterCalculation "False" 
     .SetPhaseCenterAngularLimit "3.000000e+001" 
     .SetPhaseCenterComponent "boresight" 
     .SetPhaseCenterPlane "both" 
     .ShowPhaseCenter "True" 
     .StoreSettings
End With 


