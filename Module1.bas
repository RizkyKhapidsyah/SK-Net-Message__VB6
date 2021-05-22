Attribute VB_Name = "Main"
Public Container As IADsContainer
Public ContainerName As String
Public StrComputerName As String
Public ComputerDomain As String
Public dso As IADsOpenDSObject
Public Computer As IADsComputer
Public StrDomain As String
Public Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal FileName As String, ByVal uFlags As Long) As Long
