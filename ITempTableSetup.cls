VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITempTableSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' A helper class to let us pretend we have object constructors.

' You would never do this in a language that supported object constructors... _
  but VBA doesn't, so....
  
Public Sub Setup(cn As ADODB.Connection, tableConverter As ITableConverter, Optional withAutomaticTearDown As Boolean = True)

End Sub

