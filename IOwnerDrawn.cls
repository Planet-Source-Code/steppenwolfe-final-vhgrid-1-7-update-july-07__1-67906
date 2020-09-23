VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IOwnerDrawn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Sub Draw(ByRef GridCell As clsGridItem, _
                ByVal lRow As Long, _
                ByVal lCell As Long, _
                ByVal lHdc As Long, _
                ByVal eDrawStage As EGDDrawStage, _
                ByVal lLeft As Long, _
                ByVal lTop As Long, _
                ByVal lRight As Long, _
                ByVal lBottom As Long, _
                ByRef bSkipDefault As Boolean)

End Sub

