Attribute VB_Name = "Overcut"
'===============================================================================
'   Макрос          : Overcut
'   Версия          : 2025.03.17
'   Сайты           : https://vk.com/elvin_macro
'                     https://github.com/elvin-nsk
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

'===============================================================================
' # Manifest

Public Const APP_NAME As String = "Overcut"
Public Const APP_DISPLAYNAME As String = APP_NAME
Public Const APP_FILEBASENAME As String = "elvin_" & APP_NAME
Public Const APP_VERSION As String = "2025.03.17"
Public Const APP_URL As String = "https://vk.com/elvin_macro/" & APP_NAME

'===============================================================================
' # Globals

Private Const OVERLAY_LENGTH As Double = 1 'mm

'===============================================================================
' # Entry points

Sub Start()

    #If DebugMode = 0 Then
    On Error GoTo Catch
    #End If
    
    Dim Shapes As ShapeRange
    If Not InputData.ExpectShapes.Ok(Shapes) Then GoTo Finally
        
    Set Shapes = Shapes.UngroupAllEx
    FilterValidForTrace Shapes
    If Shapes.Count = 0 Then
        Warn "Ни один из выделенных объектов не является кривой.", APP_DISPLAYNAME
        Exit Sub
    End If
    
    BoostStart APP_DISPLAYNAME
    
    ProcessShapes Shapes
    
Finally:
    BoostFinish
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Source & ": " & VBA.Err.Description, vbCritical, "Error"
    Resume Finally

End Sub

'===============================================================================
' # Helpers

Private Sub ProcessShapes(ByVal ValidShapes As ShapeRange)
    Dim Shape As Shape
    For Each Shape In ValidShapes
        ProcessShape Shape
    Next Shape
End Sub

Private Sub ProcessShape(ByVal ValidShape As Shape)
    If ValidShape.Curve.SubPaths.Count > 1 Then
        ProcessShapes ValidShape.BreakApartEx
    Else
        If ValidShape.Curve.SubPaths.First.Closed Then _
            ProcessOnePathCurve ValidShape.Curve
    End If
End Sub

Private Sub ProcessOnePathCurve(ByVal Curve As Curve)
    Curve.SubPaths.First.Closed = False
    
    MayMakeNode Curve
    
    Dim NewCurve As Curve
    Set NewCurve = Curve.Segments.First.GetCopy
    Curve.AppendCurve NewCurve
    
    'Debug
    'Curve.SubPaths.Last.Nodes.First.Move 10, 10
    'Curve.SubPaths.First.Nodes.Last.Move 10, 10
    
    Curve.SubPaths.Last.Nodes.First.JoinWith _
        Curve.SubPaths.First.Nodes.Last
End Sub

Private Sub MayMakeNode(ByVal Curve As Curve)
    If OVERLAY_LENGTH >= Curve.SubPaths.First.Segments.First.Length Then Exit Sub
    Curve.SubPaths.First.AddNodeAt OVERLAY_LENGTH, cdrAbsoluteSegmentOffset
End Sub

Private Sub FilterValidForTrace(ByRef Shapes As ShapeRange)
    Dim ValidShapes As New ShapeRange
    Dim Shape As Shape
    For Each Shape In Shapes
        If HasCurve(Shape) Then ValidShapes.Add Shape
    Next Shape
    Set Shapes = ValidShapes
End Sub

'===============================================================================
' # Tests

Private Sub TestSomething()
'
End Sub
