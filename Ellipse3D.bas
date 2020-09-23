Attribute VB_Name = "basEllipse3D"
Option Explicit

'Type structures:
Public Type PointSng    'Point structure (uses Singles for accuracy).
    X   As Single
    Y   As Single
End Type

Public Type RectSng     'Rect structure (uses Singles for accuracy).
    Left    As Single
    Top     As Single
    Right   As Single
    Bottom  As Single
End Type

Public Type EllipseData 'The Ellipse properties
    Easel       As Object   'Must be a Form or PictureBox.
    rcBounds    As RectSng  'Bounding rectangle for the ellipse.
    ptLight     As PointSng 'The point where the light hits the ellipse.
    'The following colors may be swapped to produce a concave effect.
    LightColor  As Long     'The color of the light point.
    BackColor   As Long     'The color of the ellipse.
End Type

Private Type PointAPI   'Windows API point structure (uses Longs)
    X   As Long
    Y   As Long
End Type

Private Type RectAPI    'Windows API rect structure (uses Longs)
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

'API Declares:
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RectAPI) As Long
Private Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long

Public Function DrawEllipse(uEllipse As EllipseData, Optional ByVal bShowErrorMsg As Boolean = False) As Boolean

'Note uEllipse.Easel must be a Form or PictureBox.

Dim lIdx        As Long
Dim lMaxX       As Long
Dim lMaxY       As Long
Dim lMaxSteps   As Long
Dim lRet        As Long
Dim hRgn        As Long
Dim lTemp       As Long
Dim laColors()  As Long
Dim lForeColor  As Long
Dim lFillColor  As Long
Dim iFillStyle  As Integer
Dim iDrawWidth  As Integer
Dim iDrawStyle  As Integer
Dim iDrawMode   As Integer
Dim fScale      As Single
Dim ptLight     As PointAPI
Dim rcEasel     As RectAPI
Dim rcBounds    As RectAPI
Dim rcStepAmt   As RectSng
Dim rcEllipse   As RectSng

    'Safety first! :oþ
    On Error GoTo LocalError
    
    If Not uEllipse.Easel Is Nothing Then
        
        'Convert Easel to pixels and calculate the scale factor.
        lRet = GetClientRect(uEllipse.Easel.hWnd, rcEasel)
        fScale = uEllipse.Easel.ScaleWidth / rcEasel.Right
        
        With uEllipse
            With .Easel
                'Save the Easel settings so they may be restored.
                lForeColor = .ForeColor
                lFillColor = .FillColor
                iFillStyle = .FillStyle
                iDrawWidth = .DrawWidth
                iDrawStyle = .DrawStyle
                iDrawMode = .DrawMode
                'Now set them to the correct values for this procedure.
                .DrawWidth = 1
                .DrawStyle = vbSolid
                .DrawMode = vbCopyPen
                .FillStyle = vbFSSolid
            End With
            
            'Convert all coordinates to pixels. APIs use pixels only.
            With .ptLight   'Note - ptLight is not the same as .ptLight.
                ptLight.X = Div(.X, fScale)
                ptLight.Y = Div(.Y, fScale)
            End With
            With .rcBounds  'Note - rcBounds is not the same as .rcBounds.
                rcBounds.Left = Div(.Left, fScale)
                rcBounds.Top = Div(.Top, fScale)
                rcBounds.Right = Div(.Right, fScale)
                rcBounds.Bottom = Div(.Bottom, fScale)
        
            End With
        End With
        
        'Validate coordinates.
        If rcBounds.Right < rcBounds.Left Then
            'Swap left and right sides if reversed.
            lTemp = rcBounds.Left
            rcBounds.Left = rcBounds.Right
            rcBounds.Right = lTemp
        ElseIf rcBounds.Right = rcBounds.Left Then
            'Zero width. Show error and get out.
            If bShowErrorMsg Then
                MsgBox "Ellipse cannot have a zero width", vbExclamation
            End If
            GoTo NormalExit
        End If
        If rcBounds.Bottom < rcBounds.Top Then
            'Swap top and bottom sides if reversed.
            lTemp = rcBounds.Top
            rcBounds.Top = rcBounds.Bottom
            rcBounds.Bottom = lTemp
        ElseIf rcBounds.Bottom = rcBounds.Top Then
            'Zero height. Show error and get out.
            If bShowErrorMsg Then
                MsgBox "Ellipse cannot have a zero Height", vbExclamation
            End If
            GoTo NormalExit
        End If
        
        'Validate the light point. It must be within the ellipse.
        With rcBounds
            hRgn = CreateEllipticRgn(.Left, .Top, .Right, .Bottom)
            If PtInRegion(hRgn, ptLight.X, ptLight.Y) = 0 Then
                'Light point out of bounds. Show error and get out.
                If bShowErrorMsg Then
                    MsgBox "Light point must be within the bounds of the ellipse.", vbExclamation
                End If
                GoTo NormalExit
            End If
        End With
        
        'Calculate the number of pixels from the light
        'point to the furthest border. This will be the
        'maximum number of colors needed in the blend
        'and the maximum steps needed to draw the ellipse.
        With rcEllipse
            If Abs(.Right - ptLight.X) > Abs(.Left - ptLight.X) Then
                lMaxX = Abs(.Right - ptLight.X)
            Else
                lMaxX = Abs(.Left - ptLight.X)
            End If
            If Abs(.Bottom - ptLight.Y) > Abs(.Top - ptLight.Y) Then
                lMaxY = Abs(.Bottom - ptLight.Y)
            Else
                lMaxY = Abs(.Top - ptLight.Y)
            End If
        End With
        If lMaxX > lMaxY Then
            lMaxSteps = lMaxX
        Else
            lMaxSteps = lMaxY
        End If
        
        'The BlendColors routine may adjust lMaxSteps, but it returns the
        'adjusted lMaxSteps. It also redims and fills the laColors() array.
        'Blend the colors starting from BackColor and going to LightColor.
        lMaxSteps = BlendColors(uEllipse.BackColor, uEllipse.LightColor, lMaxSteps, laColors)
        
        'All coordinate calcs must be done in singles from here on for accuracy.
        With rcEllipse  'RectSng version
            .Left = rcBounds.Left
            .Top = rcBounds.Top
            .Right = rcBounds.Right
            .Bottom = rcBounds.Bottom
        End With
        
        'Calculate the distance to move each of the four bounding
        'sides for each step of the loop that draws the ellipse.
        With rcStepAmt  'RectSng version
            .Left = Div(Abs(rcBounds.Left - CSng(ptLight.X)), CSng(lMaxSteps))
            .Top = Div(Abs(rcBounds.Top - CSng(ptLight.Y)), CSng(lMaxSteps))
            .Right = Div(Abs(rcBounds.Right - CSng(ptLight.X)), CSng(lMaxSteps))
            .Bottom = Div(Abs(rcBounds.Bottom - CSng(ptLight.Y)), CSng(lMaxSteps))
        End With
        
        'Draw the Ellipse. This starts by drawing the ellipse in the BackColor
        'using the full size rectangle. Then each loop shrinks the rectangle
        'toward the ptLight coordinates using the next color in the blend, until
        'it reaches the light point at the light color.
        With rcEllipse
            For lIdx = 0 To lMaxSteps - 1
                'Setup the color for this loop.
                uEllipse.Easel.ForeColor = laColors(lIdx)
                uEllipse.Easel.FillColor = laColors(lIdx)
                'Draw the ellipse.
                lRet = Ellipse(uEllipse.Easel.hDC, .Left, .Top, .Right, .Bottom)
                'Shrink the rectangle for the next loop.
                .Left = .Left + rcStepAmt.Left
                .Top = .Top + rcStepAmt.Top
                .Right = .Right - rcStepAmt.Right
                .Bottom = .Bottom - rcStepAmt.Bottom
            Next lIdx
        End With
        
        'APIs draw to the Image, not the Picture when AutoRedraw is on,
        'so use refresh to copy the Image to the Picture.
        If uEllipse.Easel.AutoRedraw Then
            uEllipse.Easel.Refresh
        End If
        
        'Return True - No errors if it got this far.
        DrawEllipse = True
    
    Else
        If bShowErrorMsg Then
            MsgBox "Easel has not been setup for the DrawEllipse function.", vbExclamation
        End If
    End If

NormalExit:
    'Cleanup and restore settings.
    
    'Delete the elliptic region to free resources.
    If hRgn <> 0 Then   'May hit here on error before creating region.
        lRet = DeleteObject(hRgn)
    End If
    
    'Restore the Easel settings
    If iDrawWidth > 0 Then  'May hit here on error before setting properties.
        With uEllipse.Easel
            .ForeColor = lForeColor
            .FillColor = lFillColor
            .FillStyle = iFillStyle
            .DrawWidth = iDrawWidth
            .DrawStyle = iDrawStyle
            .DrawMode = iDrawMode
        End With
    End If
    
    Exit Function
    
LocalError:
    'Show the error and resume to NormalExit for cleanup.
    If bShowErrorMsg Then
        MsgBox Err.Description, vbExclamation
    End If
    Resume NormalExit
    
End Function


Private Function BlendColors(ByVal lColor1 As Long, ByVal lColor2 As Long, ByVal lSteps As Long, laRetColors() As Long) As Long

'Creates an array of colors blending from
'Color1 to Color2 in lSteps number of steps.
'Returns the count and fills the laRetColors() array.

Dim lIdx    As Long
Dim lRed    As Long
Dim lGrn    As Long
Dim lBlu    As Long
Dim fRedStp As Single
Dim fGrnStp As Single
Dim fBluStp As Single

    'Stop possible error
    If lSteps < 2 Then lSteps = 2
    
    'Extract Red, Blue and Green values from the start and end colors.
    lRed = (lColor1 And &HFF&)
    lGrn = (lColor1 And &HFF00&) / &H100
    lBlu = (lColor1 And &HFF0000) / &H10000
    
    'Find the amount of change for each color element per color change.
    fRedStp = Div(CSng((lColor2 And &HFF&) - lRed), CSng(lSteps))
    fGrnStp = Div(CSng(((lColor2 And &HFF00&) / &H100&) - lGrn), CSng(lSteps))
    fBluStp = Div(CSng(((lColor2 And &HFF0000) / &H10000) - lBlu), CSng(lSteps))
    
    'Create the colors
    ReDim laRetColors(lSteps - 1)
    laRetColors(0) = lColor1            'First Color
    laRetColors(lSteps - 1) = lColor2   'Last Color
    For lIdx = 1 To lSteps - 2          'All Colors between
        laRetColors(lIdx) = CLng(lRed + (fRedStp * CSng(lIdx))) + _
            (CLng(lGrn + (fGrnStp * CSng(lIdx))) * &H100&) + _
            (CLng(lBlu + (fBluStp * CSng(lIdx))) * &H10000)
    Next lIdx
    
    'Return number of colors in array
    BlendColors = lSteps

End Function

Private Function Div(ByVal dNumer As Double, ByVal dDenom As Double) As Double
    
'Divides dNumer by dDenom if dDenom <> 0
'Eliminates 'Division By Zero' error.

    If dDenom <> 0 Then
        Div = dNumer / dDenom
    Else
        Div = 0
    End If

End Function

