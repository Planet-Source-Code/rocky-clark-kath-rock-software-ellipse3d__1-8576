VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "3D Ellipse - (Click to change the light point. - Resize to shape the ellipse.)"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   5790
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar sbrStatus 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4869
            MinWidth        =   159
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4869
            MinWidth        =   159
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Ellipse data (Module level variable).
Dim muEllipse   As EllipseData

Private Sub DoEllipse()

Dim bErr    As Boolean
Dim fScaleX As Single
Dim fScaleY As Single

    'Setup the Ellipse properties.
    With muEllipse
        
        'Calc the scale of the light point in reference to the ellipse.
        With .rcBounds
            If (.Right - .Left > 0) And (.Bottom - .Top > 0) Then
                fScaleX = (muEllipse.ptLight.X - .Left) / (.Right - .Left)
                fScaleY = (muEllipse.ptLight.Y - .Top) / (.Bottom - .Top)
            Else
                With muEllipse.ptLight
                    .X = Me.ScaleWidth / 2.6
                    .Y = (Me.ScaleHeight - sbrStatus.Height) / 2.6
                End With
                bErr = True
            End If
        End With
            
        Set .Easel = Me 'Must be Form or PictureBox.
        With .rcBounds  'Bounding rectangle of the ellipse.
            .Left = Me.ScaleX(10, vbPixels, Me.ScaleMode)
            .Top = Me.ScaleY(10, vbPixels, Me.ScaleMode)
            .Right = Me.ScaleWidth - .Left
            .Bottom = (Me.ScaleHeight - sbrStatus.Height) - .Top
            'Set the point where the light hits the ellipse.
            If Not bErr Then
                muEllipse.ptLight.X = .Left + (.Right - .Left) * fScaleX
                muEllipse.ptLight.Y = .Top + (.Bottom - .Top) * fScaleY
            End If
        End With
        .LightColor = &HFFC0FF  'The color of the light point.
        .BackColor = &H600060   'The color of the ellipse.
    
    End With
    
    'Clear the form and draw the ellipse.
    Me.Cls
    bErr = Not DrawEllipse(muEllipse)    'Draw the ellipse
    
    'Update the status
    With sbrStatus
        If Not bErr Then
            fScaleX = Me.ScaleX(1, Me.ScaleMode, vbPixels)
            fScaleY = Me.ScaleY(1, Me.ScaleMode, vbPixels)
            .Panels(1).Text = "Size: " & Format$((muEllipse.rcBounds.Right - muEllipse.rcBounds.Left) * fScaleX, "0") & " x " & Format$((muEllipse.rcBounds.Bottom - muEllipse.rcBounds.Top) * fScaleY, "0")
            .Panels(2).Text = "Light Pt: " & Format$(muEllipse.ptLight.X * fScaleX, "0") & ", " & Format$(muEllipse.ptLight.Y * fScaleY, "0")
        Else
            With muEllipse.ptLight
                .X = Me.ScaleWidth / 2.6
                .Y = (Me.ScaleHeight - sbrStatus.Height) / 2.6
            End With
            .Panels(1).Text = "Size: error"
            .Panels(2).Text = "Light: error"
        End If
    
    End With

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Move the light point coordinate.
    muEllipse.ptLight.X = X
    muEllipse.ptLight.Y = Y
    
    'Redraw the ellipse.
    Call DoEllipse
    
End Sub


Private Sub Form_Resize()

    'Make sure the form is not minimized.
    If Me.WindowState <> vbMinimized Then
        
        'Make sure the width and height are not too small to draw in.
        If (Me.ScaleWidth > 60) And ((Me.ScaleHeight - sbrStatus.Height) > 60) Then
            
            'Redraw the ellipse
            Call DoEllipse
        
        End If
    
    End If
    
End Sub

