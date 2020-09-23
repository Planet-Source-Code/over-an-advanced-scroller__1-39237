Attribute VB_Name = "basScroller"

'#################################################
'
' To add a Scroller to your own Program simply add
' this bas File to your Project and have a look
' at the examples in frmMain.
'
' Syntax:
'
' Scroller NameOfPictureBox, "Your Message", FontSize, Animationspeed, Oscillation, Should Colors constantly change ?, otherwise name the Textcolor here
'
' Author: Over (overkillpage@gmx.net)
'
'#################################################

Option Explicit

Public Sub Scroller(PicBox As PictureBox, ScrollTxt As String, FontSizeA As Integer, Speed As Double, Oscillation As Double, ChangingColors As Boolean, FontColor As ColorConstants)

    'Declarations
    Dim i As Double                             'Var for the Main Loop
    Dim j As Integer                            'another loop var
    Dim StartTime As Double                     'Important for the Animation Speed
        
    Dim RedCounter As Double                    'Needed
    Dim RedDown As Boolean                      'if changing
    Dim GreenCounter As Double                  'colors
    Dim GreenDown As Boolean                    'are
    Dim BlueCounter As Double                   'enabled
    Dim BlueDown As Boolean
        
    'Preparing the PictureBox
    PicBox.ScaleMode = 4                        'Set Scalemode to Font
    
    'Font Settings, make changes
    'here if you want too
    PicBox.FontName = "Courier New"             '"Courier" or other fonts work, too. But every
                                                'Letter must have the same width !
    PicBox.FontSize = FontSizeA                 'Setting Fontsize
    PicBox.FontBold = False                     'Setting Bold Mode
    PicBox.FontItalic = False                   'Setting Italic Mode
    PicBox.FontUnderline = False                'Underline looks a bit funny ;)
    PicBox.ForeColor = FontColor
        
    'Main Animation Loop
    For i = Int(PicBox.ScaleWidth) To Int(-Len(ScrollTxt) * (PicBox.FontSize / 10)) Step -0.5
        
        StartTime = Timer                       'Animation
        Do                                      'Speed
            DoEvents                            'Loop
        Loop Until Timer - StartTime >= Speed * 0.01
        
        PicBox.Cls                              'Clearing PictureBox for new Frame
        
        'Calculate Color Settings
        If ChangingColors = True Then
            If RedCounter >= 255 Then RedDown = True
            If RedCounter <= 100 Then RedDown = False
            If RedDown = True Then RedCounter = RedCounter - 3 Else RedCounter = RedCounter + 3
            If GreenCounter >= 255 Then GreenDown = True
            If GreenCounter <= 100 Then GreenDown = False
            If GreenDown = True Then GreenCounter = GreenCounter - 1 Else GreenCounter = GreenCounter + 1
            If BlueCounter >= 255 Then BlueDown = True
            If BlueCounter <= 100 Then BlueDown = False
            If BlueDown = True Then BlueCounter = BlueCounter - 2 Else BlueCounter = BlueCounter + 2
            PicBox.ForeColor = RGB(RedCounter, GreenCounter, BlueCounter)
        End If
        
        'Printing out our Text letter by letter and adding up down movement
        For j = 0 To Len(ScrollTxt) - 1
        
            PicBox.CurrentX = i + j * (PicBox.FontSize / 10)
            PicBox.CurrentY = (PicBox.ScaleHeight / 2) - 1 + (Sin((j + i) * 0.7)) * Oscillation 'Center text (vertical) + Up-Down-Movement
            PicBox.Print Mid(ScrollTxt, j + 1, 1)
            
        Next j
    
    Next i
    
End Sub
