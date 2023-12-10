VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAnchor"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'********************************************
' Based on http://www.freevbcode.com/ShowCode.asp?ID=6075
'
'
' In the first 4 characters of Tag property of the control would be a combination of L, T, R and B
' The method is case insensitive
'********************************************
'

'In the form you want to anchor:
'Private Sub Form_Resize()
'    Static oAnchor As clsAnchor
'On Error GoTo Err_Form_Resize
'    If Not oAnchor Is Nothing Then
'    Else
'        Set oAnchor = New clsAnchor
'        oAnchor.Form = Me
'    End If
'
'    oAnchor.Anchor
'
'
'    Exit Sub
'Err_Form_Resize:
'    Screen.MousePointer = vbDefault
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en  Sub Form1.Form_Resize" & vbCrLf & "[" & Err.Source & "]", vbCritical
'End Sub
'
'version 1.1
'I add the posibility to use the Tag property of the control to contains more than one 'property'
'If this is the case, you have to separate the diferent properties by the string contained in the Separator property of the class, the default value is ","
'And you have to specify the index (cero based) of the Anchor property in the PropertyNumber property of the class, the default value, -1, indicates
'that the class works like the versión 1 and take the 4 first chars for the Anchor property
'
'Now in the form you can use something like this:
'
'
'Private Sub Form_Resize()
'    Static oAnchor As clsAnchor
'On Error GoTo Err_Form_Resize
'    If Not oAnchor Is Nothing Then
'    Else
'        Set oAnchor = New clsAnchor
'        'optional if not is used the separator will be ","
'        oAnchor.Separator = ";"
'        oAnchor.PropertyNumber = 1
'        oAnchor.Form = Me
'    End If
'
'    oAnchor.Anchor
'
'
'    Exit Sub
'Err_Form_Resize:
'    Screen.MousePointer = vbDefault
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en  Sub Form1.Form_Resize" & vbCrLf & "[" & Err.Source & "]", vbCritical
'End Sub
'
'in this case the tag property of the control have to be like this:
'       "property one;TB;...                Anchor = TB
'       ";LB"                               Anchor = LB
'       "property one"                      No Anchor property
'       "property one;;other property       No Anchor property
'.- Version 1.1.1
'Fixed:
'The ListBoxes Height Property is not free, so in previous versions if you increase manually the heigth of the form, the heigth of the listboxes does not increase.
'Now the ListBoxes Height property changes correctly

Private Type POINTAPI
        x As Long
        y As Long
End Type
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Type WINDOWPLACEMENT
    Length As Long
    Flags As Long
    ShowCmd As Long
    ptMinPosition As POINTAPI
    ptMaxPosition As POINTAPI
    rcNormalPosition As RECT
End Type
Private Declare Function GetWindowPlacement Lib "user32" ( _
    ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
 
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Private frm As Form

Private pPropertyNumber As Long
Private pSeparator As String

Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long

Public Sub Anchor()
' MinWidth and MinHeight save the first value of form and don't let form to get smaller
    
    Static MinHeight As Long
    Static MinWidth As Long
' LastHeight and LastWidth are both save the last height and width
    Static LastHeight As Long
    Static LastWidth As Long

    
    Dim HeightChange As Long, WidthChange As Long
    Dim Tg As String
    Dim oControl As Control
    Dim wndpl As WINDOWPLACEMENT
    Dim sm As Long
    
    Dim ByIndex As Boolean
    
    On Error GoTo Err_Anchor
    ' Exit sub on Minimize
    If frm.WindowState = vbMinimized Then Exit Sub
    LockWindowUpdate frm.hWnd
    sm = frm.ScaleMode
    frm.ScaleMode = 1 'work with twips
    
 
    If LastHeight = 0 And LastWidth = 0 Then
        'This is the first Resize
        'Cant use frm.Height and frm.With because frm can be maximized
        'and in this case the form must be resized
  
        wndpl.Length = LenB(wndpl)
        If GetWindowPlacement(frm.hWnd, wndpl) <> 0 Then
            ' Function succeeded, Get the window dimensions
            With wndpl.rcNormalPosition
                LastHeight = (.Bottom - .Top) * Screen.TwipsPerPixelY
                LastWidth = (.Right - .Left) * Screen.TwipsPerPixelX
            End With
        End If

        MinHeight = LastHeight
        MinWidth = LastWidth
    End If
    If frm.WindowState <> vbMaximized Then
        ' Check the form for Min Values
        If frm.Height < MinHeight Then frm.Height = MinHeight
        If frm.Width < MinWidth Then frm.Width = MinWidth
    Else
        'If the window is maximized, but is smaller than the original window, change the min values to the actual size
        If frm.Height < MinHeight Then MinHeight = frm.Height
        If frm.Width < MinWidth Then MinWidth = frm.Width
    End If
    'Calculate the Changes
    HeightChange = frm.Height - LastHeight
    WidthChange = frm.Width - LastWidth
    ByIndex = pPropertyNumber > -1
    For Each oControl In frm.Controls
        On Error Resume Next
        If oControl.Container Is Nothing Then
            'If the control don't have a container, do nothing
        Else
            On Error GoTo Err_Anchor
            'v. 1.1 if pPropertyNumber <= -1, proceed like the original version
            If ByIndex Then
                'v. 1.1
                'The tag property can contain more properties codified not only the anchor
                'The tag property is a list of values separated by pSeparator, and the Anchor property is located at the postion (cero based) pPropertyNumber
                Dim Properties() As String ' As Variant
                Properties = Split(oControl.Tag, pSeparator)
                If UBound(Properties) >= pPropertyNumber Then
                    Tg = UCase(Properties(pPropertyNumber))
                Else
                    Tg = ""
                End If
            Else
                'In the v 1.0 if there more properties codified in the Tag property, the other properties should be after the 4th char o the tag property
                'Be case insensitive, and use only the 4 first chars
                Tg = UCase(Left(oControl.Tag, 4))
            End If
        
            If Not TypeOf oControl Is Line Then
          
                'Checking Tag
                If InStr(Tg, "B") Then   'Bottom
                    If InStr(Tg, "T") Then      'Top
                        'Only Anchor Top if the container is the form or the container is changes his Height
                        If HeightContainerChanges(oControl) Then
                            Select Case TypeName(oControl)
                                Case "ListBox"      'Listboxes Heitght are not free, so we have to capture the real final height
                                    Dim RealHeight As Long
                                    
                                    RealHeight = GetProp(oControl.hWnd, "RealHeight")
                                    If RealHeight <> 0 Then
                                        'not the first time, calculate the new real height
                                        RealHeight = RealHeight + HeightChange
                                    Else
                                        'Is the first time LastHeight is oControl.Height
                                        RealHeight = oControl.Height + HeightChange
                                    End If
                                    oControl.Height = RealHeight
                                    'Hold the real heigth of the control as a property of the window
                                    SetProp oControl.hWnd, "RealHeight", RealHeight
                                Case Else
                                    oControl.Height = oControl.Height + HeightChange
                            End Select
                        End If
                    Else
                        oControl.Top = oControl.Top + HeightChange
                    End If
                End If
          
                If InStr(Tg, "R") Then       'Right
                    If InStr(Tg, "L") Then      'Left
                        'Only Anchor Left if the container is the form or the container is anchored Left
                        If WidthContainerChanges(oControl) Then
                            oControl.Width = oControl.Width + WidthChange
                        End If
                    Else
                        oControl.Left = oControl.Left + WidthChange
                    End If
                End If
            Else
                'the line control don't have the Left, Top, Width and Height properties
                'it uses the X1, X2, Y1 and Y2 properties
                'Checking Tag
                If InStr(Tg, "B") Then   'Bottom
                    If InStr(Tg, "T") Then      'Top
                        'Only Anchor Top if the container is the form or the container is anchored Top
                        If HeightContainerChanges(oControl) Then
                            oControl.Y2 = oControl.Y2 + HeightChange
                        End If
                    Else
                        oControl.Y1 = oControl.Y1 + HeightChange
                    End If
                End If
          
                If InStr(Tg, "R") Then       'Right
                    If InStr(Tg, "L") Then      'Left
                        'Only Anchor Left if the container is the form or the container is anchored Left
                        If WidthContainerChanges(oControl) Then
                            oControl.X2 = oControl.X2 + WidthChange
                        End If
                    Else
                        oControl.X1 = oControl.Y2 + WidthChange
                    End If
                End If
          
            End If
        End If
    Next oControl
    'Save Last values
    LastHeight = frm.Height
    LastWidth = frm.Width
    'restore the form's salemode
    frm.ScaleMode = sm
    LockWindowUpdate 0

    Exit Sub
Err_Anchor:
    If Erl = 0 Then
        Err.Raise Err.Number, "clsAnchor.Anchor" & vbCrLf & Err.Source, Err.Description
    Else
        Err.Raise Err.Number, "clsAnchor.Anchor Linea " & Erl & vbCrLf & Err.Source, Err.Description
    End If
End Sub



Public Property Let Form(ByVal of As Form)
    Set frm = of
End Property



Private Function HeightContainerChanges(ByVal oControl As Control) As Boolean
    Dim Cont As Object
    Dim ContTag As String

    On Error GoTo Err_HeightContainerChanges
    Set Cont = oControl.Container
    ContTag = UCase(Left(oControl.Container.Tag, 4))
    'Only Anchor Top if the container is the form or the container is anchored Top
    If TypeOf Cont Is Form Then
        HeightContainerChanges = True
    Else
        HeightContainerChanges = (InStr(ContTag, "T") > 0 And InStr(ContTag, "B") > 0)
        If HeightContainerChanges Then
            HeightContainerChanges = HeightContainerChanges(Cont)
        End If
    End If


    Exit Function
Err_HeightContainerChanges:
    If Erl = 0 Then
        Err.Raise Err.Number, "clsAnchor.HeightContainerChanges" & vbCrLf & Err.Source, Err.Description
    Else
        Err.Raise Err.Number, "clsAnchor.HeightContainerChanges Linea " & Erl & vbCrLf & Err.Source, Err.Description
    End If
End Function

Private Function WidthContainerChanges(ByVal oControl As Control) As Boolean
    Dim Cont As Object
    Dim ContTag As String

    On Error GoTo Err_WidthContainerChanges
    Set Cont = oControl.Container
    ContTag = UCase(Left(oControl.Container.Tag, 4))
    'Only Anchor Top if the container is the form or the container is anchored Top
    If TypeOf Cont Is Form Then
        WidthContainerChanges = True
    Else
        WidthContainerChanges = (InStr(ContTag, "L") > 0 And InStr(ContTag, "R") > 0)
        If WidthContainerChanges Then
            WidthContainerChanges = WidthContainerChanges(Cont)
        End If
    End If

    Exit Function
Err_WidthContainerChanges:
    If Erl = 0 Then
        Err.Raise Err.Number, "clsAnchor.WidthContainerChanges" & vbCrLf & Err.Source, Err.Description
    Else
        Err.Raise Err.Number, "clsAnchor.WidthContainerChanges Linea " & Erl & vbCrLf & Err.Source, Err.Description
    End If
End Function

Public Property Get PropertyNumber() As Long

    PropertyNumber = pPropertyNumber

End Property

Public Property Let PropertyNumber(ByVal newvalue As Long)

    pPropertyNumber = newvalue

End Property

Public Property Get Separator() As String

    Separator = pSeparator

End Property

Public Property Let Separator(ByVal newvalue As String)

    pSeparator = newvalue

End Property

Private Sub Class_Initialize()

    pPropertyNumber = -1
    pSeparator = ","
    
End Sub


