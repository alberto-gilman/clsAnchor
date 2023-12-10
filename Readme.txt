'
'********************************************
' Based on http://www.freevbcode.com/ShowCode.asp?ID=6075
' The link don't works
' You can see it at https://web.archive.org/web/20210727124243/www.freevbcode.com/ShowCode.asp?ID=6075
'
.-.-Version  1.0
' Checks the control tag and resize it.
' 
' In the first 4 characters of Tag property of the control would be a combination of L, T, R and B
' If you have more values to codify in the Tag prpoerty they should be after the 4th char of the Tag property
' The method is case insensitive
'********************************************
'

In the form you want to anchor:

Private Sub Form_Resize()
    Static oAnchor As clsAnchor
On Error GoTo Err_Form_Resize
    If Not oAnchor Is Nothing Then
    Else
        Set oAnchor = New clsAnchor
        oAnchor.Form = Me
    End If

    oAnchor.Anchor


    Exit Sub
Err_Form_Resize:
    Screen.MousePointer = vbDefault
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en  Sub Form1.Form_Resize" & vbCrLf & "[" & Err.Source & "]", vbCritical
End Sub

.-Version 1.1
Added:
.- If you have more value to codify in the Tag property now you can put it in the Tag property as a list of values
If this is the case, you have to separate the diferent properties by the string contained in the Separator property of the class, the default is ","
And you have to specify the index (cero based) of the Anchor property in the PropertyNumber property of the class, the default is -1 and that value ndicates that the class works like the versión 1 and take the 4 first chars for the Anchor property

Now in the form you can use something like this:


Private Sub Form_Resize()
    Static oAnchor As clsAnchor
On Error GoTo Err_Form_Resize
    If Not oAnchor Is Nothing Then
    Else
        Set oAnchor = New clsAnchor
        'Optional, if not specified default is ","
        oAnchor.Separator = ";"
        'Mandatory, if the value of the PropertyNumber property is not specified, the separator is ignored and the operation is as in version 1.0
        oAnchor.PropertyNumber = 1
        oAnchor.Form = Me
    End If

    oAnchor.Anchor


    Exit Sub
Err_Form_Resize:
    Screen.MousePointer = vbDefault
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") en  Sub Form1.Form_Resize" & vbCrLf & "[" & Err.Source & "]", vbCritical
End Sub

in this case the tag property of the control have to be like this:
       "property one;TB;...                Anchor = TB
       ";LB"                               Anchor = LB
       "property one"                      No Anchor property
       "property one;;other property       No Anchor property

.- Lock the window update in the anchor for smoother refresh when resizing
Fixed:
.- If the form is maximized and is smallest than the original size put the minsize as the maximized size of the form
.- If a control don't has  container property, but in the tag is L, T, R or B the class try to move the control and crass asking for the conatainer

.- Version 1.1.1
Fixed:
The ListBoxes Height Property is not free, so in previous versions if you increase manually the heigth of the form, the heigth of the listboxes does not increase.
Now the ListBoxes Height property changes correctly

