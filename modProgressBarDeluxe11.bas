Attribute VB_Name = "modProgressDeluxe11"
' Code example by Rudy Alex Kohn
' Use as you like, but plz credit me for it =)
' rudyalexkohn@hotmail.com
' v1.1 (last version, i guess)
' - Example program now has better configuration
' - Added Alignment to the % display, Left, Right and Center
' - Added CustomString argument.
' (still 100% compatible)
' NOTE!!.. AutoRedraw MUST be set to TRUE!!
Option Explicit

Public Sub DrawPercent(picBox As PictureBox, lPercent As Integer, Optional ForeColor As Long = vbBlack, Optional LineColor As Long = vbBlue, Optional BackColor As Long = &H8000000F, Optional Align As AlignmentConstants = vbCenter, Optional CustomString As String = vbNullString)
' picBox = The PictureBox to use
' lPercent = Current % .. 0 to 100
' ForeColor = % Display Color (Optional)
' BackColor = BackGround Color (Optional)
' Align = Where to draw the % display (Optional, Default = Center)
' CustomString = A string to display with the %. Nice when doing diffrent tasks after each other
  
  If LenB(CustomString) <> 0 Then CustomString = CustomString & " "     ' Adds a " " if customstring ain't empty
  picBox.Scale (0, 0)-(100, 100)                                        ' Set scale
  With picBox
    .BackColor = BackColor                                              ' Set Background color
    .ForeColor = ForeColor                                              ' Sets forecolor (%)
    .Cls                                                                ' Clear
    picBox.Line (0, 0)-(lPercent, 100), LineColor, BF                   ' The line update
    Dim x As Integer
    Dim y As Integer
    Select Case Align
    Case vbCenter                                                       ' Align Center
      x = (.ScaleWidth - .TextWidth(CustomString & CStr(lPercent & " %"))) / 2
      y = (.ScaleHeight - .TextHeight(CustomString & CStr(lPercent & " %"))) / 2
    Case vbLeftJustify                                                  ' Align Left
      x = 0
      y = 0
    Case vbRightJustify                                                 ' Align Right
      x = .ScaleWidth - .TextWidth(CustomString & CStr(lPercent & " %"))
      y = .ScaleHeight - .TextHeight(CustomString & CStr(lPercent & " %"))
    End Select
    .CurrentX = x                                                       ' Draw % at desired position
    .CurrentY = y                                                       ' Draw % at desired position
    End With
    picBox.Print CustomString & CStr(lPercent) & " %"                   ' Draw CustomString (if any) and % display
End Sub
