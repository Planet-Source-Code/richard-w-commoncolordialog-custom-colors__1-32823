Attribute VB_Name = "RW_CommonDialogColor"
Option Explicit

Private Declare Function ChooseColor Lib "COMDLG32.DLL" Alias _
        "ChooseColorA" (Color As TCHOOSECOLOR) As Long
        Private Type TCHOOSECOLOR
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    Flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long
End Type
Public CustomColors(0 To 15) As Long
Public Function ColorDlg(hWndParent As Long, DefColor As Long, _
       Optional ShowExpDlg As Boolean = 0) As Long
   Dim I
   Dim C As Long
   Dim CC As TCHOOSECOLOR
   Dim arrayRegColors
   Dim ProductName
ProductName = App.EXEName

Dim SetOrNot
SetOrNot = GetSetting(ProductName, "cDialog", "RegColors", "Key")
If SetOrNot = "Key" Then
arrayRegColors = Split("16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215", ",")
Else
arrayRegColors = Split(GetSetting(ProductName, "cDialog", "RegColors", "Key"), ",")
End If

   For I = 0 To 15
      CustomColors(I) = arrayRegColors(I)
   Next I
   
   With CC
        
       .rgbResult = DefColor
       .hWndOwner = hWndParent
       .lpCustColors = VarPtr(CustomColors(0))
       .Flags = &H101
        
       If ShowExpDlg Then .Flags = .Flags Or &H2
        
       .lStructSize = Len(CC)
       C = ChooseColor(CC)
        
       If C Then
          ColorDlg = .rgbResult
       Else
          ColorDlg = -1
       End If
        
   End With
   
'save the custom colors
arrayRegColors = CustomColors(0)
For I = 1 To 15
arrayRegColors = arrayRegColors & "," & CustomColors(I)
Next I
SaveSetting ProductName, "cDialog", "RegColors", arrayRegColors
End Function


