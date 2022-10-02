VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Data_Change 
   Caption         =   "UserForm2"
   ClientHeight    =   4752
   ClientLeft      =   100
   ClientTop       =   460
   ClientWidth     =   7040
   OleObjectBlob   =   "Sales Analysis - Userform 1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Data_Change"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBack_Click()

Dim counter As Integer

For counter = 1 To 10

    Data_Change.Controls("TextBox" & counter).Visible = True
    
Next counter

VBA.Unload Me

End Sub

Private Sub Insert_Update_Click()

Dim count_data As Integer, row As Long

If Insert_Update.Caption = "Insert" Then

    row = Cells(Rows.count, 1).End(xlUp).row + 1
    
    For count_data = 1 To Cells(1, Columns.count).End(xlToLeft).Column
    
        Cells(row, count_data).Value = Controls("TextBox" & count_data).Value
        
    Next count_data
    
    MsgBox "The record is inserted!"
    
    TextBox1.SetFocus
    
Else

    row = UserForm1.lstRecord.ListIndex + 1
    
    For count_data = 1 To Cells(1, Columns.count).End(xlToLeft).Column
    
        Cells(row, 1).Offset(0, count_data - 1) = Controls("TextBox" & count_data).Value
        
    Next count_data
    
End If

With UserForm1.lstRecord

    .ColumnCount = Cells(1, Columns.count).End(xlToLeft).Column
    
    .RowSource = Range("a1", Range("A1").End(xlToRight).End(xlDown)).Address
    
End With

For count_data = 1 To Cells(1, Columns.count).End(xlToLeft).Column

    Data_Change.Controls("TextBox" & count_data).Value = ""
    
Next count_data
    

End Sub

Private Sub month_change_Change()

TextBox4.Value = month_change.Value

End Sub

Private Sub UserForm_Initialize()

month_change.SmallChange = 1

month_change.Max = 12

month_change.Min = 1

End Sub


