VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} myForm 
   Caption         =   "顧客を選んでください"
   ClientHeight    =   7785
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8670
   OleObjectBlob   =   "myForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "myForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub myComboBox_Change()
    Call 請求書作成
        
End Sub
