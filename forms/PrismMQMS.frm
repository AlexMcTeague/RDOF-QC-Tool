VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PrismMQMS 
   Caption         =   "Prism / MQMS"
   ClientHeight    =   5856
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   15105
   OleObjectBlob   =   "PrismMQMS.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PrismMQMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Button_ProjectDetailsEmail_Click()
    Dim Hyperlink As String
    Dim sht As Worksheet
    
    Set sht = ThisWorkbook.Worksheets("PrismMQMS")
    Hyperlink = sht.Range("Email_Details_Hyperlink").Value & "&body=Please update the Project Details for " & sht.Range("PID").Value & _
        ":%0D%0ATotal Footage: " & Me.TotalFootage.Text & "%0D%0APassings: " & Me.TotalHomesPassed.Text & _
        "%0D%0A%0D%0AThanks,%0D%0A"
    
    ActiveWorkbook.FollowHyperlink (Hyperlink)
End Sub

Private Sub Button_QCRejectEmail_Click()
    Dim Hyperlink As String
    Dim sht As Worksheet
    
    Set sht = ThisWorkbook.Worksheets("PrismMQMS")
    Hyperlink = sht.Range("Email_QC_Hyperlink").Value & "&body=1st QC has been completed for " & sht.Range("PID").Value & " and is rejected for:"

    ActiveWorkbook.FollowHyperlink (Hyperlink)
End Sub

Private Sub TotalFootage_Click()
    ' ClickToCopy Me.TotalFootage
End Sub

Private Sub UserForm_Initialize()
    ' Disabled; currently encountering breaking issues while scroll wheel is enabled
    ' EnableMouseScroll Me
End Sub

Private Sub UserForm_Terminate()
    ' DisableMouseScroll Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' DisableMouseScroll Me
End Sub

Private Sub ClickToCopy(label As label)
    Dim ClipData As Object
    Set ClipData = CreateObject("HtmlFile")
    ClipData.ParentWindow.ClipboardData.SetData "text", CStr(label.Text)
End Sub
