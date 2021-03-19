Option Explicit

Type PersonalData
    PName As String
    PAge As Integer
    PDate As Date
End Type

Sub typeSample()
    Dim myData As PersonalData
    
    myData.PName = "田中博人"
    myData.PAge = 27
    myData.PDate = #4/1/1995#

    MsgBox  "氏名：" & myData.PName & vbCrLf & _
            "年齢：" & myData.PAge & vbCrLf & _
            "入社日：" & myData.PDate
End SUb