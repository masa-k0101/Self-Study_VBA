Sub FSOSample()
    Dim myFSO As Object
    Dim myTS As Object

    Set myFSO = CreateObject("Scripting.FileSystemObject")

    Set myTS = myFSO.CreateTextFile("C:|FSOSample2.txt,", True)

    myTS.WriteLine "かんたんプログラミングExcel2003VBA応用編" & "作成日：" & Date

    myTS.Close
End sub