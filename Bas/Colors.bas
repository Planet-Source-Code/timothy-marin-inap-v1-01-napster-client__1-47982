Attribute VB_Name = "Colors"
Public Const FColors = "0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,34,54,31,51,32,52,36,56,35,55,37,57,33,53,30,50,-1"
Public Const FCMatch = "#FFFFFF,#000000,#00007F,#009300,#FF0000,#CCCC66,#990099,#FF6633,#FFFF00,#00FC00,#CCFFFF,#00FFFF,#99CCFF,#FF00FF,#D2D2D2,#7F7F7F,#0000FF,#99CCFF,#FF0000,#FF6633,#009900,#99FF99,#00FFFF,#CCFFFF,#990099,#9933CC,#999999,#FFFFFF,#CCCC66,#FFFF99,#000000,#336666,#000000"
Public Const BColors = "40,41,42,43,44,45,46,47"
Public Const BCMatch = "#000000,#FF0000,#00FF00,#FF9900,#0000FF,#9933CC,#CCFFFF,#999999"
Function Reverse(msg)
    For i = Len(msg) To 1 Step -1
        Reverse = Reverse & Mid(msg, i, 1)
    Next
End Function
