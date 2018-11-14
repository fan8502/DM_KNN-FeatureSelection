VERSION 5.00
Begin VB.Form Partition 
   Caption         =   "Partition"
   ClientHeight    =   9060
   ClientLeft      =   2400
   ClientTop       =   1500
   ClientWidth     =   11988
   LinkTopic       =   "Form2"
   ScaleHeight     =   9060
   ScaleWidth      =   11988
   Begin VB.CommandButton backward_k6 
      Caption         =   "K=6 / backward"
      Height          =   492
      Left            =   9600
      TabIndex        =   17
      Top             =   7560
      Width           =   1572
   End
   Begin VB.CommandButton backward_k5 
      Caption         =   "K=5 / backward"
      Height          =   492
      Left            =   9600
      TabIndex        =   16
      Top             =   6720
      Width           =   1572
   End
   Begin VB.CommandButton backward_k4 
      Caption         =   "K=4 / backward"
      Height          =   492
      Left            =   9600
      TabIndex        =   15
      Top             =   5880
      Width           =   1572
   End
   Begin VB.CommandButton forward_k6 
      Caption         =   "K=6 / forward"
      Height          =   492
      Left            =   7320
      TabIndex        =   14
      Top             =   7560
      Width           =   1572
   End
   Begin VB.CommandButton forward_k5 
      Caption         =   "K=5 / forward"
      Height          =   492
      Left            =   7320
      TabIndex        =   13
      Top             =   6720
      Width           =   1572
   End
   Begin VB.CommandButton forward_k4 
      Caption         =   "K=4 / forward"
      Height          =   492
      Left            =   7320
      TabIndex        =   12
      Top             =   5880
      Width           =   1572
   End
   Begin VB.CommandButton random 
      Caption         =   "RandomData"
      Height          =   492
      Left            =   9000
      TabIndex        =   11
      Top             =   840
      Width           =   2172
   End
   Begin VB.CommandButton six_nearst 
      Caption         =   "K=6 NN  / 5-fold"
      Height          =   612
      Left            =   9600
      TabIndex        =   10
      Top             =   3240
      Width           =   1572
   End
   Begin VB.CommandButton five_nearst 
      Caption         =   "K=5 NN / 5-fold"
      Height          =   612
      Left            =   7320
      TabIndex        =   9
      Top             =   3240
      Width           =   1572
   End
   Begin VB.CommandButton four_nearst 
      Caption         =   "K=4 NN  / 5-fold"
      Height          =   612
      Left            =   9600
      TabIndex        =   8
      Top             =   2280
      Width           =   1572
   End
   Begin VB.CommandButton backward 
      Caption         =   "K=3 / backward"
      Height          =   492
      Left            =   9600
      TabIndex        =   7
      Top             =   5040
      Width           =   1572
   End
   Begin VB.CommandButton forward 
      Caption         =   "K=3 / forward"
      Height          =   492
      Left            =   7320
      TabIndex        =   6
      Top             =   5040
      Width           =   1572
   End
   Begin VB.CommandButton three_nearst 
      Caption         =   "K=3 NN / 5-fold"
      Height          =   612
      Left            =   7320
      TabIndex        =   5
      Top             =   2280
      Width           =   1572
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7488
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   6852
   End
   Begin VB.TextBox infile 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Text            =   "yeast.txt"
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton Partition 
      Caption         =   "Read"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5160
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "ps.如果老師覺得跑太久資料夾內有附上我自己等他      確實跑完的成果截圖QQ"
      Height          =   420
      Left            =   7320
      TabIndex        =   22
      Top             =   8280
      Width           =   4032
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "#backward一個按紐約需跑2-3分鐘"
      Height          =   180
      Left            =   7440
      TabIndex        =   21
      Top             =   4560
      Width           =   2580
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "#forward一個按鈕約需跑5-6分鐘"
      Height          =   180
      Left            =   7440
      TabIndex        =   20
      Top             =   4200
      Width           =   2472
   End
   Begin VB.Label Label3 
      Caption         =   "#以下4個KNN按鈕及8個Feature Selection按鈕皆可直      接點選不需要再從新read及random資料"
      Height          =   372
      Left            =   7320
      TabIndex        =   19
      Top             =   1680
      Width           =   4092
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      Caption         =   "#Read完後請先點  選此random按鈕"
      Height          =   420
      Left            =   7440
      TabIndex        =   18
      Top             =   960
      Width           =   1392
   End
   Begin VB.Label Label5 
      Caption         =   "Data"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3360
      TabIndex        =   3
      Top             =   1080
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "Input file :"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Partition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim in_file As String, out_file As String, nstr As String
Dim attribute_name(9) As String '存八個屬性的名字，從index=1開始
Dim class_name(11) As String '存十個類別的名字，從index=1開始
Dim fileCount As Integer
Dim fileArray(1484, 10) As Variant
Dim learnArray(1484, 10) As Variant '存學習後的class的陣列
Dim fold1(297, 10) As Variant
Dim fold2(297, 10) As Variant
Dim fold3(297, 10) As Variant
Dim fold4(297, 10) As Variant
Dim fold5(296, 10) As Variant
Dim foldDis(1188, 297, 6) As Variant
Dim trainingArray() As Variant
Dim trainingArray2() As Variant
Dim trainingArray3() As Variant
Dim trainingArray4() As Variant
Dim trainingArray5() As Variant
Private Sub backward_k4_Click()
'K=3時
    Dim select_array(9) As Boolean '從index(1)開始存八個attribute
    Dim accuracy As Double
    Dim select_index(9) As Integer '紀錄選到哪幾個attribute，從1開始，共8個屬性，故宣告為9
    Dim temp_max As Double
    Dim i As Integer, k As Integer
    Dim output As Variant
    temp_max = 0
    '初始化八個屬性的是否選擇
    For i = 0 To 8
        select_array(i) = True
    Next i
    accuracy = K4Function(select_array)
    temp_max = accuracy
    List1.AddItem "Remove Attribute : " & select_index(0) & vbTab & "Accuracy : " & temp_max
'移除1個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        select_array(k) = False
        select_array(k - 1) = True
        accuracy = K4Function(select_array)
        If accuracy > temp_max Then
            select_index(1) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = True
    select_array(select_index(1)) = False
    If select_index(1) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Remove Attribute : " & select_index(1) & vbTab & "Accuracy : " & temp_max
    End If
'移除2個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = False
        Else
            select_array(k) = False
            select_array(k - 1) = True
        End If
        accuracy = K4Function(select_array)
        If accuracy > temp_max Then
            select_index(2) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = True
    select_array(select_index(1)) = False
    select_array(select_index(2)) = False
    If select_index(2) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Remove Attribute : " & select_index(2) & vbTab & "Accuracy : " & temp_max
    End If
'移除3個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = False
        Else
            select_array(k) = False
            select_array(k - 1) = True
        End If
        accuracy = K4Function(select_array)
        If accuracy > temp_max Then
            select_index(3) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = True
    select_array(select_index(1)) = False
    select_array(select_index(2)) = False
    select_array(select_index(3)) = False
    If select_index(3) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Remove Attribute : " & select_index(3) & vbTab & "Accuracy : " & temp_max
    End If
'移除4個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = False
        Else
            select_array(k) = False
            select_array(k - 1) = True
        End If
        accuracy = K4Function(select_array)
        If accuracy > temp_max Then
            select_index(4) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = True
    select_array(select_index(1)) = False
    select_array(select_index(2)) = False
    select_array(select_index(3)) = False
    select_array(select_index(4)) = False
    If select_index(4) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Remove Attribute : " & select_index(4) & vbTab & "Accuracy : " & temp_max
    End If
'移除5個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(4) + 1 Then
            select_array(k) = False
        Else
            select_array(k) = False
            select_array(k - 1) = True
        End If
        accuracy = K4Function(select_array)
        If accuracy > temp_max Then
            select_index(5) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = True
    select_array(select_index(1)) = False
    select_array(select_index(2)) = False
    select_array(select_index(3)) = False
    select_array(select_index(4)) = False
    select_array(select_index(5)) = False
    If select_index(5) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Remove Attribute : " & select_index(5) & vbTab & "Accuracy : " & temp_max
    End If
'移除6個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(4) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(5) + 1 Then
            select_array(k) = False
        Else
            select_array(k) = False
            select_array(k - 1) = True
        End If
        accuracy = K4Function(select_array)
        If accuracy > temp_max Then
            select_index(6) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = True
    select_array(select_index(1)) = False
    select_array(select_index(2)) = False
    select_array(select_index(3)) = False
    select_array(select_index(4)) = False
    select_array(select_index(5)) = False
    select_array(select_index(6)) = False
    If select_index(6) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Remove Attribute : " & select_index(6) & vbTab & "Accuracy : " & temp_max
    End If
'移除7個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(4) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(5) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(6) + 1 Then
            select_array(k) = False
        Else
            select_array(k) = False
            select_array(k - 1) = True
        End If
        accuracy = K4Function(select_array)
        If accuracy > temp_max Then
            select_index(7) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = True
    select_array(select_index(1)) = False
    select_array(select_index(2)) = False
    select_array(select_index(3)) = False
    select_array(select_index(4)) = False
    select_array(select_index(5)) = False
    select_array(select_index(6)) = False
    select_array(select_index(7)) = False
    If select_index(7) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Remove Attribute : " & select_index(7) & vbTab & "Accuracy : " & temp_max
    End If
line_end:
    List1.AddItem "K=4 , END"
    output = set_output(select_array)
End Sub
Private Sub backward_k5_Click()
'K=5時
    Dim select_array(9) As Boolean '從index(1)開始存八個attribute
    Dim accuracy As Double
    Dim select_index(9) As Integer '紀錄選到哪幾個attribute，從1開始，共8個屬性，故宣告為9
    Dim temp_max As Double
    Dim i As Integer, k As Integer
    Dim output As Variant
    temp_max = 0
    '初始化八個屬性的是否選擇
    For i = 0 To 8
        select_array(i) = True
    Next i
    accuracy = K5Function(select_array)
    temp_max = accuracy
    List1.AddItem "Remove Attribute : " & select_index(0) & vbTab & "Accuracy : " & temp_max
'移除1個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        select_array(k) = False
        select_array(k - 1) = True
        accuracy = K5Function(select_array)
        If accuracy > temp_max Then
            select_index(1) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = True
    select_array(select_index(1)) = False
    If select_index(1) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Remove Attribute : " & select_index(1) & vbTab & "Accuracy : " & temp_max
    End If
'移除2個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = False
        Else
            select_array(k) = False
            select_array(k - 1) = True
        End If
        accuracy = K5Function(select_array)
        If accuracy > temp_max Then
            select_index(2) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = True
    select_array(select_index(1)) = False
    select_array(select_index(2)) = False
    If select_index(2) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Remove Attribute : " & select_index(2) & vbTab & "Accuracy : " & temp_max
    End If
'移除3個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = False
        Else
            select_array(k) = False
            select_array(k - 1) = True
        End If
        accuracy = K5Function(select_array)
        If accuracy > temp_max Then
            select_index(3) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = True
    select_array(select_index(1)) = False
    select_array(select_index(2)) = False
    select_array(select_index(3)) = False
    If select_index(3) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Remove Attribute : " & select_index(3) & vbTab & "Accuracy : " & temp_max
    End If
'移除4個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = False
        Else
            select_array(k) = False
            select_array(k - 1) = True
        End If
        accuracy = K5Function(select_array)
        If accuracy > temp_max Then
            select_index(4) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = True
    select_array(select_index(1)) = False
    select_array(select_index(2)) = False
    select_array(select_index(3)) = False
    select_array(select_index(4)) = False
    If select_index(4) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Remove Attribute : " & select_index(4) & vbTab & "Accuracy : " & temp_max
    End If
'移除5個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(4) + 1 Then
            select_array(k) = False
        Else
            select_array(k) = False
            select_array(k - 1) = True
        End If
        accuracy = K5Function(select_array)
        If accuracy > temp_max Then
            select_index(5) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = True
    select_array(select_index(1)) = False
    select_array(select_index(2)) = False
    select_array(select_index(3)) = False
    select_array(select_index(4)) = False
    select_array(select_index(5)) = False
    If select_index(5) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Remove Attribute : " & select_index(5) & vbTab & "Accuracy : " & temp_max
    End If
'移除6個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(4) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(5) + 1 Then
            select_array(k) = False
        Else
            select_array(k) = False
            select_array(k - 1) = True
        End If
        accuracy = K5Function(select_array)
        If accuracy > temp_max Then
            select_index(6) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = True
    select_array(select_index(1)) = False
    select_array(select_index(2)) = False
    select_array(select_index(3)) = False
    select_array(select_index(4)) = False
    select_array(select_index(5)) = False
    select_array(select_index(6)) = False
    If select_index(6) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Remove Attribute : " & select_index(6) & vbTab & "Accuracy : " & temp_max
    End If
'移除7個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(4) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(5) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(6) + 1 Then
            select_array(k) = False
        Else
            select_array(k) = False
            select_array(k - 1) = True
        End If
        accuracy = K5Function(select_array)
        If accuracy > temp_max Then
            select_index(7) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = True
    select_array(select_index(1)) = False
    select_array(select_index(2)) = False
    select_array(select_index(3)) = False
    select_array(select_index(4)) = False
    select_array(select_index(5)) = False
    select_array(select_index(6)) = False
    select_array(select_index(7)) = False
    If select_index(7) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Remove Attribute : " & select_index(7) & vbTab & "Accuracy : " & temp_max
    End If
line_end:
    List1.AddItem "K=5 , END"
    output = set_output(select_array)
End Sub
Public Function set_output(final_select_array)
    Dim final_set As String, i As Integer
    For i = 1 To 8
        If final_select_array(i) <> 0 Then '表示第i個屬性有被選擇，例如(0,1,0,0)為第2個屬性被選擇
            final_set = final_set & i & "," '將被選的屬性名稱丟到同一個字串內
        End If
    Next i
    final_set = Left(final_set, Len(final_set) - 1)
    List1.AddItem "final set: { " & final_set & " }"
    List1.AddItem "----------------------------------"
End Function
Private Sub backward_k6_Click()
'K=6時
    Dim select_array(9) As Boolean '從index(1)開始存八個attribute
    Dim accuracy As Double
    Dim select_index(9) As Integer '紀錄選到哪幾個attribute，從1開始，共8個屬性，故宣告為9
    Dim temp_max As Double
    Dim i As Integer, k As Integer
    Dim output As Variant
    temp_max = 0
    '初始化八個屬性的是否選擇
    For i = 0 To 8
        select_array(i) = True
    Next i
    accuracy = K6Function(select_array)
    temp_max = accuracy
    List1.AddItem "Remove Attribute : " & select_index(0) & vbTab & "Accuracy : " & temp_max
'移除1個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        select_array(k) = False
        select_array(k - 1) = True
        accuracy = K6Function(select_array)
        If accuracy > temp_max Then
            select_index(1) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = True
    select_array(select_index(1)) = False
    If select_index(1) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Remove Attribute : " & select_index(1) & vbTab & "Accuracy : " & temp_max
    End If
'移除2個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = False
        Else
            select_array(k) = False
            select_array(k - 1) = True
        End If
        accuracy = K6Function(select_array)
        If accuracy > temp_max Then
            select_index(2) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = True
    select_array(select_index(1)) = False
    select_array(select_index(2)) = False
    If select_index(2) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Remove Attribute : " & select_index(2) & vbTab & "Accuracy : " & temp_max
    End If
'移除3個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = False
        Else
            select_array(k) = False
            select_array(k - 1) = True
        End If
        accuracy = K6Function(select_array)
        If accuracy > temp_max Then
            select_index(3) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = True
    select_array(select_index(1)) = False
    select_array(select_index(2)) = False
    select_array(select_index(3)) = False
    If select_index(3) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Remove Attribute : " & select_index(3) & vbTab & "Accuracy : " & temp_max
    End If
'移除4個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = False
        Else
            select_array(k) = False
            select_array(k - 1) = True
        End If
        accuracy = K6Function(select_array)
        If accuracy > temp_max Then
            select_index(4) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = True
    select_array(select_index(1)) = False
    select_array(select_index(2)) = False
    select_array(select_index(3)) = False
    select_array(select_index(4)) = False
    If select_index(4) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Remove Attribute : " & select_index(4) & vbTab & "Accuracy : " & temp_max
    End If
'移除5個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(4) + 1 Then
            select_array(k) = False
        Else
            select_array(k) = False
            select_array(k - 1) = True
        End If
        accuracy = K6Function(select_array)
        If accuracy > temp_max Then
            select_index(5) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = True
    select_array(select_index(1)) = False
    select_array(select_index(2)) = False
    select_array(select_index(3)) = False
    select_array(select_index(4)) = False
    select_array(select_index(5)) = False
    If select_index(5) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Remove Attribute : " & select_index(5) & vbTab & "Accuracy : " & temp_max
    End If
'移除6個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(4) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(5) + 1 Then
            select_array(k) = False
        Else
            select_array(k) = False
            select_array(k - 1) = True
        End If
        accuracy = K6Function(select_array)
        If accuracy > temp_max Then
            select_index(6) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = True
    select_array(select_index(1)) = False
    select_array(select_index(2)) = False
    select_array(select_index(3)) = False
    select_array(select_index(4)) = False
    select_array(select_index(5)) = False
    select_array(select_index(6)) = False
    If select_index(6) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Remove Attribute : " & select_index(6) & vbTab & "Accuracy : " & temp_max
    End If
'移除7個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(4) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(5) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(6) + 1 Then
            select_array(k) = False
        Else
            select_array(k) = False
            select_array(k - 1) = True
        End If
        accuracy = K6Function(select_array)
        If accuracy > temp_max Then
            select_index(7) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = True
    select_array(select_index(1)) = False
    select_array(select_index(2)) = False
    select_array(select_index(3)) = False
    select_array(select_index(4)) = False
    select_array(select_index(5)) = False
    select_array(select_index(6)) = False
    select_array(select_index(7)) = False
    If select_index(7) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Remove Attribute : " & select_index(7) & vbTab & "Accuracy : " & temp_max
    End If
line_end:
    List1.AddItem "END"
    output = set_output(select_array)
End Sub

Private Sub Form_Load()
    attribute_name(1) = "mcg"
    attribute_name(2) = "gvh"
    attribute_name(3) = "alm"
    attribute_name(4) = "mit"
    attribute_name(5) = "erl"
    attribute_name(6) = "pox"
    attribute_name(7) = "vac"
    attribute_name(8) = "nuc"
    class_name(1) = "CYT"
    class_name(2) = "ERL"
    class_name(3) = "EXC"
    class_name(4) = "ME1"
    class_name(5) = "ME2"
    class_name(6) = "ME3"
    class_name(7) = "MIT"
    class_name(8) = "NUC"
    class_name(9) = "POX"
    class_name(10) = "VAC"
'    forward.Enabled = False
'    backward.Enabled = False
End Sub
Public Function K3Function(att() As Boolean) As Double
    Dim i As Integer, j As Integer, n As Integer, m As Integer, k As Integer, x As Integer, y As Integer, t As Integer
'-----------------------------------------------------------------------------------------------------------fold 1當testing
Dim eduDistance(1188) As Double
Dim correct(6) As Integer
Dim accuracy(6) As Single
Dim maxDis(3) As Double
Dim maxInd(3) As Double
Dim class_count(11) As Integer
Dim maxclasscount As Integer
Dim maxclass As Integer
    For n = 0 To 296
        Erase eduDistance
        For j = 0 To 1187 '一個n和其他j個data的距離
            For m = 1 To 8
                If att(m) Then
                    eduDistance(j) = eduDistance(j) + ((Val(fold1(n, m)) - Val(trainingArray(j, m))) ^ 2)
                End If
            Next m
            eduDistance(j) = eduDistance(j) ^ 0.5
        Next j
        For i = 0 To 2
            maxDis(i) = 1
        Next i
        For k = 0 To 1187
            If eduDistance(k) < maxDis(0) Then
                maxDis(2) = maxDis(1)
                maxDis(1) = maxDis(0)
                maxDis(0) = eduDistance(k)
                maxInd(2) = maxInd(1)
                maxInd(1) = maxInd(0)
                maxInd(0) = k
            ElseIf eduDistance(k) >= maxDis(0) And eduDistance(k) <= maxDis(1) Then
                maxDis(2) = maxDis(1)
                maxDis(1) = eduDistance(k)
                maxInd(2) = maxInd(1)
                maxInd(1) = k
            ElseIf eduDistance(k) >= maxDis(1) And eduDistance(k) <= maxDis(2) Then
                maxDis(2) = eduDistance(k)
                maxInd(2) = k
            End If
        Next k
        Erase class_count
        For t = 0 To 2
            Select Case trainingArray(maxInd(t), 9)
                Case "CYT"
                    class_count(1) = class_count(1) + 1
                Case "ERL"
                    class_count(2) = class_count(2) + 1
                Case "EXC"
                    class_count(3) = class_count(3) + 1
                Case "ME1"
                    class_count(4) = class_count(4) + 1
                Case "ME2"
                    class_count(5) = class_count(5) + 1
                Case "ME3"
                   class_count(6) = class_count(6) + 1
                Case "MIT"
                   class_count(7) = class_count(7) + 1
                Case "NUC"
                   class_count(8) = class_count(8) + 1
                Case "POX"
                    class_count(9) = class_count(9) + 1
                Case "VAC"
                    class_count(10) = class_count(10) + 1
            End Select
        Next t
        maxclasscount = 0
        maxclass = 0 '要預測的class值
        For m = 1 To 10
            If class_count(m) <> 0 Then
                If class_count(m) > maxclasscount Then
                    maxclasscount = class_count(m)
                    maxclass = m
                End If
            End If
        Next m
        If class_name(maxclass) = fold1(n, 9) Then
            correct(1) = correct(1) + 1
        End If
        accuracy(1) = correct(1) / 297
    Next n
'-----------------------------------------------------------------------------------------------------------fold 2當testing
Dim eduDistance2(1188) As Double
Dim maxDis2(3) As Double
Dim maxInd2(3) As Double
Dim class_count2(11) As Integer
Dim maxclasscount2 As Integer
Dim maxclass2 As Integer
    For n = 0 To 296
        Erase eduDistance2
        For j = 0 To 1187 '一個n和其他j個data的距離
            For m = 1 To 8
                If att(m) Then
                    eduDistance2(j) = eduDistance2(j) + ((Val(fold2(n, m)) - Val(trainingArray2(j, m))) ^ 2)
                End If
            Next m
            eduDistance2(j) = eduDistance2(j) ^ 0.5
        Next j

        For i = 0 To 2
            maxDis2(i) = 1
        Next i
        For k = 0 To 1187
            If eduDistance2(k) < maxDis2(0) Then
                maxDis2(2) = maxDis2(1)
                maxDis2(1) = maxDis2(0)
                maxDis2(0) = eduDistance2(k)
                maxInd2(2) = maxInd2(1)
                maxInd2(1) = maxInd2(0)
                maxInd2(0) = k
            ElseIf eduDistance2(k) >= maxDis2(0) And eduDistance2(k) <= maxDis2(1) Then
                maxDis2(2) = maxDis2(1)
                maxDis2(1) = eduDistance2(k)
                maxInd2(2) = maxInd2(1)
                maxInd2(1) = k
            ElseIf eduDistance2(k) >= maxDis2(1) And eduDistance2(k) <= maxDis2(2) Then
                maxDis2(2) = eduDistance2(k)
                maxInd2(2) = k
            End If
        Next k
        Erase class_count2
        For t = 0 To 2
            Select Case trainingArray2(maxInd2(t), 9)
                Case "CYT"
                    class_count2(1) = class_count2(1) + 1
                Case "ERL"
                    class_count2(2) = class_count2(2) + 1
                Case "EXC"
                    class_count2(3) = class_count2(3) + 1
                Case "ME1"
                    class_count2(4) = class_count2(4) + 1
                Case "ME2"
                    class_count2(5) = class_count2(5) + 1
                Case "ME3"
                   class_count2(6) = class_count2(6) + 1
                Case "MIT"
                   class_count2(7) = class_count2(7) + 1
                Case "NUC"
                   class_count2(8) = class_count2(8) + 1
                Case "POX"
                    class_count2(9) = class_count2(9) + 1
                Case "VAC"
                    class_count2(10) = class_count2(10) + 1
            End Select
        Next t

        maxclasscount2 = 0
        maxclass2 = 0 '要預測的class值
        For m = 1 To 10
            If class_count2(m) <> 0 Then
                If class_count2(m) > maxclasscount2 Then
                    maxclasscount2 = class_count2(m)
                    maxclass2 = m
                End If
            End If
        Next m
        If class_name(maxclass2) = fold2(n, 9) Then
            correct(2) = correct(2) + 1
        End If
        accuracy(2) = correct(2) / 297
    Next n
'-----------------------------------------------------------------------------------------------------------fold 3當testing
Dim eduDistance3(1188) As Double
Dim maxDis3(3) As Double
Dim maxInd3(3) As Double
Dim class_count3(11) As Integer
Dim maxclasscount3 As Integer
Dim maxclass3 As Integer
    For n = 0 To 296
        Erase eduDistance3
        For j = 0 To 1187 '一個n和其他j個data的距離
            For m = 1 To 8
                If att(m) Then
                    eduDistance3(j) = eduDistance3(j) + ((Val(fold3(n, m)) - Val(trainingArray3(j, m))) ^ 2)
                End If
            Next m
            eduDistance3(j) = eduDistance3(j) ^ 0.5
        Next j

        For i = 0 To 2
            maxDis3(i) = 1
        Next i
        For k = 0 To 1187
            If eduDistance3(k) < maxDis3(0) Then
                maxDis3(2) = maxDis3(1)
                maxDis3(1) = maxDis3(0)
                maxDis3(0) = eduDistance3(k)
                maxInd3(2) = maxInd3(1)
                maxInd3(1) = maxInd3(0)
                maxInd3(0) = k
            ElseIf eduDistance3(k) >= maxDis3(0) And eduDistance3(k) <= maxDis3(1) Then
                maxDis3(2) = maxDis3(1)
                maxDis3(1) = eduDistance3(k)
                maxInd3(2) = maxInd3(1)
                maxInd3(1) = k
            ElseIf eduDistance3(k) >= maxDis3(1) And eduDistance3(k) <= maxDis3(2) Then
                maxDis3(2) = eduDistance3(k)
                maxInd3(2) = k
            End If
        Next k
        Erase class_count3
        For t = 0 To 2
            Select Case trainingArray3(maxInd3(t), 9)
                Case "CYT"
                    class_count3(1) = class_count3(1) + 1
                Case "ERL"
                    class_count3(2) = class_count3(2) + 1
                Case "EXC"
                    class_count3(3) = class_count3(3) + 1
                Case "ME1"
                    class_count3(4) = class_count3(4) + 1
                Case "ME2"
                    class_count3(5) = class_count3(5) + 1
                Case "ME3"
                   class_count3(6) = class_count3(6) + 1
                Case "MIT"
                   class_count3(7) = class_count3(7) + 1
                Case "NUC"
                   class_count3(8) = class_count3(8) + 1
                Case "POX"
                    class_count3(9) = class_count3(9) + 1
                Case "VAC"
                    class_count3(10) = class_count3(10) + 1
            End Select
        Next t

        maxclasscount3 = 0
        maxclass3 = 0 '要預測的class值
        For m = 1 To 10
            If class_count3(m) <> 0 Then
                If class_count3(m) > maxclasscount3 Then
                    maxclasscount3 = class_count3(m)
                    maxclass3 = m
                End If
            End If
        Next m
        If class_name(maxclass3) = fold3(n, 9) Then
            correct(3) = correct(3) + 1
        End If
        accuracy(3) = correct(3) / 297
    Next n
'-----------------------------------------------------------------------------------------------------------fold 4當testing
Dim eduDistance4(1188) As Double
Dim maxDis4(3) As Double
Dim maxInd4(3) As Double
Dim class_count4(11) As Integer
Dim maxclasscount4 As Integer
Dim maxclass4 As Integer
    For n = 0 To 296
        Erase eduDistance4
        For j = 0 To 1187 '一個n和其他j個data的距離
            For m = 1 To 8
                If att(m) Then
                    eduDistance4(j) = eduDistance4(j) + ((Val(fold4(n, m)) - Val(trainingArray4(j, m))) ^ 2)
                End If
            Next m
            eduDistance4(j) = eduDistance4(j) ^ 0.5
        Next j

        For i = 0 To 2
            maxDis4(i) = 1
        Next i
        For k = 0 To 1187
            If eduDistance4(k) < maxDis4(0) Then
                maxDis4(2) = maxDis4(1)
                maxDis4(1) = maxDis4(0)
                maxDis4(0) = eduDistance4(k)
                maxInd4(2) = maxInd4(1)
                maxInd4(1) = maxInd4(0)
                maxInd4(0) = k
            ElseIf eduDistance4(k) >= maxDis4(0) And eduDistance4(k) <= maxDis4(1) Then
                maxDis4(2) = maxDis4(1)
                maxDis4(1) = eduDistance4(k)
                maxInd4(2) = maxInd4(1)
                maxInd4(1) = k
            ElseIf eduDistance4(k) >= maxDis4(1) And eduDistance4(k) <= maxDis4(2) Then
                maxDis4(2) = eduDistance4(k)
                maxInd4(2) = k
            End If
        Next k
        Erase class_count4
        For t = 0 To 2
            Select Case trainingArray4(maxInd4(t), 9)
                Case "CYT"
                    class_count4(1) = class_count4(1) + 1
                Case "ERL"
                    class_count4(2) = class_count4(2) + 1
                Case "EXC"
                    class_count4(3) = class_count4(3) + 1
                Case "ME1"
                    class_count4(4) = class_count4(4) + 1
                Case "ME2"
                    class_count4(5) = class_count4(5) + 1
                Case "ME3"
                   class_count4(6) = class_count4(6) + 1
                Case "MIT"
                   class_count4(7) = class_count4(7) + 1
                Case "NUC"
                   class_count4(8) = class_count4(8) + 1
                Case "POX"
                    class_count4(9) = class_count4(9) + 1
                Case "VAC"
                    class_count4(10) = class_count4(10) + 1
            End Select
        Next t

        maxclasscount4 = 0
        maxclass4 = 0 '要預測的class值
        For m = 1 To 10
            If class_count4(m) <> 0 Then
                If class_count4(m) > maxclasscount4 Then
                    maxclasscount4 = class_count(m)
                    maxclass4 = m
                End If
            End If
        Next m
        If class_name(maxclass4) = fold4(n, 9) Then
            correct(4) = correct(4) + 1
        End If
        accuracy(4) = correct(4) / 297
    Next n
'-----------------------------------------------------------------------------------------------------------fold 5當testing
Dim eduDistance5(1188) As Double
Dim maxDis5(3) As Double
Dim maxInd5(3) As Double
Dim class_count5(11) As Integer
Dim maxclasscount5 As Integer
Dim maxclass5 As Integer
    For n = 0 To 296
        Erase eduDistance5
        For j = 0 To 1187 '一個n和其他j個data的距離
            For m = 1 To 8
                If att(m) Then
                    eduDistance5(j) = eduDistance5(j) + ((Val(fold5(n, m)) - Val(trainingArray5(j, m))) ^ 2)
                End If
            Next m
            eduDistance5(j) = eduDistance5(j) ^ 0.5
        Next j

        For i = 0 To 2
            maxDis5(i) = 1
        Next i
        For k = 0 To 1187
            If eduDistance5(k) < maxDis5(0) Then
                maxDis5(2) = maxDis5(1)
                maxDis5(1) = maxDis5(0)
                maxDis5(0) = eduDistance5(k)
                maxInd5(2) = maxInd5(1)
                maxInd5(1) = maxInd5(0)
                maxInd5(0) = k
            ElseIf eduDistance5(k) >= maxDis5(0) And eduDistance5(k) <= maxDis5(1) Then
                maxDis5(2) = maxDis5(1)
                maxDis5(1) = eduDistance5(k)
                maxInd5(2) = maxInd5(1)
                maxInd5(1) = k
            ElseIf eduDistance5(k) >= maxDis5(1) And eduDistance5(k) <= maxDis5(2) Then
                maxDis5(2) = eduDistance5(k)
                maxInd5(2) = k
            End If
        Next k
        Erase class_count5
        For t = 0 To 2
            Select Case trainingArray5(maxInd5(t), 9)
                Case "CYT"
                    class_count5(1) = class_count5(1) + 1
                Case "ERL"
                    class_count5(2) = class_count5(2) + 1
                Case "EXC"
                    class_count5(3) = class_count5(3) + 1
                Case "ME1"
                    class_count5(4) = class_count5(4) + 1
                Case "ME2"
                    class_count5(5) = class_count5(5) + 1
                Case "ME3"
                   class_count5(6) = class_count5(6) + 1
                Case "MIT"
                   class_count5(7) = class_count5(7) + 1
                Case "NUC"
                   class_count5(8) = class_count5(8) + 1
                Case "POX"
                    class_count5(9) = class_count5(9) + 1
                Case "VAC"
                    class_count5(10) = class_count5(10) + 1
            End Select
        Next t
        maxclasscount5 = 0
        maxclass5 = 0 '要預測的class值
        For m = 1 To 10
            If class_count5(m) <> 0 Then
                If class_count5(m) > maxclasscount5 Then
                    maxclasscount5 = class_count5(m)
                    maxclass5 = m
                End If
            End If
        Next m
        If class_name(maxclass5) = fold5(n, 9) Then
            correct(5) = correct(5) + 1
        End If
        accuracy(5) = correct(5) / 297
    Next n
Dim average As Single
average = (accuracy(1) + accuracy(2) + accuracy(3) + accuracy(4) + accuracy(5)) / 5
K3Function = average
End Function
Public Function K4Function(att() As Boolean) As Double
    Dim i As Integer, j As Integer, n As Integer, m As Integer, k As Integer, x As Integer, y As Integer, t As Integer
'-----------------------------------------------------------------------------------------------------------fold 1當testing
Dim eduDistance(1188) As Double
Dim maxDis(4) As Double
Dim maxInd(4) As Double
Dim class_count(11) As Integer
Dim maxclasscount As Integer
Dim maxclass As Integer
Dim correct(6) As Integer
Dim accuracy(6) As Single
    For n = 0 To 296
        Erase eduDistance
        For j = 0 To 1187
            For m = 1 To 8
                If att(m) Then
                    eduDistance(j) = eduDistance(j) + ((Val(fold1(n, m)) - Val(trainingArray(j, m))) ^ 2)
                End If
            Next m
            eduDistance(j) = eduDistance(j) ^ 0.5
        Next j
        For i = 0 To 3
            maxDis(i) = 1
        Next i
        For k = 0 To 1187
            If eduDistance(k) < maxDis(0) Then
                maxDis(3) = maxDis(2)
                maxDis(2) = maxDis(1)
                maxDis(1) = maxDis(0)
                maxDis(0) = eduDistance(k)
                maxInd(3) = maxInd(2)
                maxInd(2) = maxInd(1)
                maxInd(1) = maxInd(0)
                maxInd(0) = k
            ElseIf eduDistance(k) >= maxDis(0) And eduDistance(k) <= maxDis(1) Then
                maxDis(3) = maxDis(2)
                maxDis(2) = maxDis(1)
                maxDis(1) = eduDistance(k)
                maxInd(3) = maxInd(2)
                maxInd(2) = maxInd(1)
                maxInd(1) = k
            ElseIf eduDistance(k) >= maxDis(1) And eduDistance(k) <= maxDis(2) Then
                maxDis(3) = maxDis(2)
                maxDis(2) = eduDistance(k)
                maxInd(3) = maxInd(2)
                maxInd(2) = k
            ElseIf eduDistance(k) >= maxDis(2) And eduDistance(k) <= maxDis(3) Then
                maxDis(3) = eduDistance(k)
                maxInd(3) = k
            End If
        Next k
        Erase class_count
        For t = 0 To 3
            Select Case trainingArray(maxInd(t), 9)
                Case "CYT"
                    class_count(1) = class_count(1) + 1
                Case "ERL"
                    class_count(2) = class_count(2) + 1
                Case "EXC"
                    class_count(3) = class_count(3) + 1
                Case "ME1"
                    class_count(4) = class_count(4) + 1
                Case "ME2"
                    class_count(5) = class_count(5) + 1
                Case "ME3"
                   class_count(6) = class_count(6) + 1
                Case "MIT"
                   class_count(7) = class_count(7) + 1
                Case "NUC"
                   class_count(8) = class_count(8) + 1
                Case "POX"
                    class_count(9) = class_count(9) + 1
                Case "VAC"
                    class_count(10) = class_count(10) + 1
            End Select
        Next t
        maxclasscount = 0
        maxclass = 0 '要預測的class值
        For m = 1 To 10
            If class_count(m) > maxclasscount Then
                maxclasscount = class_count(m)
                maxclass = m
            End If
        Next m
        If class_name(maxclass) = fold1(n, 9) Then
            correct(1) = correct(1) + 1
        End If
        accuracy(1) = correct(1) / 297
    Next n
'-----------------------------------------------------------------------------------------------------------fold 2當testing
Dim eduDistance2(1188) As Double
Dim maxDis2(4) As Double
Dim maxInd2(4) As Double
Dim class_count2(11) As Integer
Dim maxclasscount2 As Integer
Dim maxclass2 As Integer
    For n = 0 To 296
        Erase eduDistance2
        For j = 0 To 1187 '一個n和其他j個data的距離
            For m = 1 To 8
                If att(m) Then
                    eduDistance2(j) = eduDistance2(j) + ((Val(fold2(n, m)) - Val(trainingArray2(j, m))) ^ 2)
                End If
            Next m
            eduDistance2(j) = eduDistance2(j) ^ 0.5
        Next j

        For i = 0 To 3
            maxDis2(i) = 1
        Next i
        For k = 0 To 1187
            If eduDistance2(k) < maxDis2(0) Then
                maxDis2(3) = maxDis2(2)
                maxDis2(2) = maxDis2(1)
                maxDis2(1) = maxDis2(0)
                maxDis2(0) = eduDistance2(k)
                maxInd2(3) = maxInd2(2)
                maxInd2(2) = maxInd2(1)
                maxInd2(1) = maxInd2(0)
                maxInd2(0) = k
            ElseIf eduDistance2(k) >= maxDis2(0) And eduDistance2(k) <= maxDis2(1) Then
                maxDis2(3) = maxDis2(2)
                maxDis2(2) = maxDis2(1)
                maxDis2(1) = eduDistance2(k)
                maxInd2(3) = maxInd2(2)
                maxInd2(2) = maxInd2(1)
                maxInd2(1) = k
            ElseIf eduDistance2(k) >= maxDis2(1) And eduDistance2(k) <= maxDis2(2) Then
                maxDis2(3) = maxDis2(2)
                maxDis2(2) = eduDistance2(k)
                maxInd2(3) = maxInd2(2)
                maxInd2(2) = k
            ElseIf eduDistance2(k) >= maxDis2(2) And eduDistance2(k) <= maxDis2(3) Then
                maxDis2(3) = eduDistance2(k)
                maxInd2(3) = k
            End If
        Next k
        Erase class_count2
        For t = 0 To 3
            Select Case trainingArray2(maxInd2(t), 9)
                Case "CYT"
                    class_count2(1) = class_count2(1) + 1
                Case "ERL"
                    class_count2(2) = class_count2(2) + 1
                Case "EXC"
                    class_count2(3) = class_count2(3) + 1
                Case "ME1"
                    class_count2(4) = class_count2(4) + 1
                Case "ME2"
                    class_count2(5) = class_count2(5) + 1
                Case "ME3"
                   class_count2(6) = class_count2(6) + 1
                Case "MIT"
                   class_count2(7) = class_count2(7) + 1
                Case "NUC"
                   class_count2(8) = class_count2(8) + 1
                Case "POX"
                    class_count2(9) = class_count2(9) + 1
                Case "VAC"
                    class_count2(10) = class_count2(10) + 1
            End Select
        Next t
        maxclasscount2 = 0
        maxclass2 = 0 '要預測的class值
        For m = 1 To 10
            If class_count2(m) > maxclasscount2 Then
                maxclasscount2 = class_count2(m)
                maxclass2 = m
            End If
        Next m
        If class_name(maxclass2) = fold2(n, 9) Then
            correct(2) = correct(2) + 1
        End If
        accuracy(2) = correct(2) / 297
    Next n
'-----------------------------------------------------------------------------------------------------------fold 3當testing
Dim eduDistance3(1188) As Double
Dim maxDis3(4) As Double
Dim maxInd3(4) As Double
Dim class_count3(11) As Integer
Dim maxclasscount3 As Integer
Dim maxclass3 As Integer
    For n = 0 To 296
        Erase eduDistance3
        For j = 0 To 1187 '一個n和其他j個data的距離
            For m = 1 To 8
                If att(m) Then
                    eduDistance3(j) = eduDistance3(j) + ((Val(fold3(n, m)) - Val(trainingArray3(j, m))) ^ 2)
                End If
            Next m
            eduDistance3(j) = eduDistance3(j) ^ 0.5
        Next j

        For i = 0 To 3
            maxDis3(i) = 1
        Next i
        For k = 0 To 1187
            If eduDistance3(k) < maxDis3(0) Then
                maxDis3(3) = maxDis3(2)
                maxDis3(2) = maxDis3(1)
                maxDis3(1) = maxDis3(0)
                maxDis3(0) = eduDistance3(k)
                maxInd3(3) = maxInd3(2)
                maxInd3(2) = maxInd3(1)
                maxInd3(1) = maxInd3(0)
                maxInd3(0) = k
            ElseIf eduDistance3(k) >= maxDis3(0) And eduDistance3(k) < maxDis3(1) Then
                maxDis3(3) = maxDis3(2)
                maxDis3(2) = maxDis3(1)
                maxDis3(1) = eduDistance3(k)
                maxInd3(3) = maxInd3(2)
                maxInd3(2) = maxInd3(1)
                maxInd3(1) = k
            ElseIf eduDistance3(k) >= maxDis3(1) And eduDistance3(k) < maxDis3(2) Then
                maxDis3(3) = maxDis3(2)
                maxDis3(2) = eduDistance3(k)
                maxInd3(3) = maxInd3(2)
                maxInd3(2) = k
            ElseIf eduDistance3(k) >= maxDis3(2) And eduDistance3(k) < maxDis3(3) Then
                maxDis3(3) = eduDistance3(k)
                maxInd3(3) = k
            End If
        Next k
        Erase class_count3
        For t = 0 To 3
            Select Case trainingArray3(maxInd3(t), 9)
                Case "CYT"
                    class_count3(1) = class_count3(1) + 1
                Case "ERL"
                    class_count3(2) = class_count3(2) + 1
                Case "EXC"
                    class_count3(3) = class_count3(3) + 1
                Case "ME1"
                    class_count3(4) = class_count3(4) + 1
                Case "ME2"
                    class_count3(5) = class_count3(5) + 1
                Case "ME3"
                   class_count3(6) = class_count3(6) + 1
                Case "MIT"
                   class_count3(7) = class_count3(7) + 1
                Case "NUC"
                   class_count3(8) = class_count3(8) + 1
                Case "POX"
                    class_count3(9) = class_count3(9) + 1
                Case "VAC"
                    class_count3(10) = class_count3(10) + 1
            End Select
        Next t

        maxclasscount3 = 0
        maxclass3 = 0 '要預測的class值
        For m = 1 To 10
            If class_count3(m) > maxclasscount3 Then
                maxclasscount3 = class_count3(m)
                maxclass3 = m
            End If
        Next m
        If class_name(maxclass3) = fold3(n, 9) Then
            correct(3) = correct(3) + 1
        End If
        accuracy(3) = correct(3) / 297
    Next n
'-----------------------------------------------------------------------------------------------------------fold 4當testing
Dim eduDistance4(1188) As Double
Dim maxDis4(4) As Double
Dim maxInd4(4) As Double
Dim class_count4(11) As Integer
Dim maxclasscount4 As Integer
Dim maxclass4 As Integer
    For n = 0 To 296
        Erase eduDistance4
        For j = 0 To 1187 '一個n和其他j個data的距離
            For m = 1 To 8
                If att(m) Then
                    eduDistance4(j) = eduDistance4(j) + ((Val(fold4(n, m)) - Val(trainingArray4(j, m))) ^ 2)
                End If
            Next m
            eduDistance4(j) = eduDistance4(j) ^ 0.5
        Next j

        For i = 0 To 3
            maxDis4(i) = 1
        Next i
        For k = 0 To 1187
            If eduDistance4(k) < maxDis4(0) Then
                maxDis4(3) = maxDis4(2)
                maxDis4(2) = maxDis4(1)
                maxDis4(1) = maxDis4(0)
                maxDis4(0) = eduDistance4(k)
                maxInd4(3) = maxInd4(2)
                maxInd4(2) = maxInd4(1)
                maxInd4(1) = maxInd4(0)
                maxInd4(0) = k
            ElseIf eduDistance4(k) >= maxDis4(0) And eduDistance4(k) < maxDis4(1) Then
                maxDis4(3) = maxDis4(2)
                maxDis4(2) = maxDis4(1)
                maxDis4(1) = eduDistance4(k)
                maxInd4(3) = maxInd4(2)
                maxInd4(2) = maxInd4(1)
                maxInd4(1) = k
            ElseIf eduDistance4(k) >= maxDis4(1) And eduDistance4(k) < maxDis4(2) Then
                maxDis4(3) = maxDis4(2)
                maxDis4(2) = eduDistance4(k)
                maxInd4(3) = maxInd4(2)
                maxInd4(2) = k
            ElseIf eduDistance4(k) >= maxDis4(2) And eduDistance4(k) < maxDis4(3) Then
                maxDis4(3) = eduDistance4(k)
                maxInd4(3) = k
            End If
        Next k
        Erase class_count4
        For t = 0 To 3
            Select Case trainingArray4(maxInd4(t), 9)
                Case "CYT"
                    class_count4(1) = class_count4(1) + 1
                Case "ERL"
                    class_count4(2) = class_count4(2) + 1
                Case "EXC"
                    class_count4(3) = class_count4(3) + 1
                Case "ME1"
                    class_count4(4) = class_count4(4) + 1
                Case "ME2"
                    class_count4(5) = class_count4(5) + 1
                Case "ME3"
                   class_count4(6) = class_count4(6) + 1
                Case "MIT"
                   class_count4(7) = class_count4(7) + 1
                Case "NUC"
                   class_count4(8) = class_count4(8) + 1
                Case "POX"
                    class_count4(9) = class_count4(9) + 1
                Case "VAC"
                    class_count4(10) = class_count4(10) + 1
            End Select
        Next t

        maxclasscount4 = 0
        maxclass4 = 0 '要預測的class值
        For m = 1 To 10
            If class_count4(m) > maxclasscount4 Then
                maxclasscount4 = class_count(m)
                maxclass4 = m
            End If
        Next m
        If class_name(maxclass4) = fold4(n, 9) Then
            correct(4) = correct(4) + 1
        End If
        accuracy(4) = correct(4) / 297
    Next n
'-----------------------------------------------------------------------------------------------------------fold 5當testing
Dim eduDistance5(1188) As Double
Dim maxDis5(4) As Double
Dim maxInd5(4) As Double
Dim class_count5(11) As Integer
Dim maxclasscount5 As Integer
Dim maxclass5 As Integer
    For n = 0 To 296
        Erase eduDistance5
        For j = 0 To 1187 '一個n和其他j個data的距離
            For m = 1 To 8
                If att(m) Then
                    eduDistance5(j) = eduDistance5(j) + ((Val(fold5(n, m)) - Val(trainingArray5(j, m))) ^ 2)
                End If
            Next m
            eduDistance5(j) = eduDistance5(j) ^ 0.5
        Next j

        For i = 0 To 3
            maxDis5(i) = 1
        Next i
        For k = 0 To 1187
            If eduDistance5(k) < maxDis5(0) Then
                maxDis5(3) = maxDis5(2)
                maxDis5(2) = maxDis5(1)
                maxDis5(1) = maxDis5(0)
                maxDis5(0) = eduDistance5(k)
                maxInd5(3) = maxInd5(2)
                maxInd5(2) = maxInd5(1)
                maxInd5(1) = maxInd5(0)
                maxInd5(0) = k
            ElseIf eduDistance5(k) >= maxDis5(0) And eduDistance5(k) < maxDis5(1) Then
                maxDis5(3) = maxDis5(2)
                maxDis5(2) = maxDis5(1)
                maxDis5(1) = eduDistance5(k)
                maxInd5(3) = maxInd5(2)
                maxInd5(2) = maxInd5(1)
                maxInd5(1) = k
            ElseIf eduDistance5(k) >= maxDis5(1) And eduDistance5(k) < maxDis5(2) Then
                maxDis5(3) = maxDis5(2)
                maxDis5(2) = eduDistance5(k)
                maxInd5(3) = maxInd5(2)
                maxInd5(2) = k
            ElseIf eduDistance5(k) >= maxDis5(2) And eduDistance5(k) < maxDis5(3) Then
                maxDis5(3) = eduDistance5(k)
                maxInd5(3) = k
            End If
        Next k
        Erase class_count5
        For t = 0 To 3
            Select Case trainingArray5(maxInd5(t), 9)
                Case "CYT"
                    class_count5(1) = class_count5(1) + 1
                Case "ERL"
                    class_count5(2) = class_count5(2) + 1
                Case "EXC"
                    class_count5(3) = class_count5(3) + 1
                Case "ME1"
                    class_count5(4) = class_count5(4) + 1
                Case "ME2"
                    class_count5(5) = class_count5(5) + 1
                Case "ME3"
                   class_count5(6) = class_count5(6) + 1
                Case "MIT"
                   class_count5(7) = class_count5(7) + 1
                Case "NUC"
                   class_count5(8) = class_count5(8) + 1
                Case "POX"
                    class_count5(9) = class_count5(9) + 1
                Case "VAC"
                    class_count5(10) = class_count5(10) + 1
            End Select
        Next t
        maxclasscount5 = 0
        maxclass5 = 0 '要預測的class值
        For m = 1 To 10
            If class_count5(m) > maxclasscount5 Then
                maxclasscount5 = class_count5(m)
                maxclass5 = m
            End If
        Next m
        If class_name(maxclass5) = fold5(n, 9) Then
            correct(5) = correct(5) + 1
        End If
        accuracy(5) = correct(5) / 297
    Next n
Dim average As Single
average = (accuracy(1) + accuracy(2) + accuracy(3) + accuracy(4) + accuracy(5)) / 5
K4Function = average
End Function
Public Function K5Function(att() As Boolean) As Double
    Dim i As Integer, j As Integer, n As Integer, m As Integer, k As Integer, x As Integer, y As Integer, t As Integer
'-----------------------------------------------------------------------------------------------------------fold 1當testing
Dim eduDistance(1188) As Double
Dim maxDis(5) As Double
Dim maxInd(5) As Double
Dim class_count(11) As Integer
Dim maxclasscount As Integer
Dim maxclass As Integer
Dim correct(6) As Integer
Dim accuracy(6) As Single
    For n = 0 To 296
        Erase eduDistance
        For j = 0 To 1187
            For m = 1 To 8
                If att(m) Then
                    eduDistance(j) = eduDistance(j) + ((Val(fold1(n, m)) - Val(trainingArray(j, m))) ^ 2)
                End If
            Next m
            eduDistance(j) = eduDistance(j) ^ 0.5
        Next j
        For i = 0 To 4
            maxDis(i) = 1
        Next i
        For k = 0 To 1187
            If eduDistance(k) < maxDis(0) Then
                maxDis(4) = maxDis(3)
                maxDis(3) = maxDis(2)
                maxDis(2) = maxDis(1)
                maxDis(1) = maxDis(0)
                maxDis(0) = eduDistance(k)
                maxInd(4) = maxInd(3)
                maxInd(3) = maxInd(2)
                maxInd(2) = maxInd(1)
                maxInd(1) = maxInd(0)
                maxInd(0) = k
            ElseIf eduDistance(k) >= maxDis(0) And eduDistance(k) < maxDis(1) Then
                maxDis(4) = maxDis(3)
                maxDis(3) = maxDis(2)
                maxDis(2) = maxDis(1)
                maxDis(1) = eduDistance(k)
                maxInd(4) = maxInd(3)
                maxInd(3) = maxInd(2)
                maxInd(2) = maxInd(1)
                maxInd(1) = k
            ElseIf eduDistance(k) >= maxDis(1) And eduDistance(k) <= maxDis(2) Then
                maxDis(4) = maxDis(3)
                maxDis(3) = maxDis(2)
                maxDis(2) = eduDistance(k)
                maxInd(4) = maxInd(3)
                maxInd(3) = maxInd(2)
                maxInd(2) = k
            ElseIf eduDistance(k) >= maxDis(2) And eduDistance(k) <= maxDis(3) Then
                maxDis(4) = maxDis(3)
                maxDis(3) = eduDistance(k)
                maxInd(4) = maxInd(3)
                maxInd(3) = k
            ElseIf eduDistance(k) >= maxDis(3) And eduDistance(k) <= maxDis(4) Then
                maxDis(4) = eduDistance(k)
                maxInd(4) = k
            End If
        Next k
         Erase class_count
        For t = 0 To 4
            Select Case trainingArray(maxInd(t), 9)
                Case "CYT"
                    class_count(1) = class_count(1) + 1
                Case "ERL"
                    class_count(2) = class_count(2) + 1
                Case "EXC"
                    class_count(3) = class_count(3) + 1
                Case "ME1"
                    class_count(4) = class_count(4) + 1
                Case "ME2"
                    class_count(5) = class_count(5) + 1
                Case "ME3"
                   class_count(6) = class_count(6) + 1
                Case "MIT"
                   class_count(7) = class_count(7) + 1
                Case "NUC"
                   class_count(8) = class_count(8) + 1
                Case "POX"
                    class_count(9) = class_count(9) + 1
                Case "VAC"
                    class_count(10) = class_count(10) + 1
            End Select
        Next t
        maxclasscount = 0
        maxclass = 0 '要預測的class值
        For m = 1 To 10
            If class_count(m) > maxclasscount Then
                maxclasscount = class_count(m)
                maxclass = m
            End If
        Next m

        If class_name(maxclass) = fold1(n, 9) Then
            correct(1) = correct(1) + 1
        End If
        accuracy(1) = correct(1) / 297
    Next n
'-----------------------------------------------------------------------------------------------------------fold 2當testing
Dim eduDistance2(1188) As Double
Dim maxDis2(5) As Double
Dim maxInd2(5) As Double
Dim class_count2(11) As Integer
Dim maxclasscount2 As Integer
Dim maxclass2 As Integer
    For n = 0 To 296
        Erase eduDistance2
        For j = 0 To 1187 '一個n和其他j個data的距離
            For m = 1 To 8
                If att(m) Then
                    eduDistance2(j) = eduDistance2(j) + ((Val(fold2(n, m)) - Val(trainingArray2(j, m))) ^ 2)
                End If
            Next m
            eduDistance2(j) = eduDistance2(j) ^ 0.5
        Next j
        For i = 0 To 4
            maxDis2(i) = 1
        Next i
        For k = 0 To 1187
            If eduDistance2(k) < maxDis2(0) Then
                maxDis2(4) = maxDis2(3)
                maxDis2(3) = maxDis2(2)
                maxDis2(2) = maxDis2(1)
                maxDis2(1) = maxDis2(0)
                maxDis2(0) = eduDistance2(k)
                maxInd2(4) = maxInd2(3)
                maxInd2(3) = maxInd2(2)
                maxInd2(2) = maxInd2(1)
                maxInd2(1) = maxInd2(0)
                maxInd2(0) = k
            ElseIf eduDistance2(k) >= maxDis2(0) And eduDistance2(k) < maxDis2(1) Then
                maxDis2(4) = maxDis2(3)
                maxDis2(3) = maxDis2(2)
                maxDis2(2) = maxDis2(1)
                maxDis2(1) = eduDistance2(k)
                maxInd2(4) = maxInd2(3)
                maxInd2(3) = maxInd2(2)
                maxInd2(2) = maxInd2(1)
                maxInd2(1) = k
            ElseIf eduDistance2(k) >= maxDis2(1) And eduDistance2(k) <= maxDis2(2) Then
                maxDis2(4) = maxDis2(3)
                maxDis2(3) = maxDis2(2)
                maxDis2(2) = eduDistance2(k)
                maxInd2(4) = maxInd2(3)
                maxInd2(3) = maxInd2(2)
                maxInd2(2) = k
            ElseIf eduDistance2(k) >= maxDis2(2) And eduDistance2(k) <= maxDis2(3) Then
                maxDis2(4) = maxDis2(3)
                maxDis2(3) = eduDistance2(k)
                maxInd2(4) = maxInd2(3)
                maxInd2(3) = k
            ElseIf eduDistance2(k) >= maxDis2(3) And eduDistance2(k) <= maxDis2(4) Then
                maxDis2(4) = eduDistance2(k)
                maxInd2(4) = k
            End If
        Next k
        Erase class_count2
        For t = 0 To 4
            Select Case trainingArray2(maxInd2(t), 9)
                Case "CYT"
                    class_count2(1) = class_count2(1) + 1
                Case "ERL"
                    class_count2(2) = class_count2(2) + 1
                Case "EXC"
                    class_count2(3) = class_count2(3) + 1
                Case "ME1"
                    class_count2(4) = class_count2(4) + 1
                Case "ME2"
                    class_count2(5) = class_count2(5) + 1
                Case "ME3"
                   class_count2(6) = class_count2(6) + 1
                Case "MIT"
                   class_count2(7) = class_count2(7) + 1
                Case "NUC"
                   class_count2(8) = class_count2(8) + 1
                Case "POX"
                    class_count2(9) = class_count2(9) + 1
                Case "VAC"
                    class_count2(10) = class_count2(10) + 1
            End Select
        Next t
        maxclasscount2 = 0
        maxclass2 = 0 '要預測的class值
        For m = 1 To 10
            If class_count2(m) > maxclasscount2 Then
                maxclasscount2 = class_count2(m)
                maxclass2 = m
            End If
        Next m
        If class_name(maxclass2) = fold2(n, 9) Then
            correct(2) = correct(2) + 1
        End If
        accuracy(2) = correct(2) / 297
    Next n
'-----------------------------------------------------------------------------------------------------------fold 3當testing
Dim eduDistance3(1188) As Double
Dim maxDis3(5) As Double
Dim maxInd3(5) As Double
Dim class_count3(11) As Integer
Dim maxclasscount3 As Integer
Dim maxclass3 As Integer
    For n = 0 To 296
        Erase eduDistance3
        For j = 0 To 1187 '一個n和其他j個data的距離
            For m = 1 To 8
                If att(m) Then
                    eduDistance3(j) = eduDistance3(j) + ((Val(fold3(n, m)) - Val(trainingArray3(j, m))) ^ 2)
                End If
            Next m
            eduDistance3(j) = eduDistance3(j) ^ 0.5
        Next j
        For i = 0 To 4
            maxDis3(i) = 1
        Next i
        For k = 0 To 1187
            If eduDistance3(k) < maxDis3(0) Then
                maxDis3(4) = maxDis3(3)
                maxDis3(3) = maxDis3(2)
                maxDis3(2) = maxDis3(1)
                maxDis3(1) = maxDis3(0)
                maxDis3(0) = eduDistance3(k)
                maxInd3(4) = maxInd3(3)
                maxInd3(3) = maxInd3(2)
                maxInd3(2) = maxInd3(1)
                maxInd3(1) = maxInd3(0)
                maxInd3(0) = k
            ElseIf eduDistance3(k) >= maxDis3(0) And eduDistance3(k) < maxDis3(1) Then
                maxDis3(4) = maxDis3(3)
                maxDis3(3) = maxDis3(2)
                maxDis3(2) = maxDis3(1)
                maxDis3(1) = eduDistance3(k)
                maxInd3(4) = maxInd3(3)
                maxInd3(3) = maxInd3(2)
                maxInd3(2) = maxInd3(1)
                maxInd3(1) = k
            ElseIf eduDistance3(k) >= maxDis3(1) And eduDistance3(k) < maxDis3(2) Then
                maxDis3(4) = maxDis3(3)
                maxDis3(3) = maxDis3(2)
                maxDis3(2) = eduDistance3(k)
                maxInd3(4) = maxInd3(3)
                maxInd3(3) = maxInd3(2)
                maxInd3(2) = k
            ElseIf eduDistance3(k) >= maxDis3(2) And eduDistance3(k) < maxDis3(3) Then
                maxDis3(4) = maxDis3(3)
                maxDis3(3) = eduDistance3(k)
                maxInd3(4) = maxInd3(3)
                maxInd3(3) = k
            ElseIf eduDistance3(k) >= maxDis3(3) And eduDistance3(k) < maxDis3(4) Then
                maxDis3(4) = eduDistance3(k)
                maxInd3(4) = k
            End If
        Next k
         Erase class_count3
        For t = 0 To 4
            Select Case trainingArray3(maxInd3(t), 9)
                Case "CYT"
                    class_count3(1) = class_count3(1) + 1
                Case "ERL"
                    class_count3(2) = class_count3(2) + 1
                Case "EXC"
                    class_count3(3) = class_count3(3) + 1
                Case "ME1"
                    class_count3(4) = class_count3(4) + 1
                Case "ME2"
                    class_count3(5) = class_count3(5) + 1
                Case "ME3"
                   class_count3(6) = class_count3(6) + 1
                Case "MIT"
                   class_count3(7) = class_count3(7) + 1
                Case "NUC"
                   class_count3(8) = class_count3(8) + 1
                Case "POX"
                    class_count3(9) = class_count3(9) + 1
                Case "VAC"
                    class_count3(10) = class_count3(10) + 1
            End Select
        Next t
        maxclasscount3 = 0
        maxclass3 = 0 '要預測的class值
        For m = 1 To 10
            If class_count3(m) > maxclasscount3 Then
                maxclasscount3 = class_count3(m)
                maxclass3 = m
            End If
        Next m
        If class_name(maxclass3) = fold3(n, 9) Then
            correct(3) = correct(3) + 1
        End If
        accuracy(3) = correct(3) / 297
    Next n
'-----------------------------------------------------------------------------------------------------------fold 4當testing
Dim eduDistance4(1188) As Double
Dim maxDis4(5) As Double
Dim maxInd4(5) As Double
Dim class_count4(11) As Integer
Dim maxclasscount4 As Integer
Dim maxclass4 As Integer
    For n = 0 To 296
        Erase eduDistance4
        For j = 0 To 1187 '一個n和其他j個data的距離
            For m = 1 To 8
                If att(m) Then
                    eduDistance4(j) = eduDistance4(j) + ((Val(fold4(n, m)) - Val(trainingArray4(j, m))) ^ 2)
                End If
            Next m
            eduDistance4(j) = eduDistance4(j) ^ 0.5
        Next j
        For i = 0 To 4
            maxDis4(i) = 1
        Next i
        For k = 0 To 1187
            If eduDistance4(k) < maxDis4(0) Then
                maxDis4(4) = maxDis4(3)
                maxDis4(3) = maxDis4(2)
                maxDis4(2) = maxDis4(1)
                maxDis4(1) = maxDis4(0)
                maxDis4(0) = eduDistance4(k)
                maxInd4(4) = maxInd4(3)
                maxInd4(3) = maxInd4(2)
                maxInd4(2) = maxInd4(1)
                maxInd4(1) = maxInd4(0)
                maxInd4(0) = k
            ElseIf eduDistance4(k) >= maxDis4(0) And eduDistance4(k) < maxDis4(1) Then
                maxDis4(4) = maxDis4(3)
                maxDis4(3) = maxDis4(2)
                maxDis4(2) = maxDis4(1)
                maxDis4(1) = eduDistance4(k)
                maxInd4(4) = maxInd4(3)
                maxInd4(3) = maxInd4(2)
                maxInd4(2) = maxInd4(1)
                maxInd4(1) = k
            ElseIf eduDistance4(k) >= maxDis4(1) And eduDistance4(k) < maxDis4(2) Then
                maxDis4(4) = maxDis4(3)
                maxDis4(3) = maxDis4(2)
                maxDis4(2) = eduDistance4(k)
                maxInd4(4) = maxInd4(3)
                maxInd4(3) = maxInd4(2)
                maxInd4(2) = k
            ElseIf eduDistance4(k) >= maxDis4(2) And eduDistance4(k) < maxDis4(3) Then
                maxDis4(4) = maxDis4(3)
                maxDis4(3) = eduDistance4(k)
                maxInd4(4) = maxInd4(3)
                maxInd4(3) = k
            ElseIf eduDistance4(k) >= maxDis4(3) And eduDistance4(k) < maxDis4(4) Then
                maxDis4(4) = eduDistance4(k)
                maxInd4(4) = k
            End If
        Next k
        Erase class_count4
        For t = 0 To 4
            Select Case trainingArray4(maxInd4(t), 9)
                Case "CYT"
                    class_count4(1) = class_count4(1) + 1
                Case "ERL"
                    class_count4(2) = class_count4(2) + 1
                Case "EXC"
                    class_count4(3) = class_count4(3) + 1
                Case "ME1"
                    class_count4(4) = class_count4(4) + 1
                Case "ME2"
                    class_count4(5) = class_count4(5) + 1
                Case "ME3"
                   class_count4(6) = class_count4(6) + 1
                Case "MIT"
                   class_count4(7) = class_count4(7) + 1
                Case "NUC"
                   class_count4(8) = class_count4(8) + 1
                Case "POX"
                    class_count4(9) = class_count4(9) + 1
                Case "VAC"
                    class_count4(10) = class_count4(10) + 1
            End Select
        Next t
        maxclasscount4 = 0
        maxclass4 = 0 '要預測的class值
        For m = 1 To 10
            If class_count4(m) > maxclasscount4 Then
                maxclasscount4 = class_count(m)
                maxclass4 = m
            End If
        Next m
        If class_name(maxclass4) = fold4(n, 9) Then
            correct(4) = correct(4) + 1
        End If
        accuracy(4) = correct(4) / 297
    Next n
'-----------------------------------------------------------------------------------------------------------fold 5當testing
Dim eduDistance5(1188) As Double
Dim maxDis5(5) As Double
Dim maxInd5(5) As Double
Dim class_count5(11) As Integer
Dim maxclasscount5 As Integer
Dim maxclass5 As Integer
    For n = 0 To 296
        Erase eduDistance5
        For j = 0 To 1187 '一個n和其他j個data的距離
            For m = 1 To 8
                If att(m) Then
                    eduDistance5(j) = eduDistance5(j) + ((Val(fold5(n, m)) - Val(trainingArray5(j, m))) ^ 2)
                End If
            Next m
            eduDistance5(j) = eduDistance5(j) ^ 0.5
        Next j
        For i = 0 To 4
            maxDis5(i) = 1
        Next i
        For k = 0 To 1187
            If eduDistance5(k) < maxDis5(0) Then
                maxDis5(4) = maxDis5(3)
                maxDis5(3) = maxDis5(2)
                maxDis5(2) = maxDis5(1)
                maxDis5(1) = maxDis5(0)
                maxDis5(0) = eduDistance5(k)
                maxInd5(4) = maxInd5(3)
                maxInd5(3) = maxInd5(2)
                maxInd5(2) = maxInd5(1)
                maxInd5(1) = maxInd5(0)
                maxInd5(0) = k
            ElseIf eduDistance5(k) >= maxDis5(0) And eduDistance5(k) < maxDis5(1) Then
                maxDis5(4) = maxDis5(3)
                maxDis5(3) = maxDis5(2)
                maxDis5(2) = maxDis5(1)
                maxDis5(1) = eduDistance5(k)
                maxInd5(4) = maxInd5(3)
                maxInd5(3) = maxInd5(2)
                maxInd5(2) = maxInd5(1)
                maxInd5(1) = k
            ElseIf eduDistance5(k) >= maxDis5(1) And eduDistance5(k) < maxDis5(2) Then
                maxDis5(4) = maxDis5(3)
                maxDis5(3) = maxDis5(2)
                maxDis5(2) = eduDistance5(k)
                maxInd5(4) = maxInd5(3)
                maxInd5(3) = maxInd5(2)
                maxInd5(2) = k
            ElseIf eduDistance5(k) >= maxDis5(2) And eduDistance5(k) < maxDis5(3) Then
                maxDis5(4) = maxDis5(3)
                maxDis5(3) = eduDistance5(k)
                maxInd5(4) = maxInd5(3)
                maxInd5(3) = k
            ElseIf eduDistance5(k) >= maxDis5(3) And eduDistance5(k) < maxDis5(4) Then
                maxDis5(4) = eduDistance5(k)
                maxInd5(4) = k
            End If
        Next k
        Erase class_count5
        For t = 0 To 4
            Select Case trainingArray5(maxInd5(t), 9)
                Case "CYT"
                    class_count5(1) = class_count5(1) + 1
                Case "ERL"
                    class_count5(2) = class_count5(2) + 1
                Case "EXC"
                    class_count5(3) = class_count5(3) + 1
                Case "ME1"
                    class_count5(4) = class_count5(4) + 1
                Case "ME2"
                    class_count5(5) = class_count5(5) + 1
                Case "ME3"
                   class_count5(6) = class_count5(6) + 1
                Case "MIT"
                   class_count5(7) = class_count5(7) + 1
                Case "NUC"
                   class_count5(8) = class_count5(8) + 1
                Case "POX"
                    class_count5(9) = class_count5(9) + 1
                Case "VAC"
                    class_count5(10) = class_count5(10) + 1
            End Select
        Next t
        maxclasscount5 = 0
        maxclass5 = 0 '要預測的class值
        For m = 1 To 10
            If class_count5(m) > maxclasscount5 Then
                maxclasscount5 = class_count5(m)
                maxclass5 = m
            End If
        Next m
        If class_name(maxclass5) = fold5(n, 9) Then
            correct(5) = correct(5) + 1
        End If
        accuracy(5) = correct(5) / 297
    Next n
Dim average As Single
average = (accuracy(1) + accuracy(2) + accuracy(3) + accuracy(4) + accuracy(5)) / 5
K5Function = average
End Function
Public Function K6Function(att() As Boolean) As Double
    Dim i As Integer, j As Integer, n As Integer, m As Integer, k As Integer, x As Integer, y As Integer, t As Integer
'-----------------------------------------------------------------------------------------------------------fold 1當testing
Dim eduDistance(1188) As Double
Dim maxDis(6) As Double
Dim maxInd(6) As Double
Dim class_count(11) As Integer
Dim maxclasscount As Integer
Dim maxclass As Integer
Dim correct(6) As Integer
Dim accuracy(6) As Single
    For n = 0 To 296
        Erase eduDistance
        For j = 0 To 1187
            For m = 1 To 8
                If att(m) Then
                    eduDistance(j) = eduDistance(j) + ((Val(fold1(n, m)) - Val(trainingArray(j, m))) ^ 2)
                End If
            Next m
            eduDistance(j) = eduDistance(j) ^ 0.5
        Next j
        For i = 0 To 5
            maxDis(i) = 1
        Next i
        For k = 0 To 1187
            If eduDistance(k) < maxDis(0) Then
                maxDis(5) = maxDis(4)
                maxDis(4) = maxDis(3)
                maxDis(3) = maxDis(2)
                maxDis(2) = maxDis(1)
                maxDis(1) = maxDis(0)
                maxDis(0) = eduDistance(k)
                maxInd(5) = maxInd(4)
                maxInd(4) = maxInd(3)
                maxInd(3) = maxInd(2)
                maxInd(2) = maxInd(1)
                maxInd(1) = maxInd(0)
                maxInd(0) = k
            ElseIf eduDistance(k) >= maxDis(0) And eduDistance(k) < maxDis(1) Then
                maxDis(5) = maxDis(4)
                maxDis(4) = maxDis(3)
                maxDis(3) = maxDis(2)
                maxDis(2) = maxDis(1)
                maxDis(1) = eduDistance(k)
                maxInd(5) = maxInd(4)
                maxInd(4) = maxInd(3)
                maxInd(3) = maxInd(2)
                maxInd(2) = maxInd(1)
                maxInd(1) = k
            ElseIf eduDistance(k) >= maxDis(1) And eduDistance(k) <= maxDis(2) Then
                maxDis(5) = maxDis(4)
                maxDis(4) = maxDis(3)
                maxDis(3) = maxDis(2)
                maxDis(2) = eduDistance(k)
                maxInd(5) = maxInd(4)
                maxInd(4) = maxInd(3)
                maxInd(3) = maxInd(2)
                maxInd(2) = k
            ElseIf eduDistance(k) >= maxDis(2) And eduDistance(k) <= maxDis(3) Then
                maxDis(5) = maxDis(4)
                maxDis(4) = maxDis(3)
                maxDis(3) = eduDistance(k)
                maxInd(5) = maxInd(4)
                maxInd(4) = maxInd(3)
                maxInd(3) = k
            ElseIf eduDistance(k) >= maxDis(3) And eduDistance(k) <= maxDis(4) Then
                maxDis(5) = maxDis(4)
                maxDis(4) = eduDistance(k)
                maxInd(5) = maxInd(4)
                maxInd(4) = k
            ElseIf eduDistance(k) >= maxDis(4) And eduDistance(k) <= maxDis(5) Then
                maxDis(5) = eduDistance(k)
                maxInd(5) = k
            End If
        Next k
        Erase class_count
        For t = 0 To 5
            Select Case trainingArray(maxInd(t), 9)
                Case "CYT"
                    class_count(1) = class_count(1) + 1
                Case "ERL"
                    class_count(2) = class_count(2) + 1
                Case "EXC"
                    class_count(3) = class_count(3) + 1
                Case "ME1"
                    class_count(4) = class_count(4) + 1
                Case "ME2"
                    class_count(5) = class_count(5) + 1
                Case "ME3"
                   class_count(6) = class_count(6) + 1
                Case "MIT"
                   class_count(7) = class_count(7) + 1
                Case "NUC"
                   class_count(8) = class_count(8) + 1
                Case "POX"
                    class_count(9) = class_count(9) + 1
                Case "VAC"
                    class_count(10) = class_count(10) + 1
            End Select
        Next t
        maxclasscount = 0
        maxclass = 0 '要預測的class值
        For m = 1 To 10
            If class_count(m) > maxclasscount Then
                maxclasscount = class_count(m)
                maxclass = m
            End If
        Next m

        If class_name(maxclass) = fold1(n, 9) Then
            correct(1) = correct(1) + 1
        End If
        accuracy(1) = correct(1) / 297
    Next n
'-----------------------------------------------------------------------------------------------------------fold 2當testing
Dim eduDistance2(1188) As Double
Dim maxDis2(6) As Double
Dim maxInd2(6) As Double
Dim class_count2(11) As Integer
Dim maxclasscount2 As Integer
Dim maxclass2 As Integer
    For n = 0 To 296
        Erase eduDistance2
        For j = 0 To 1187 '一個n和其他j個data的距離
            For m = 1 To 8
                If att(m) Then
                    eduDistance2(j) = eduDistance2(j) + ((Val(fold2(n, m)) - Val(trainingArray2(j, m))) ^ 2)
                End If
            Next m
            eduDistance2(j) = eduDistance2(j) ^ 0.5
        Next j
        For i = 0 To 5
            maxDis2(i) = 1
        Next i
        For k = 0 To 1187
            If eduDistance2(k) < maxDis2(0) Then
                maxDis2(5) = maxDis2(4)
                maxDis2(4) = maxDis2(3)
                maxDis2(3) = maxDis2(2)
                maxDis2(2) = maxDis2(1)
                maxDis2(1) = maxDis2(0)
                maxDis2(0) = eduDistance2(k)
                maxInd2(5) = maxInd2(4)
                maxInd2(4) = maxInd2(3)
                maxInd2(3) = maxInd2(2)
                maxInd2(2) = maxInd2(1)
                maxInd2(1) = maxInd2(0)
                maxInd2(0) = k
            ElseIf eduDistance2(k) >= maxDis2(0) And eduDistance2(k) < maxDis2(1) Then
                maxDis2(5) = maxDis2(4)
                maxDis2(4) = maxDis2(3)
                maxDis2(3) = maxDis2(2)
                maxDis2(2) = maxDis2(1)
                maxDis2(1) = eduDistance2(k)
                maxInd2(5) = maxInd2(4)
                maxInd2(4) = maxInd2(3)
                maxInd2(3) = maxInd2(2)
                maxInd2(2) = maxInd2(1)
                maxInd2(1) = k
            ElseIf eduDistance2(k) >= maxDis2(1) And eduDistance2(k) <= maxDis2(2) Then
                maxDis2(5) = maxDis2(4)
                maxDis2(4) = maxDis2(3)
                maxDis2(3) = maxDis2(2)
                maxDis2(2) = eduDistance2(k)
                maxInd2(5) = maxInd2(4)
                maxInd2(4) = maxInd2(3)
                maxInd2(3) = maxInd2(2)
                maxInd2(2) = k
            ElseIf eduDistance2(k) >= maxDis2(2) And eduDistance2(k) <= maxDis2(3) Then
                maxDis2(5) = maxDis2(4)
                maxDis2(4) = maxDis2(3)
                maxDis2(3) = eduDistance2(k)
                maxInd2(5) = maxInd2(4)
                maxInd2(4) = maxInd2(3)
                maxInd2(3) = k
            ElseIf eduDistance2(k) >= maxDis2(3) And eduDistance2(k) <= maxDis2(4) Then
                maxDis2(5) = maxDis2(4)
                maxDis2(4) = eduDistance2(k)
                maxInd2(5) = maxInd2(4)
                maxInd2(4) = k
            ElseIf eduDistance2(k) >= maxDis2(4) And eduDistance2(k) <= maxDis2(5) Then
                maxDis2(5) = eduDistance2(k)
                maxInd2(5) = k
            End If
        Next k
        Erase class_count2
        For t = 0 To 5
            Select Case trainingArray2(maxInd2(t), 9)
                Case "CYT"
                    class_count2(1) = class_count2(1) + 1
                Case "ERL"
                    class_count2(2) = class_count2(2) + 1
                Case "EXC"
                    class_count2(3) = class_count2(3) + 1
                Case "ME1"
                    class_count2(4) = class_count2(4) + 1
                Case "ME2"
                    class_count2(5) = class_count2(5) + 1
                Case "ME3"
                   class_count2(6) = class_count2(6) + 1
                Case "MIT"
                   class_count2(7) = class_count2(7) + 1
                Case "NUC"
                   class_count2(8) = class_count2(8) + 1
                Case "POX"
                    class_count2(9) = class_count2(9) + 1
                Case "VAC"
                    class_count2(10) = class_count2(10) + 1
            End Select
        Next t
        maxclasscount2 = 0
        maxclass2 = 0 '要預測的class值
        For m = 1 To 10
            If class_count2(m) > maxclasscount2 Then
                maxclasscount2 = class_count2(m)
                maxclass2 = m
            End If
        Next m
        If class_name(maxclass2) = fold2(n, 9) Then
            correct(2) = correct(2) + 1
        End If
        accuracy(2) = correct(2) / 297
    Next n
'-----------------------------------------------------------------------------------------------------------fold 3當testing
Dim eduDistance3(1188) As Double
Dim maxDis3(6) As Double
Dim maxInd3(6) As Double
Dim class_count3(11) As Integer
Dim maxclasscount3 As Integer
Dim maxclass3 As Integer
    For n = 0 To 296
        Erase eduDistance3
        For j = 0 To 1187 '一個n和其他j個data的距離
            For m = 1 To 8
                If att(m) Then
                    eduDistance3(j) = eduDistance3(j) + ((Val(fold3(n, m)) - Val(trainingArray3(j, m))) ^ 2)
                End If
            Next m
            eduDistance3(j) = eduDistance3(j) ^ 0.5
        Next j
        For i = 0 To 5
            maxDis3(i) = 1
        Next i
        For k = 0 To 1187
            If eduDistance3(k) < maxDis3(0) Then
                maxDis3(5) = maxDis3(4)
                maxDis3(4) = maxDis3(3)
                maxDis3(3) = maxDis3(2)
                maxDis3(2) = maxDis3(1)
                maxDis3(1) = maxDis3(0)
                maxDis3(0) = eduDistance3(k)
                maxInd3(5) = maxInd3(4)
                maxInd3(4) = maxInd3(3)
                maxInd3(3) = maxInd3(2)
                maxInd3(2) = maxInd3(1)
                maxInd3(1) = maxInd3(0)
                maxInd3(0) = k
            ElseIf eduDistance3(k) >= maxDis3(0) And eduDistance3(k) < maxDis3(1) Then
                maxDis3(5) = maxDis3(4)
                maxDis3(4) = maxDis3(3)
                maxDis3(3) = maxDis3(2)
                maxDis3(2) = maxDis3(1)
                maxDis3(1) = eduDistance3(k)
                maxInd3(5) = maxInd3(4)
                maxInd3(4) = maxInd3(3)
                maxInd3(3) = maxInd3(2)
                maxInd3(2) = maxInd3(1)
                maxInd3(1) = k
            ElseIf eduDistance3(k) >= maxDis3(1) And eduDistance3(k) < maxDis3(2) Then
                maxDis3(5) = maxDis3(4)
                maxDis3(4) = maxDis3(3)
                maxDis3(3) = maxDis3(2)
                maxDis3(2) = eduDistance3(k)
                maxInd3(5) = maxInd3(4)
                maxInd3(4) = maxInd3(3)
                maxInd3(3) = maxInd3(2)
                maxInd3(2) = k
            ElseIf eduDistance3(k) >= maxDis3(2) And eduDistance3(k) < maxDis3(3) Then
                maxDis3(5) = maxDis3(4)
                maxDis3(4) = maxDis3(3)
                maxDis3(3) = eduDistance3(k)
                maxInd3(5) = maxInd3(4)
                maxInd3(4) = maxInd3(3)
                maxInd3(3) = k
            ElseIf eduDistance3(k) >= maxDis3(3) And eduDistance3(k) < maxDis3(4) Then
                maxDis3(5) = maxDis3(4)
                maxDis3(4) = eduDistance3(k)
                maxInd3(5) = maxInd3(4)
                maxInd3(4) = k
            ElseIf eduDistance3(k) >= maxDis3(4) And eduDistance3(k) < maxDis3(5) Then
                maxDis3(5) = eduDistance3(k)
                maxInd3(5) = k
            End If
        Next k
        Erase class_count3
        For t = 0 To 5
            Select Case trainingArray3(maxInd3(t), 9)
                Case "CYT"
                    class_count3(1) = class_count3(1) + 1
                Case "ERL"
                    class_count3(2) = class_count3(2) + 1
                Case "EXC"
                    class_count3(3) = class_count3(3) + 1
                Case "ME1"
                    class_count3(4) = class_count3(4) + 1
                Case "ME2"
                    class_count3(5) = class_count3(5) + 1
                Case "ME3"
                   class_count3(6) = class_count3(6) + 1
                Case "MIT"
                   class_count3(7) = class_count3(7) + 1
                Case "NUC"
                   class_count3(8) = class_count3(8) + 1
                Case "POX"
                    class_count3(9) = class_count3(9) + 1
                Case "VAC"
                    class_count3(10) = class_count3(10) + 1
            End Select
        Next t
        maxclasscount3 = 0
        maxclass3 = 0 '要預測的class值
        For m = 1 To 10
            If class_count3(m) > maxclasscount3 Then
                maxclasscount3 = class_count3(m)
                maxclass3 = m
            End If
        Next m
        If class_name(maxclass3) = fold3(n, 9) Then
            correct(3) = correct(3) + 1
        End If
        accuracy(3) = correct(3) / 297
    Next n
'-----------------------------------------------------------------------------------------------------------fold 4當testing
Dim eduDistance4(1188) As Double
Dim maxDis4(6) As Double
Dim maxInd4(6) As Double
Dim class_count4(11) As Integer
Dim maxclasscount4 As Integer
Dim maxclass4 As Integer
    For n = 0 To 296
        Erase eduDistance4
        For j = 0 To 1187 '一個n和其他j個data的距離
            For m = 1 To 8
                If att(m) Then
                    eduDistance4(j) = eduDistance4(j) + ((Val(fold4(n, m)) - Val(trainingArray4(j, m))) ^ 2)
                End If
            Next m
            eduDistance4(j) = eduDistance4(j) ^ 0.5
        Next j
        For i = 0 To 5
            maxDis4(i) = 1
        Next i
        For k = 0 To 1187
            If eduDistance4(k) < maxDis4(0) Then
                maxDis4(5) = maxDis4(4)
                maxDis4(4) = maxDis4(3)
                maxDis4(3) = maxDis4(2)
                maxDis4(2) = maxDis4(1)
                maxDis4(1) = maxDis4(0)
                maxDis4(0) = eduDistance4(k)
                maxInd4(5) = maxInd4(4)
                maxInd4(4) = maxInd4(3)
                maxInd4(3) = maxInd4(2)
                maxInd4(2) = maxInd4(1)
                maxInd4(1) = maxInd4(0)
                maxInd4(0) = k
            ElseIf eduDistance4(k) >= maxDis4(0) And eduDistance4(k) < maxDis4(1) Then
                maxDis4(5) = maxDis4(4)
                maxDis4(4) = maxDis4(3)
                maxDis4(3) = maxDis4(2)
                maxDis4(2) = maxDis4(1)
                maxDis4(1) = eduDistance4(k)
                maxInd4(5) = maxInd4(4)
                maxInd4(4) = maxInd4(3)
                maxInd4(3) = maxInd4(2)
                maxInd4(2) = maxInd4(1)
                maxInd4(1) = k
            ElseIf eduDistance4(k) >= maxDis4(1) And eduDistance4(k) < maxDis4(2) Then
                maxDis4(5) = maxDis4(4)
                maxDis4(4) = maxDis4(3)
                maxDis4(3) = maxDis4(2)
                maxDis4(2) = eduDistance4(k)
                maxInd4(5) = maxInd4(4)
                maxInd4(4) = maxInd4(3)
                maxInd4(3) = maxInd4(2)
                maxInd4(2) = k
            ElseIf eduDistance4(k) >= maxDis4(2) And eduDistance4(k) < maxDis4(3) Then
                maxDis4(5) = maxDis4(4)
                maxDis4(4) = maxDis4(3)
                maxDis4(3) = eduDistance4(k)
                maxInd4(5) = maxInd4(4)
                maxInd4(4) = maxInd4(3)
                maxInd4(3) = k
            ElseIf eduDistance4(k) >= maxDis4(3) And eduDistance4(k) < maxDis4(4) Then
                maxDis4(5) = maxDis4(4)
                maxDis4(4) = eduDistance4(k)
                maxInd4(5) = maxInd4(4)
                maxInd4(4) = k
            ElseIf eduDistance4(k) >= maxDis4(4) And eduDistance4(k) < maxDis4(5) Then
                maxDis4(5) = eduDistance4(k)
                maxInd4(5) = k
            End If
        Next k
        Erase class_count4
        For t = 0 To 5
            Select Case trainingArray4(maxInd4(t), 9)
                Case "CYT"
                    class_count4(1) = class_count4(1) + 1
                Case "ERL"
                    class_count4(2) = class_count4(2) + 1
                Case "EXC"
                    class_count4(3) = class_count4(3) + 1
                Case "ME1"
                    class_count4(4) = class_count4(4) + 1
                Case "ME2"
                    class_count4(5) = class_count4(5) + 1
                Case "ME3"
                   class_count4(6) = class_count4(6) + 1
                Case "MIT"
                   class_count4(7) = class_count4(7) + 1
                Case "NUC"
                   class_count4(8) = class_count4(8) + 1
                Case "POX"
                    class_count4(9) = class_count4(9) + 1
                Case "VAC"
                    class_count4(10) = class_count4(10) + 1
            End Select
        Next t
        maxclasscount4 = 0
        maxclass4 = 0 '要預測的class值
        For m = 1 To 10
            If class_count4(m) > maxclasscount4 Then
                maxclasscount4 = class_count(m)
                maxclass4 = m
            End If
        Next m
        If class_name(maxclass4) = fold4(n, 9) Then
            correct(4) = correct(4) + 1
        End If
        accuracy(4) = correct(4) / 297
    Next n
'-----------------------------------------------------------------------------------------------------------fold 5當testing
Dim eduDistance5(1188) As Double
Dim maxDis5(6) As Double
Dim maxInd5(6) As Double
Dim class_count5(11) As Integer
Dim maxclasscount5 As Integer
Dim maxclass5 As Integer
    For n = 0 To 296
        Erase eduDistance5
        For j = 0 To 1187 '一個n和其他j個data的距離
            For m = 1 To 8
                If att(m) Then
                    eduDistance5(j) = eduDistance5(j) + ((Val(fold5(n, m)) - Val(trainingArray5(j, m))) ^ 2)
                End If
            Next m
            eduDistance5(j) = eduDistance5(j) ^ 0.5
        Next j
        For i = 0 To 5
            maxDis5(i) = 1
        Next i
        For k = 0 To 1187
            If eduDistance5(k) < maxDis5(0) Then
                maxDis5(5) = maxDis5(4)
                maxDis5(4) = maxDis5(3)
                maxDis5(3) = maxDis5(2)
                maxDis5(2) = maxDis5(1)
                maxDis5(1) = maxDis5(0)
                maxDis5(0) = eduDistance5(k)
                maxInd5(5) = maxInd5(4)
                maxInd5(4) = maxInd5(3)
                maxInd5(3) = maxInd5(2)
                maxInd5(2) = maxInd5(1)
                maxInd5(1) = maxInd5(0)
                maxInd5(0) = k
            ElseIf eduDistance5(k) >= maxDis5(0) And eduDistance5(k) < maxDis5(1) Then
                maxDis5(5) = maxDis5(4)
                maxDis5(4) = maxDis5(3)
                maxDis5(3) = maxDis5(2)
                maxDis5(2) = maxDis5(1)
                maxDis5(1) = eduDistance5(k)
                maxInd5(5) = maxInd5(4)
                maxInd5(4) = maxInd5(3)
                maxInd5(3) = maxInd5(2)
                maxInd5(2) = maxInd5(1)
                maxInd5(1) = k
            ElseIf eduDistance5(k) >= maxDis5(1) And eduDistance5(k) < maxDis5(2) Then
                maxDis5(5) = maxDis5(4)
                maxDis5(4) = maxDis5(3)
                maxDis5(3) = maxDis5(2)
                maxDis5(2) = eduDistance5(k)
                maxInd5(5) = maxInd5(4)
                maxInd5(4) = maxInd5(3)
                maxInd5(3) = maxInd5(2)
                maxInd5(2) = k
            ElseIf eduDistance5(k) >= maxDis5(2) And eduDistance5(k) < maxDis5(3) Then
                maxDis5(5) = maxDis5(4)
                maxDis5(4) = maxDis5(3)
                maxDis5(3) = eduDistance5(k)
                maxInd5(5) = maxInd5(4)
                maxInd5(4) = maxInd5(3)
                maxInd5(3) = k
            ElseIf eduDistance5(k) >= maxDis5(3) And eduDistance5(k) < maxDis5(4) Then
                maxDis5(5) = maxDis5(4)
                maxDis5(4) = eduDistance5(k)
                maxInd5(5) = maxInd5(4)
                maxInd5(4) = k
            ElseIf eduDistance5(k) >= maxDis5(4) And eduDistance5(k) < maxDis5(5) Then
                maxDis5(5) = eduDistance5(k)
                maxInd5(5) = k
            End If
        Next k
        Erase class_count5
        For t = 0 To 5
            Select Case trainingArray5(maxInd5(t), 9)
                Case "CYT"
                    class_count5(1) = class_count5(1) + 1
                Case "ERL"
                    class_count5(2) = class_count5(2) + 1
                Case "EXC"
                    class_count5(3) = class_count5(3) + 1
                Case "ME1"
                    class_count5(4) = class_count5(4) + 1
                Case "ME2"
                    class_count5(5) = class_count5(5) + 1
                Case "ME3"
                   class_count5(6) = class_count5(6) + 1
                Case "MIT"
                   class_count5(7) = class_count5(7) + 1
                Case "NUC"
                   class_count5(8) = class_count5(8) + 1
                Case "POX"
                    class_count5(9) = class_count5(9) + 1
                Case "VAC"
                    class_count5(10) = class_count5(10) + 1
            End Select
        Next t
        maxclasscount5 = 0
        maxclass5 = 0 '要預測的class值
        For m = 1 To 10
            If class_count5(m) > maxclasscount5 Then
                maxclasscount5 = class_count5(m)
                maxclass5 = m
            End If
        Next m
        If class_name(maxclass5) = fold5(n, 9) Then
            correct(5) = correct(5) + 1
        End If
        accuracy(5) = correct(5) / 297
    Next n
Dim average As Single
average = (accuracy(1) + accuracy(2) + accuracy(3) + accuracy(4) + accuracy(5)) / 5
K6Function = average
End Function
Private Sub forward_k6_Click()
'K=6時
    Dim select_array(9) As Boolean '從index(1)開始存八個attribute
    Dim accuracy As Double
    Dim select_index(9) As Integer '紀錄選到哪幾個attribute，從1開始，共8個屬性，故宣告為9
    Dim temp_max As Double
    Dim i As Integer, k As Integer
    Dim output As Variant
    temp_max = 0
    '初始化八個屬性的是否選擇
    For i = 0 To 8
        select_array(i) = False
    Next i
'選1個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        select_array(k) = True
        select_array(k - 1) = False
        accuracy = K6Function(select_array)
        If accuracy > temp_max Then
            select_index(1) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = False '把最後一個屬性初始化
    select_array(select_index(1)) = True '把紀錄到的屬性選擇起來掉 (0 0 1 0 0 0 0 0)
    If select_index(1) = 0 Then '等於0時，代表沒有K值輸入，沒有原MAX他更大的值，所以直接結束選取
        GoTo line_end
    Else
        List1.AddItem "Attribute : " & select_index(1) & vbTab & "Accuracy : " & temp_max
    End If
'選2個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = True
        Else
            select_array(k) = True
            select_array(k - 1) = False
        End If
        accuracy = K6Function(select_array)
        If accuracy > temp_max Then
            select_index(2) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = False
    select_array(select_index(1)) = True
    select_array(select_index(2)) = True
    If select_index(2) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Attribute : " & select_index(2) & vbTab & "Accuracy : " & temp_max
    End If
'選3個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = True
        Else
            select_array(k) = True
            select_array(k - 1) = False
        End If
        accuracy = K6Function(select_array)
        If accuracy > temp_max Then
            select_index(3) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = False
    select_array(select_index(1)) = True
    select_array(select_index(2)) = True
    select_array(select_index(3)) = True
    If select_index(3) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Attribute : " & select_index(3) & vbTab & "Accuracy : " & temp_max
    End If
'選4個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = True
        Else
            select_array(k) = True
            select_array(k - 1) = False
        End If
        accuracy = K6Function(select_array)
        If accuracy > temp_max Then
            select_index(4) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = False
    select_array(select_index(1)) = True
    select_array(select_index(2)) = True
    select_array(select_index(3)) = True
    select_array(select_index(4)) = True
    If select_index(4) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Attribute : " & select_index(4) & vbTab & "Accuracy : " & temp_max
    End If
'選5個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(4) + 1 Then
            select_array(k) = True
        Else
            select_array(k) = True
            select_array(k - 1) = False
        End If
        accuracy = K6Function(select_array)
        If accuracy > temp_max Then
            select_index(5) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = False
    select_array(select_index(1)) = True
    select_array(select_index(2)) = True
    select_array(select_index(3)) = True
    select_array(select_index(4)) = True
    select_array(select_index(5)) = True
    If select_index(5) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Attribute : " & select_index(5) & vbTab & "Accuracy : " & temp_max
    End If
'選6個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(4) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(5) + 1 Then
            select_array(k) = True
        Else
            select_array(k) = True
            select_array(k - 1) = False
        End If
        accuracy = K6Function(select_array)
        If accuracy > temp_max Then
            select_index(6) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = False
    select_array(select_index(1)) = True
    select_array(select_index(2)) = True
    select_array(select_index(3)) = True
    select_array(select_index(4)) = True
    select_array(select_index(5)) = True
    select_array(select_index(6)) = True
    If select_index(6) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Attribute : " & select_index(6) & vbTab & "Accuracy : " & temp_max
    End If
'選7個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(4) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(5) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(6) + 1 Then
            select_array(k) = True
        Else
            select_array(k) = True
            select_array(k - 1) = False
        End If
        accuracy = K6Function(select_array)
        If accuracy > temp_max Then
            select_index(7) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = False
    select_array(select_index(1)) = True
    select_array(select_index(2)) = True
    select_array(select_index(3)) = True
    select_array(select_index(4)) = True
    select_array(select_index(5)) = True
    select_array(select_index(6)) = True
    select_array(select_index(7)) = True
    If select_index(7) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Attribute : " & select_index(7) & vbTab & "Accuracy : " & temp_max
    End If
'選8個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(4) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(5) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(6) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(7) + 1 Then
            select_array(k) = True
        Else
            select_array(k) = True
            select_array(k - 1) = False
        End If
        accuracy = K6Function(select_array)
        If accuracy > temp_max Then
            select_index(8) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = False
    select_array(select_index(1)) = True
    select_array(select_index(2)) = True
    select_array(select_index(3)) = True
    select_array(select_index(4)) = True
    select_array(select_index(5)) = True
    select_array(select_index(6)) = True
    select_array(select_index(7)) = True
    select_array(select_index(8)) = True
    If select_index(8) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Attribute : " & select_index(8) & vbTab & "Accuracy : " & temp_max
    End If
line_end:
    List1.AddItem "K=6 , END"
    output = set_output(select_array)
End Sub
Private Sub forward_k5_Click()
'K=5時
    Dim select_array(9) As Boolean '從index(1)開始存八個attribute
    Dim accuracy As Double
    Dim select_index(9) As Integer '紀錄選到哪幾個attribute，從1開始，共8個屬性，故宣告為9
    Dim temp_max As Double
    Dim i As Integer, k As Integer
    Dim output As Variant
    temp_max = 0
    '初始化八個屬性的是否選擇
    For i = 0 To 8
        select_array(i) = False
    Next i
'選1個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        select_array(k) = True
        select_array(k - 1) = False
        accuracy = K5Function(select_array)
        If accuracy > temp_max Then
            select_index(1) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = False '把最後一個屬性初始化
    select_array(select_index(1)) = True '把紀錄到的屬性選擇起來掉 (0 0 1 0 0 0 0 0)
    If select_index(1) = 0 Then '等於0時，代表沒有K值輸入，沒有原MAX他更大的值，所以直接結束選取
        GoTo line_end
    Else
        List1.AddItem "Attribute : " & select_index(1) & vbTab & "Accuracy : " & temp_max
    End If
'選2個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = True
        Else
            select_array(k) = True
            select_array(k - 1) = False
        End If
        accuracy = K5Function(select_array)
        If accuracy > temp_max Then
            select_index(2) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = False
    select_array(select_index(1)) = True
    select_array(select_index(2)) = True
    If select_index(2) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Attribute : " & select_index(2) & vbTab & "Accuracy : " & temp_max
    End If
'選3個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = True
        Else
            select_array(k) = True
            select_array(k - 1) = False
        End If
        accuracy = K5Function(select_array)
        If accuracy > temp_max Then
            select_index(3) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = False
    select_array(select_index(1)) = True
    select_array(select_index(2)) = True
    select_array(select_index(3)) = True
    If select_index(3) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Attribute : " & select_index(3) & vbTab & "Accuracy : " & temp_max
    End If
'選4個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = True
        Else
            select_array(k) = True
            select_array(k - 1) = False
        End If
        accuracy = K5Function(select_array)
        If accuracy > temp_max Then
            select_index(4) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = False
    select_array(select_index(1)) = True
    select_array(select_index(2)) = True
    select_array(select_index(3)) = True
    select_array(select_index(4)) = True
    If select_index(4) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Attribute : " & select_index(4) & vbTab & "Accuracy : " & temp_max
    End If
'選5個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(4) + 1 Then
            select_array(k) = True
        Else
            select_array(k) = True
            select_array(k - 1) = False
        End If
        accuracy = K5Function(select_array)
        If accuracy > temp_max Then
            select_index(5) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = False
    select_array(select_index(1)) = True
    select_array(select_index(2)) = True
    select_array(select_index(3)) = True
    select_array(select_index(4)) = True
    select_array(select_index(5)) = True
    If select_index(5) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Attribute : " & select_index(5) & vbTab & "Accuracy : " & temp_max
    End If
'選6個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(4) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(5) + 1 Then
            select_array(k) = True
        Else
            select_array(k) = True
            select_array(k - 1) = False
        End If
        accuracy = K5Function(select_array)
        If accuracy > temp_max Then
            select_index(6) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = False
    select_array(select_index(1)) = True
    select_array(select_index(2)) = True
    select_array(select_index(3)) = True
    select_array(select_index(4)) = True
    select_array(select_index(5)) = True
    select_array(select_index(6)) = True
    If select_index(6) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Attribute : " & select_index(6) & vbTab & "Accuracy : " & temp_max
    End If
'選7個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(4) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(5) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(6) + 1 Then
            select_array(k) = True
        Else
            select_array(k) = True
            select_array(k - 1) = False
        End If
        accuracy = K5Function(select_array)
        If accuracy > temp_max Then
            select_index(7) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = False
    select_array(select_index(1)) = True
    select_array(select_index(2)) = True
    select_array(select_index(3)) = True
    select_array(select_index(4)) = True
    select_array(select_index(5)) = True
    select_array(select_index(6)) = True
    select_array(select_index(7)) = True
    If select_index(7) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Attribute : " & select_index(7) & vbTab & "Accuracy : " & temp_max
    End If
'選8個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(4) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(5) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(6) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(7) + 1 Then
            select_array(k) = True
        Else
            select_array(k) = True
            select_array(k - 1) = False
        End If
        accuracy = K5Function(select_array)
        If accuracy > temp_max Then
            select_index(8) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = False
    select_array(select_index(1)) = True
    select_array(select_index(2)) = True
    select_array(select_index(3)) = True
    select_array(select_index(4)) = True
    select_array(select_index(5)) = True
    select_array(select_index(6)) = True
    select_array(select_index(7)) = True
    select_array(select_index(8)) = True
    If select_index(8) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Attribute : " & select_index(8) & vbTab & "Accuracy : " & temp_max
    End If
line_end:
    List1.AddItem "K=5 , END"
    output = set_output(select_array)
End Sub

Private Sub forward_k4_Click()
'K=4時
    Dim select_array(9) As Boolean '從index(1)開始存八個attribute
    Dim accuracy As Double
    Dim select_index(9) As Integer '紀錄選到哪幾個attribute，從1開始，共8個屬性，故宣告為9
    Dim temp_max As Double
    Dim i As Integer, k As Integer
    Dim output As Variant
    temp_max = 0
    '初始化八個屬性的是否選擇
    For i = 0 To 8
        select_array(i) = False
    Next i
'選1個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        select_array(k) = True
        select_array(k - 1) = False
        accuracy = K4Function(select_array)
        If accuracy > temp_max Then
            select_index(1) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = False '把最後一個屬性初始化
    select_array(select_index(1)) = True '把紀錄到的屬性選擇起來掉 (0 0 1 0 0 0 0 0)
    If select_index(1) = 0 Then '等於0時，代表沒有K值輸入，沒有原MAX他更大的值，所以直接結束選取
        GoTo line_end
    Else
       List1.AddItem "Attribute : " & select_index(1) & vbTab & "Accuracy : " & temp_max
    End If
'選2個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = True
        Else
            select_array(k) = True
            select_array(k - 1) = False
        End If
        accuracy = K4Function(select_array)
        If accuracy > temp_max Then
            select_index(2) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = False
    select_array(select_index(1)) = True
    select_array(select_index(2)) = True
    If select_index(2) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Attribute : " & select_index(2) & vbTab & "Accuracy : " & temp_max
    End If
'選3個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = True
        Else
            select_array(k) = True
            select_array(k - 1) = False
        End If
        accuracy = K4Function(select_array)
        If accuracy > temp_max Then
            select_index(3) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = False
    select_array(select_index(1)) = True
    select_array(select_index(2)) = True
    select_array(select_index(3)) = True
    If select_index(3) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Attribute : " & select_index(3) & vbTab & "Accuracy : " & temp_max
    End If
'選4個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = True
        Else
            select_array(k) = True
            select_array(k - 1) = False
        End If
        accuracy = K4Function(select_array)
        If accuracy > temp_max Then
            select_index(4) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = False
    select_array(select_index(1)) = True
    select_array(select_index(2)) = True
    select_array(select_index(3)) = True
    select_array(select_index(4)) = True
    If select_index(4) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Attribute : " & select_index(4) & vbTab & "Accuracy : " & temp_max
    End If
'選5個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(4) + 1 Then
            select_array(k) = True
        Else
            select_array(k) = True
            select_array(k - 1) = False
        End If
        accuracy = K4Function(select_array)
        If accuracy > temp_max Then
            select_index(5) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = False
    select_array(select_index(1)) = True
    select_array(select_index(2)) = True
    select_array(select_index(3)) = True
    select_array(select_index(4)) = True
    select_array(select_index(5)) = True
    If select_index(5) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Attribute : " & select_index(5) & vbTab & "Accuracy : " & temp_max
    End If
'選6個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(4) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(5) + 1 Then
            select_array(k) = True
        Else
            select_array(k) = True
            select_array(k - 1) = False
        End If
        accuracy = K4Function(select_array)
        If accuracy > temp_max Then
            select_index(6) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = False
    select_array(select_index(1)) = True
    select_array(select_index(2)) = True
    select_array(select_index(3)) = True
    select_array(select_index(4)) = True
    select_array(select_index(5)) = True
    select_array(select_index(6)) = True
    If select_index(6) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Attribute : " & select_index(6) & vbTab & "Accuracy : " & temp_max
    End If
'選7個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(4) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(5) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(6) + 1 Then
            select_array(k) = True
        Else
            select_array(k) = True
            select_array(k - 1) = False
        End If
        accuracy = K4Function(select_array)
        If accuracy > temp_max Then
            select_index(7) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = False
    select_array(select_index(1)) = True
    select_array(select_index(2)) = True
    select_array(select_index(3)) = True
    select_array(select_index(4)) = True
    select_array(select_index(5)) = True
    select_array(select_index(6)) = True
    select_array(select_index(7)) = True
    If select_index(7) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Attribute : " & select_index(7) & vbTab & "Accuracy : " & temp_max
    End If
'選8個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(4) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(5) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(6) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(7) + 1 Then
            select_array(k) = True
        Else
            select_array(k) = True
            select_array(k - 1) = False
        End If
        accuracy = K4Function(select_array)
        If accuracy > temp_max Then
            select_index(8) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = False
    select_array(select_index(1)) = True
    select_array(select_index(2)) = True
    select_array(select_index(3)) = True
    select_array(select_index(4)) = True
    select_array(select_index(5)) = True
    select_array(select_index(6)) = True
    select_array(select_index(7)) = True
    select_array(select_index(8)) = True
    If select_index(8) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Attribute : " & select_index(8) & vbTab & "Accuracy : " & temp_max
    End If
    
line_end:
    List1.AddItem "K=4 , END"
    output = set_output(select_array)
End Sub
Private Sub forward_Click()
List1.Clear
'K=3時
    Dim select_array(9) As Boolean '從index(1)開始存八個attribute
    Dim accuracy As Double
    Dim select_index(9) As Integer '紀錄選到哪幾個attribute，從1開始，共8個屬性，故宣告為9
    Dim temp_max As Double
    Dim i As Integer, k As Integer
    Dim output As Variant
    temp_max = 0
    '初始化八個屬性的是否選擇
    For i = 0 To 8
        select_array(i) = False
    Next i
'選1個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        select_array(k) = True
        select_array(k - 1) = False
        accuracy = K3Function(select_array)
        If accuracy > temp_max Then
            select_index(1) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = False
    select_array(select_index(1)) = True
    If select_index(1) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Attribute : " & select_index(1) & vbTab & "Accuracy : " & temp_max
    End If
'選2個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = True
        Else
            select_array(k) = True
            select_array(k - 1) = False
        End If
        accuracy = K3Function(select_array)
        If accuracy > temp_max Then
            select_index(2) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = False
    select_array(select_index(1)) = True
    select_array(select_index(2)) = True
    If select_index(2) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Attribute : " & select_index(2) & vbTab & "Accuracy : " & temp_max
    End If
''選3個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = True
        Else
            select_array(k) = True
            select_array(k - 1) = False
        End If
        accuracy = K3Function(select_array)
        If accuracy > temp_max Then
            select_index(3) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = False
    select_array(select_index(1)) = True
    select_array(select_index(2)) = True
    select_array(select_index(3)) = True
    If select_index(3) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Attribute : " & select_index(3) & vbTab & "Accuracy : " & temp_max
    End If
''選4個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = True
        Else
            select_array(k) = True
            select_array(k - 1) = False
        End If
        accuracy = K3Function(select_array)
        If accuracy > temp_max Then
            select_index(4) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = False
    select_array(select_index(1)) = True
    select_array(select_index(2)) = True
    select_array(select_index(3)) = True
    select_array(select_index(4)) = True
    If select_index(4) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Attribute : " & select_index(4) & vbTab & "Accuracy : " & temp_max
    End If
''選5個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(4) + 1 Then
            select_array(k) = True
        Else
            select_array(k) = True
            select_array(k - 1) = False
        End If
        accuracy = K3Function(select_array)
        If accuracy > temp_max Then
            select_index(5) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = False
    select_array(select_index(1)) = True
    select_array(select_index(2)) = True
    select_array(select_index(3)) = True
    select_array(select_index(4)) = True
    select_array(select_index(5)) = True
    If select_index(5) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Attribute : " & select_index(5) & vbTab & "Accuracy : " & temp_max
    End If
'選6個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(4) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(5) + 1 Then
            select_array(k) = True
        Else
            select_array(k) = True
            select_array(k - 1) = False
        End If
        accuracy = K3Function(select_array)
        If accuracy > temp_max Then
            select_index(6) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = False
    select_array(select_index(1)) = True
    select_array(select_index(2)) = True
    select_array(select_index(3)) = True
    select_array(select_index(4)) = True
    select_array(select_index(5)) = True
    select_array(select_index(6)) = True
    If select_index(6) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Attribute : " & select_index(6) & vbTab & "Accuracy : " & temp_max
    End If
'選7個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(4) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(5) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(6) + 1 Then
            select_array(k) = True
        Else
            select_array(k) = True
            select_array(k - 1) = False
        End If
        accuracy = K3Function(select_array)
        If accuracy > temp_max Then
            select_index(7) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = False
    select_array(select_index(1)) = True
    select_array(select_index(2)) = True
    select_array(select_index(3)) = True
    select_array(select_index(4)) = True
    select_array(select_index(5)) = True
    select_array(select_index(6)) = True
    select_array(select_index(7)) = True
    If select_index(7) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Attribute : " & select_index(7) & vbTab & "Accuracy : " & temp_max
    End If
'選8個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(4) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(5) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(6) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(7) + 1 Then
            select_array(k) = True
        Else
            select_array(k) = True
            select_array(k - 1) = False
        End If
        accuracy = K3Function(select_array)
        If accuracy > temp_max Then
            select_index(8) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = False
    select_array(select_index(1)) = True
    select_array(select_index(2)) = True
    select_array(select_index(3)) = True
    select_array(select_index(4)) = True
    select_array(select_index(5)) = True
    select_array(select_index(6)) = True
    select_array(select_index(7)) = True
    select_array(select_index(8)) = True
    If select_index(8) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Attribute : " & select_index(8) & vbTab & "Accuracy : " & temp_max
    End If
line_end:
    List1.AddItem "K=3 , END"
    output = set_output(select_array)
End Sub
Private Sub backward_Click()
List1.Clear
'K=3時
    Dim select_array(9) As Boolean '從index(1)開始存八個attribute
    Dim accuracy As Double
    Dim select_index(9) As Integer '紀錄選到哪幾個attribute，從1開始，共8個屬性，故宣告為9
    Dim temp_max As Double
    Dim i As Integer, k As Integer
    Dim output As Variant
    temp_max = 0
    '初始化八個屬性的是否選擇
    For i = 0 To 8
        select_array(i) = True
    Next i
    accuracy = K3Function(select_array)
    temp_max = accuracy
    List1.AddItem "Remove Attribute : " & select_index(0) & vbTab & "Accuracy : " & temp_max
'移除1個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        select_array(k) = False
        select_array(k - 1) = True
        accuracy = K3Function(select_array)
        If accuracy > temp_max Then
            select_index(1) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = True
    select_array(select_index(1)) = False
    If select_index(1) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Remove Attribute : " & select_index(1) & vbTab & "Accuracy : " & temp_max
    End If
'移除2個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = False
        Else
            select_array(k) = False
            select_array(k - 1) = True
        End If
        accuracy = K3Function(select_array)
        If accuracy > temp_max Then
            select_index(2) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = True
    select_array(select_index(1)) = False
    select_array(select_index(2)) = False
    If select_index(2) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Remove Attribute : " & select_index(2) & vbTab & "Accuracy : " & temp_max
    End If
'移除3個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = False
        Else
            select_array(k) = False
            select_array(k - 1) = True
        End If
        accuracy = K3Function(select_array)
        If accuracy > temp_max Then
            select_index(3) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = True
    select_array(select_index(1)) = False
    select_array(select_index(2)) = False
    select_array(select_index(3)) = False
    If select_index(3) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Remove Attribute : " & select_index(3) & vbTab & "Accuracy : " & temp_max
    End If
'移除4個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = False
        Else
            select_array(k) = False
            select_array(k - 1) = True
        End If
        accuracy = K3Function(select_array)
        If accuracy > temp_max Then
            select_index(4) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = True
    select_array(select_index(1)) = False
    select_array(select_index(2)) = False
    select_array(select_index(3)) = False
    select_array(select_index(4)) = False
    If select_index(4) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Remove Attribute : " & select_index(4) & vbTab & "Accuracy : " & temp_max
    End If
'移除5個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(4) + 1 Then
            select_array(k) = False
        Else
            select_array(k) = False
            select_array(k - 1) = True
        End If
        accuracy = K3Function(select_array)
        If accuracy > temp_max Then
            select_index(5) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = True
    select_array(select_index(1)) = False
    select_array(select_index(2)) = False
    select_array(select_index(3)) = False
    select_array(select_index(4)) = False
    select_array(select_index(5)) = False
    If select_index(5) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Remove Attribute : " & select_index(5) & vbTab & "Accuracy : " & temp_max
    End If
'移除6個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(4) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(5) + 1 Then
            select_array(k) = False
        Else
            select_array(k) = False
            select_array(k - 1) = True
        End If
        accuracy = K3Function(select_array)
        If accuracy > temp_max Then
            select_index(6) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = True
    select_array(select_index(1)) = False
    select_array(select_index(2)) = False
    select_array(select_index(3)) = False
    select_array(select_index(4)) = False
    select_array(select_index(5)) = False
    select_array(select_index(6)) = False
    If select_index(6) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Remove Attribute : " & select_index(6) & vbTab & "Accuracy : " & temp_max
    End If
'移除7個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(4) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(5) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(6) + 1 Then
            select_array(k) = False
        Else
            select_array(k) = False
            select_array(k - 1) = True
        End If
        accuracy = K3Function(select_array)
        If accuracy > temp_max Then
            select_index(7) = k '紀錄選擇的attribute
            temp_max = accuracy
        End If
    Next k
    select_array(8) = True
    select_array(select_index(1)) = False
    select_array(select_index(2)) = False
    select_array(select_index(3)) = False
    select_array(select_index(4)) = False
    select_array(select_index(5)) = False
    select_array(select_index(6)) = False
    select_array(select_index(7)) = False
    If select_index(7) = 0 Then
        GoTo line_end
    Else
        List1.AddItem "Remove Attribute : " & select_index(7) & vbTab & "Accuracy : " & temp_max
    End If
line_end:
    List1.AddItem "K=3 , END"
    output = set_output(select_array)
End Sub
Private Sub random_Click()
    Dim n As Integer, i As Integer, num As Integer, temp As Integer, k As Integer, j As Integer, x As Integer, f As Integer, m As Integer
    Dim random_array(1485) As Integer
    n = 1484
    For i = 1 To n
        random_array(i) = i '存1-1484到index(1)~(1484)
    Next
    For i = n To 1 Step -1
        num = Fix(i * Rnd) + 1
        temp = random_array(i)
        random_array(i) = random_array(num)
        random_array(num) = temp
    Next
    x = 1485
    For j = 0 To 296 '放random(1)-(297)
        x = x - 1
        For k = 1 To 9
            fold3(j, k) = fileArray(random_array(x), k)
        Next k
    Next j
    For j = 0 To 296  '放random(298)-(594)
        x = x - 1
        For k = 1 To 9
            fold2(j, k) = fileArray(random_array(x), k)
        Next k
    Next j
    For j = 0 To 296 '放random(595)-(891)
        x = x - 1
        For k = 1 To 9
            fold1(j, k) = fileArray(random_array(x), k)
        Next k
    Next j
    For j = 0 To 296   '放random(892)-(1188)
        x = x - 1
        For k = 1 To 9
            fold4(j, k) = fileArray(random_array(x), k)
        Next k
    Next j
    For j = 0 To 295 '放random(1189)-(1484)
        x = x - 1
        For k = 1 To 9
            fold5(j, k) = fileArray(random_array(x), k)
        Next k
    Next j
    
    trainingArray() = trainingdata(fold2, fold3, fold4, fold5)
    For n = 0 To 296
        For j = 0 To 1187 '一個n和其他j個data的距離
            For m = 1 To 8
                foldDis(j, n, 1) = foldDis(j, n, 1) + ((Val(fold1(n, m)) - Val(trainingArray(j, m))) ^ 2)
            Next m
            foldDis(j, n, 1) = foldDis(j, n, 1) ^ 0.5
        Next j
    Next n
    trainingArray2() = trainingdata(fold1, fold3, fold4, fold5)
    For n = 0 To 296
        For j = 0 To 1187 '一個n和其他j個data的距離
            For m = 1 To 8
                foldDis(j, n, 2) = foldDis(j, n, 2) + ((Val(fold2(n, m)) - Val(trainingArray2(j, m))) ^ 2)
            Next m
            foldDis(j, n, 2) = foldDis(j, n, 2) ^ 0.5
        Next j
    Next n
    trainingArray3() = trainingdata(fold1, fold2, fold4, fold5)
    For n = 0 To 296
        For j = 0 To 1187 '一個n和其他j個data的距離
            For m = 1 To 8
                foldDis(j, n, 3) = foldDis(j, n, 3) + ((Val(fold3(n, m)) - Val(trainingArray3(j, m))) ^ 2)
            Next m
            foldDis(j, n, 3) = foldDis(j, n, 3) ^ 0.5
        Next j
    Next n
    trainingArray4() = trainingdata(fold1, fold2, fold3, fold5)
    For n = 0 To 296
        For j = 0 To 1187 '一個n和其他j個data的距離
            For m = 1 To 8
                foldDis(j, n, 4) = foldDis(j, n, 4) + ((Val(fold4(n, m)) - Val(trainingArray4(j, m))) ^ 2)
            Next m
            foldDis(j, n, 4) = foldDis(j, n, 4) ^ 0.5
        Next j
    Next n
    trainingArray5() = trainingdata(fold1, fold2, fold3, fold4)
    For n = 0 To 296
        For j = 0 To 1187 '一個n和其他j個data的距離
            For m = 1 To 8
                foldDis(j, n, 5) = foldDis(j, n, 5) + ((Val(fold5(n, m)) - Val(trainingArray5(j, m))) ^ 2)
            Next m
            foldDis(j, n, 5) = foldDis(j, n, 5) ^ 0.5
        Next j
    Next n
    
    MsgBox "OK"
    List1.Clear
End Sub
Public Function trainingdata(foldArr1, foldArr2, foldArr3, foldArr4)
Dim t As Integer, k As Integer, y As Integer
Dim training(1188, 10) As Variant
    y = 0
    For t = 0 To 296  '放fold2(0-296)
         For k = 1 To 9
             training(t, k) = foldArr1(y, k)
         Next k
         y = y + 1
     Next t
     y = 0
     For t = 297 To 593  '放fold3(0-296)
         For k = 1 To 9
             training(t, k) = foldArr2(y, k)
         Next k
         y = y + 1
     Next t
     y = 0
     For t = 594 To 890  '放fold4(0-296)
         For k = 1 To 9
             training(t, k) = foldArr3(y, k)
         Next k
         y = y + 1
     Next t
      y = 0
     For t = 891 To 1187  '放fold5(0-296)
         For k = 1 To 9
             training(t, k) = foldArr4(y, k)
         Next k
         y = y + 1
     Next t
    trainingdata = training
End Function
Private Sub three_nearst_Click()
List1.AddItem "K=3"
    Dim i As Integer, j As Integer, k As Integer, m As Integer, t As Integer, n As Integer, x As Integer, y As Integer
    Dim temp As Double, temp1 As Double, temp2 As Double, temp_index As Double, maxclass As Integer
    For n = 0 To 296
Dim maxDis(3) As Double
Dim maxInd(3) As Double
        For i = 0 To 2
            maxDis(i) = 1
        Next i
        For k = 0 To 1187
            If foldDis(k, n, 1) < maxDis(0) Then
                maxDis(2) = maxDis(1)
                maxDis(1) = maxDis(0)
                maxDis(0) = foldDis(k, n, 1)
                maxInd(2) = maxInd(1)
                maxInd(1) = maxInd(0)
                maxInd(0) = k
            ElseIf foldDis(k, n, 1) >= maxDis(0) And foldDis(k, n, 1) <= maxDis(1) Then
                maxDis(2) = maxDis(1)
                maxDis(1) = foldDis(k, n, 1)
                maxInd(2) = maxInd(1)
                maxInd(1) = k
            ElseIf foldDis(k, n, 1) >= maxDis(1) And foldDis(k, n, 1) <= maxDis(2) Then
                maxDis(2) = foldDis(k, n, 1)
                maxInd(2) = k
            End If
        Next k
        '取前三近的data
Dim class_count(11) As Integer
        For x = 0 To 10
            class_count(x) = 0
        Next x
        For t = 0 To 2
            Select Case trainingArray(maxInd(t), 9)
                Case "CYT"
                    class_count(1) = class_count(1) + 1
                Case "ERL"
                    class_count(2) = class_count(2) + 1
                Case "EXC"
                    class_count(3) = class_count(3) + 1
                Case "ME1"
                    class_count(4) = class_count(4) + 1
                Case "ME2"
                    class_count(5) = class_count(5) + 1
                Case "ME3"
                   class_count(6) = class_count(6) + 1
                Case "MIT"
                   class_count(7) = class_count(7) + 1
                Case "NUC"
                   class_count(8) = class_count(8) + 1
                Case "POX"
                    class_count(9) = class_count(9) + 1
                Case "VAC"
                    class_count(10) = class_count(10) + 1
            End Select
        Next t
Dim maxclasscount As Integer
        maxclasscount = 0
        maxclass = 0 '要預測的class值
        For m = 1 To 10
            If class_count(m) > maxclasscount Then '由小到大排序
                maxclasscount = class_count(m)
                maxclass = m
            End If
        Next m
Dim correct(6) As Integer
Dim accuracy(6) As Single
        If class_name(maxclass) = fold1(n, 9) Then
            correct(1) = correct(1) + 1
        End If
        accuracy(1) = correct(1) / 297
    Next n
    List1.AddItem "1st fold : 297" & " " & "correct : " & correct(1) & " " & "accuracy : " & accuracy(1)
'---------------------------------------------------------------------------------------------------------------
    For n = 0 To 296
    Erase maxDis
        For i = 0 To 2
            maxDis(i) = 1
        Next i
        For k = 0 To 1187
            If foldDis(k, n, 2) < maxDis(0) Then
                maxDis(2) = maxDis(1)
                maxDis(1) = maxDis(0)
                maxDis(0) = foldDis(k, n, 2)
                maxInd(2) = maxInd(1)
                maxInd(1) = maxInd(0)
                maxInd(0) = k
            ElseIf foldDis(k, n, 2) >= maxDis(0) And foldDis(k, n, 2) <= maxDis(1) Then
                maxDis(2) = maxDis(1)
                maxDis(1) = foldDis(k, n, 2)
                maxInd(2) = maxInd(1)
                maxInd(1) = k
            ElseIf foldDis(k, n, 2) >= maxDis(1) And foldDis(k, n, 2) <= maxDis(2) Then
                maxDis(2) = foldDis(k, n, 2)
                maxInd(2) = k
            End If
        Next k
        '取前三近的data
    Erase class_count
        For t = 0 To 2
            Select Case trainingArray2(maxInd(t), 9)
                Case "CYT"
                    class_count(1) = class_count(1) + 1
                Case "ERL"
                    class_count(2) = class_count(2) + 1
                Case "EXC"
                    class_count(3) = class_count(3) + 1
                Case "ME1"
                    class_count(4) = class_count(4) + 1
                Case "ME2"
                    class_count(5) = class_count(5) + 1
                Case "ME3"
                   class_count(6) = class_count(6) + 1
                Case "MIT"
                   class_count(7) = class_count(7) + 1
                Case "NUC"
                   class_count(8) = class_count(8) + 1
                Case "POX"
                    class_count(9) = class_count(9) + 1
                Case "VAC"
                    class_count(10) = class_count(10) + 1
            End Select
        Next t
        maxclasscount = 0
        maxclass = 0 '要預測的class值
        For m = 1 To 10
            If class_count(m) > maxclasscount Then '由小到大排序
                maxclasscount = class_count(m)
                maxclass = m
            End If
        Next m
        If class_name(maxclass) = fold2(n, 9) Then
            correct(2) = correct(2) + 1
        End If
        accuracy(2) = correct(2) / 297
    Next n
    List1.AddItem "2nd fold : 297" & " " & "correct : " & correct(2) & " " & "accuracy : " & accuracy(2)
'---------------------------------------------------------------------------------------------------------------
    For n = 0 To 296
    Erase maxDis
        For i = 0 To 2
            maxDis(i) = 1
        Next i
        For k = 0 To 1187
            If foldDis(k, n, 3) < maxDis(0) Then
                maxDis(2) = maxDis(1)
                maxDis(1) = maxDis(0)
                maxDis(0) = foldDis(k, n, 3)
                maxInd(2) = maxInd(1)
                maxInd(1) = maxInd(0)
                maxInd(0) = k
            ElseIf foldDis(k, n, 3) >= maxDis(0) And foldDis(k, n, 3) <= maxDis(1) Then
                maxDis(2) = maxDis(1)
                maxDis(1) = foldDis(k, n, 3)
                maxInd(2) = maxInd(1)
                maxInd(1) = k
            ElseIf foldDis(k, n, 3) >= maxDis(1) And foldDis(k, n, 3) <= maxDis(2) Then
                maxDis(2) = foldDis(k, n, 3)
                maxInd(2) = k
            End If
        Next k
        '取前三近的data
    Erase class_count
        For t = 0 To 2
            Select Case trainingArray3(maxInd(t), 9)
                Case "CYT"
                    class_count(1) = class_count(1) + 1
                Case "ERL"
                    class_count(2) = class_count(2) + 1
                Case "EXC"
                    class_count(3) = class_count(3) + 1
                Case "ME1"
                    class_count(4) = class_count(4) + 1
                Case "ME2"
                    class_count(5) = class_count(5) + 1
                Case "ME3"
                   class_count(6) = class_count(6) + 1
                Case "MIT"
                   class_count(7) = class_count(7) + 1
                Case "NUC"
                   class_count(8) = class_count(8) + 1
                Case "POX"
                    class_count(9) = class_count(9) + 1
                Case "VAC"
                    class_count(10) = class_count(10) + 1
            End Select
        Next t
        maxclasscount = 0
        maxclass = 0 '要預測的class值
        For m = 1 To 10
            If class_count(m) > maxclasscount Then '由小到大排序
                maxclasscount = class_count(m)
                maxclass = m
            End If
        Next m
        If class_name(maxclass) = fold3(n, 9) Then
            correct(3) = correct(3) + 1
        End If
        accuracy(3) = correct(3) / 297
    Next n
    List1.AddItem "3rd fold : 297" & " " & "correct : " & correct(3) & " " & "accuracy : " & accuracy(3)
'---------------------------------------------------------------------------------------------------------------
    For n = 0 To 296
    Erase maxDis
        For i = 0 To 2
            maxDis(i) = 1
        Next i
        For k = 0 To 1187
            If foldDis(k, n, 4) < maxDis(0) Then
                maxDis(2) = maxDis(1)
                maxDis(1) = maxDis(0)
                maxDis(0) = foldDis(k, n, 4)
                maxInd(2) = maxInd(1)
                maxInd(1) = maxInd(0)
                maxInd(0) = k
            ElseIf foldDis(k, n, 4) >= maxDis(0) And foldDis(k, n, 4) <= maxDis(1) Then
                maxDis(2) = maxDis(1)
                maxDis(1) = foldDis(k, n, 4)
                maxInd(2) = maxInd(1)
                maxInd(1) = k
            ElseIf foldDis(k, n, 4) >= maxDis(1) And foldDis(k, n, 4) <= maxDis(2) Then
                maxDis(2) = foldDis(k, n, 4)
                maxInd(2) = k
            End If
        Next k
        '取前三近的data
    Erase class_count
        For t = 0 To 2
            Select Case trainingArray4(maxInd(t), 9)
                Case "CYT"
                    class_count(1) = class_count(1) + 1
                Case "ERL"
                    class_count(2) = class_count(2) + 1
                Case "EXC"
                    class_count(3) = class_count(3) + 1
                Case "ME1"
                    class_count(4) = class_count(4) + 1
                Case "ME2"
                    class_count(5) = class_count(5) + 1
                Case "ME3"
                   class_count(6) = class_count(6) + 1
                Case "MIT"
                   class_count(7) = class_count(7) + 1
                Case "NUC"
                   class_count(8) = class_count(8) + 1
                Case "POX"
                    class_count(9) = class_count(9) + 1
                Case "VAC"
                    class_count(10) = class_count(10) + 1
            End Select
        Next t
        maxclasscount = 0
        maxclass = 0 '要預測的class值
        For m = 1 To 10
            If class_count(m) > maxclasscount Then '由小到大排序
                maxclasscount = class_count(m)
                maxclass = m
            End If
        Next m
        If class_name(maxclass) = fold4(n, 9) Then
            correct(4) = correct(4) + 1
        End If
        accuracy(4) = correct(4) / 297
    Next n
    List1.AddItem "4th fold : 297" & " " & "correct : " & correct(4) & " " & "accuracy : " & accuracy(4)
'---------------------------------------------------------------------------------------------------------------
    For n = 0 To 296
        Erase maxDis
        For i = 0 To 2
            maxDis(i) = 1
        Next i
        For k = 0 To 1187
            If foldDis(k, n, 5) < maxDis(0) Then
                maxDis(2) = maxDis(1)
                maxDis(1) = maxDis(0)
                maxDis(0) = foldDis(k, n, 5)
                maxInd(2) = maxInd(1)
                maxInd(1) = maxInd(0)
                maxInd(0) = k
            ElseIf foldDis(k, n, 5) >= maxDis(0) And foldDis(k, n, 5) <= maxDis(1) Then
                maxDis(2) = maxDis(1)
                maxDis(1) = foldDis(k, n, 5)
                maxInd(2) = maxInd(1)
                maxInd(1) = k
            ElseIf foldDis(k, n, 5) >= maxDis(1) And foldDis(k, n, 5) <= maxDis(2) Then
                maxDis(2) = foldDis(k, n, 5)
                maxInd(2) = k
            End If
        Next k
        '取前三近的data
    Erase class_count
        For t = 0 To 2
            Select Case trainingArray5(maxInd(t), 9)
                Case "CYT"
                    class_count(1) = class_count(1) + 1
                Case "ERL"
                    class_count(2) = class_count(2) + 1
                Case "EXC"
                    class_count(3) = class_count(3) + 1
                Case "ME1"
                    class_count(4) = class_count(4) + 1
                Case "ME2"
                    class_count(5) = class_count(5) + 1
                Case "ME3"
                   class_count(6) = class_count(6) + 1
                Case "MIT"
                   class_count(7) = class_count(7) + 1
                Case "NUC"
                   class_count(8) = class_count(8) + 1
                Case "POX"
                    class_count(9) = class_count(9) + 1
                Case "VAC"
                    class_count(10) = class_count(10) + 1
            End Select
        Next t
        maxclasscount = 0
        maxclass = 0 '要預測的class值
        For m = 1 To 10
            If class_count(m) > maxclasscount Then '由小到大排序
                maxclasscount = class_count(m)
                maxclass = m
            End If
        Next m
Dim average As Single
        If class_name(maxclass) = fold5(n, 9) Then
            correct(5) = correct(5) + 1
        End If
        accuracy(5) = correct(5) / 296
    Next n
    List1.AddItem "5th fold : 296" & " " & "correct : " & correct(5) & " " & "accuracy : " & accuracy(5)
    average = (accuracy(1) + accuracy(2) + accuracy(3) + accuracy(4) + accuracy(5)) / 5
    List1.AddItem "average accuracy : " & average
End Sub
Private Sub four_nearst_Click()
List1.AddItem "K=4"
    Dim i As Integer, j As Integer, k As Integer, m As Integer, t As Integer, n As Integer, x As Integer, y As Integer
    Dim temp As Double, temp1 As Double, temp2 As Double, temp_index As Double, maxclass As Integer
    For n = 0 To 296
Dim maxDis(4) As Double
Dim maxInd(4) As Double
        For i = 0 To 3
            maxDis(i) = 1
        Next i
        For k = 0 To 1187
            If foldDis(k, n, 1) < maxDis(0) Then
                maxDis(3) = maxDis(2)
                maxDis(2) = maxDis(1)
                maxDis(1) = maxDis(0)
                maxDis(0) = foldDis(k, n, 1)
                maxInd(3) = maxInd(2)
                maxInd(2) = maxInd(1)
                maxInd(1) = maxInd(0)
                maxInd(0) = k
            ElseIf foldDis(k, n, 1) >= maxDis(0) And foldDis(k, n, 1) <= maxDis(1) Then
                maxDis(3) = maxDis(2)
                maxDis(2) = maxDis(1)
                maxDis(1) = foldDis(k, n, 1)
                maxInd(3) = maxInd(2)
                maxInd(2) = maxInd(1)
                maxInd(1) = k
            ElseIf foldDis(k, n, 1) >= maxDis(1) And foldDis(k, n, 1) <= maxDis(2) Then
                maxDis(3) = maxDis(2)
                maxDis(2) = foldDis(k, n, 1)
                maxInd(3) = maxInd(2)
                maxInd(2) = k
            ElseIf foldDis(k, n, 1) >= maxDis(2) And foldDis(k, n, 1) <= maxDis(3) Then
                maxDis(3) = foldDis(k, n, 1)
                maxInd(3) = k
            End If
        Next k
        '取前三近的data
Dim class_count(11) As Integer
        For x = 0 To 10
            class_count(x) = 0
        Next x
        For t = 0 To 3
            Select Case trainingArray(maxInd(t), 9)
                Case "CYT"
                    class_count(1) = class_count(1) + 1
                Case "ERL"
                    class_count(2) = class_count(2) + 1
                Case "EXC"
                    class_count(3) = class_count(3) + 1
                Case "ME1"
                    class_count(4) = class_count(4) + 1
                Case "ME2"
                    class_count(5) = class_count(5) + 1
                Case "ME3"
                   class_count(6) = class_count(6) + 1
                Case "MIT"
                   class_count(7) = class_count(7) + 1
                Case "NUC"
                   class_count(8) = class_count(8) + 1
                Case "POX"
                    class_count(9) = class_count(9) + 1
                Case "VAC"
                    class_count(10) = class_count(10) + 1
            End Select
        Next t
Dim maxclasscount As Integer
        maxclasscount = 0
        maxclass = 0 '要預測的class值
        For m = 1 To 10
            If class_count(m) > maxclasscount Then '由小到大排序
                maxclasscount = class_count(m)
                maxclass = m
            End If
        Next m
Dim correct(6) As Integer
Dim accuracy(6) As Single
        If class_name(maxclass) = fold1(n, 9) Then
            correct(1) = correct(1) + 1
        End If
        accuracy(1) = correct(1) / 297
    Next n
    List1.AddItem "1st fold : 297" & " " & "correct : " & correct(1) & " " & "accuracy : " & accuracy(1)
'---------------------------------------------------------------------------------------------------------------
    For n = 0 To 296
    Erase maxDis
        For i = 0 To 3
            maxDis(i) = 1
        Next i
        For k = 0 To 1187
            If foldDis(k, n, 2) < maxDis(0) Then
                maxDis(3) = maxDis(2)
                maxDis(2) = maxDis(1)
                maxDis(1) = maxDis(0)
                maxDis(0) = foldDis(k, n, 2)
                maxInd(3) = maxInd(2)
                maxInd(2) = maxInd(1)
                maxInd(1) = maxInd(0)
                maxInd(0) = k
            ElseIf foldDis(k, n, 2) >= maxDis(0) And foldDis(k, n, 2) <= maxDis(1) Then
                maxDis(3) = maxDis(2)
                maxDis(2) = maxDis(1)
                maxDis(1) = foldDis(k, n, 2)
                maxInd(3) = maxInd(2)
                maxInd(2) = maxInd(1)
                maxInd(1) = k
            ElseIf foldDis(k, n, 2) >= maxDis(1) And foldDis(k, n, 2) <= maxDis(2) Then
                maxDis(3) = maxDis(2)
                maxDis(2) = foldDis(k, n, 2)
                maxInd(3) = maxInd(2)
                maxInd(2) = k
            ElseIf foldDis(k, n, 2) >= maxDis(2) And foldDis(k, n, 2) <= maxDis(3) Then
                maxDis(3) = foldDis(k, n, 2)
                maxInd(3) = k
            End If
        Next k
        '取前三近的data
    Erase class_count
        For t = 0 To 3
            Select Case trainingArray2(maxInd(t), 9)
                Case "CYT"
                    class_count(1) = class_count(1) + 1
                Case "ERL"
                    class_count(2) = class_count(2) + 1
                Case "EXC"
                    class_count(3) = class_count(3) + 1
                Case "ME1"
                    class_count(4) = class_count(4) + 1
                Case "ME2"
                    class_count(5) = class_count(5) + 1
                Case "ME3"
                   class_count(6) = class_count(6) + 1
                Case "MIT"
                   class_count(7) = class_count(7) + 1
                Case "NUC"
                   class_count(8) = class_count(8) + 1
                Case "POX"
                    class_count(9) = class_count(9) + 1
                Case "VAC"
                    class_count(10) = class_count(10) + 1
            End Select
        Next t
        maxclasscount = 0
        maxclass = 0 '要預測的class值
        For m = 1 To 10
            If class_count(m) > maxclasscount Then '由小到大排序
                maxclasscount = class_count(m)
                maxclass = m
            End If
        Next m
        If class_name(maxclass) = fold2(n, 9) Then
            correct(2) = correct(2) + 1
        End If
        accuracy(2) = correct(2) / 297
    Next n
    List1.AddItem "2nd fold : 297" & " " & "correct : " & correct(2) & " " & "accuracy : " & accuracy(2)
'---------------------------------------------------------------------------------------------------------------
    For n = 0 To 296
    Erase maxDis
        For i = 0 To 3
            maxDis(i) = 1
        Next i
        For k = 0 To 1187
            If foldDis(k, n, 3) < maxDis(0) Then
                maxDis(3) = maxDis(2)
                maxDis(2) = maxDis(1)
                maxDis(1) = maxDis(0)
                maxDis(0) = foldDis(k, n, 3)
                maxInd(3) = maxInd(2)
                maxInd(2) = maxInd(1)
                maxInd(1) = maxInd(0)
                maxInd(0) = k
            ElseIf foldDis(k, n, 3) >= maxDis(0) And foldDis(k, n, 3) <= maxDis(1) Then
                maxDis(3) = maxDis(2)
                maxDis(2) = maxDis(1)
                maxDis(1) = foldDis(k, n, 3)
                maxInd(3) = maxInd(2)
                maxInd(2) = maxInd(1)
                maxInd(1) = k
            ElseIf foldDis(k, n, 3) >= maxDis(1) And foldDis(k, n, 3) <= maxDis(2) Then
                maxDis(3) = maxDis(2)
                maxDis(2) = foldDis(k, n, 3)
                maxInd(3) = maxInd(2)
                maxInd(2) = k
            ElseIf foldDis(k, n, 3) >= maxDis(2) And foldDis(k, n, 3) <= maxDis(3) Then
                maxDis(3) = foldDis(k, n, 3)
                maxInd(3) = k
            End If
        Next k
        '取前三近的data
    Erase class_count
        For t = 0 To 3
            Select Case trainingArray3(maxInd(t), 9)
                Case "CYT"
                    class_count(1) = class_count(1) + 1
                Case "ERL"
                    class_count(2) = class_count(2) + 1
                Case "EXC"
                    class_count(3) = class_count(3) + 1
                Case "ME1"
                    class_count(4) = class_count(4) + 1
                Case "ME2"
                    class_count(5) = class_count(5) + 1
                Case "ME3"
                   class_count(6) = class_count(6) + 1
                Case "MIT"
                   class_count(7) = class_count(7) + 1
                Case "NUC"
                   class_count(8) = class_count(8) + 1
                Case "POX"
                    class_count(9) = class_count(9) + 1
                Case "VAC"
                    class_count(10) = class_count(10) + 1
            End Select
        Next t
        maxclasscount = 0
        maxclass = 0 '要預測的class值
        For m = 1 To 10
            If class_count(m) > maxclasscount Then '由小到大排序
                maxclasscount = class_count(m)
                maxclass = m
            End If
        Next m
        If class_name(maxclass) = fold3(n, 9) Then
            correct(3) = correct(3) + 1
        End If
        accuracy(3) = correct(3) / 297
    Next n
    List1.AddItem "3rd fold : 297" & " " & "correct : " & correct(3) & " " & "accuracy : " & accuracy(3)
'---------------------------------------------------------------------------------------------------------------
    For n = 0 To 296
    Erase maxDis
        For i = 0 To 3
            maxDis(i) = 1
        Next i
        For k = 0 To 1187
            If foldDis(k, n, 4) < maxDis(0) Then
                maxDis(3) = maxDis(2)
                maxDis(2) = maxDis(1)
                maxDis(1) = maxDis(0)
                maxDis(0) = foldDis(k, n, 4)
                maxInd(3) = maxInd(2)
                maxInd(2) = maxInd(1)
                maxInd(1) = maxInd(0)
                maxInd(0) = k
            ElseIf foldDis(k, n, 4) >= maxDis(0) And foldDis(k, n, 4) <= maxDis(1) Then
                maxDis(3) = maxDis(2)
                maxDis(2) = maxDis(1)
                maxDis(1) = foldDis(k, n, 4)
                maxInd(3) = maxInd(2)
                maxInd(2) = maxInd(1)
                maxInd(1) = k
            ElseIf foldDis(k, n, 4) >= maxDis(1) And foldDis(k, n, 4) <= maxDis(2) Then
                maxDis(3) = maxDis(2)
                maxDis(2) = foldDis(k, n, 4)
                maxInd(3) = maxInd(2)
                maxInd(2) = k
            ElseIf foldDis(k, n, 4) >= maxDis(2) And foldDis(k, n, 4) <= maxDis(3) Then
                maxDis(3) = foldDis(k, n, 4)
                maxInd(3) = k
            End If
        Next k
        '取前三近的data
    Erase class_count
        For t = 0 To 3
            Select Case trainingArray4(maxInd(t), 9)
                Case "CYT"
                    class_count(1) = class_count(1) + 1
                Case "ERL"
                    class_count(2) = class_count(2) + 1
                Case "EXC"
                    class_count(3) = class_count(3) + 1
                Case "ME1"
                    class_count(4) = class_count(4) + 1
                Case "ME2"
                    class_count(5) = class_count(5) + 1
                Case "ME3"
                   class_count(6) = class_count(6) + 1
                Case "MIT"
                   class_count(7) = class_count(7) + 1
                Case "NUC"
                   class_count(8) = class_count(8) + 1
                Case "POX"
                    class_count(9) = class_count(9) + 1
                Case "VAC"
                    class_count(10) = class_count(10) + 1
            End Select
        Next t
        maxclasscount = 0
        maxclass = 0 '要預測的class值
        For m = 1 To 10
            If class_count(m) > maxclasscount Then '由小到大排序
                maxclasscount = class_count(m)
                maxclass = m
            End If
        Next m
        If class_name(maxclass) = fold4(n, 9) Then
            correct(4) = correct(4) + 1
        End If
        accuracy(4) = correct(4) / 297
    Next n
    List1.AddItem "4th fold : 297" & " " & "correct : " & correct(4) & " " & "accuracy : " & accuracy(4)
'---------------------------------------------------------------------------------------------------------------
    For n = 0 To 296
        Erase maxDis
        For i = 0 To 3
            maxDis(i) = 1
        Next i
        For k = 0 To 1187
            If foldDis(k, n, 5) < maxDis(0) Then
                maxDis(3) = maxDis(2)
                maxDis(2) = maxDis(1)
                maxDis(1) = maxDis(0)
                maxDis(0) = foldDis(k, n, 5)
                maxInd(3) = maxInd(2)
                maxInd(2) = maxInd(1)
                maxInd(1) = maxInd(0)
                maxInd(0) = k
            ElseIf foldDis(k, n, 5) >= maxDis(0) And foldDis(k, n, 5) <= maxDis(1) Then
                maxDis(3) = maxDis(2)
                maxDis(2) = maxDis(1)
                maxDis(1) = foldDis(k, n, 5)
                maxInd(3) = maxInd(2)
                maxInd(2) = maxInd(1)
                maxInd(1) = k
            ElseIf foldDis(k, n, 5) >= maxDis(1) And foldDis(k, n, 5) <= maxDis(2) Then
                maxDis(3) = maxDis(2)
                maxDis(2) = foldDis(k, n, 5)
                maxInd(3) = maxInd(2)
                maxInd(2) = k
            ElseIf foldDis(k, n, 5) >= maxDis(2) And foldDis(k, n, 5) <= maxDis(3) Then
                maxDis(3) = foldDis(k, n, 5)
                maxInd(3) = k
            End If
        Next k
        '取前三近的data
    Erase class_count
        For t = 0 To 3
            Select Case trainingArray5(maxInd(t), 9)
                Case "CYT"
                    class_count(1) = class_count(1) + 1
                Case "ERL"
                    class_count(2) = class_count(2) + 1
                Case "EXC"
                    class_count(3) = class_count(3) + 1
                Case "ME1"
                    class_count(4) = class_count(4) + 1
                Case "ME2"
                    class_count(5) = class_count(5) + 1
                Case "ME3"
                   class_count(6) = class_count(6) + 1
                Case "MIT"
                   class_count(7) = class_count(7) + 1
                Case "NUC"
                   class_count(8) = class_count(8) + 1
                Case "POX"
                    class_count(9) = class_count(9) + 1
                Case "VAC"
                    class_count(10) = class_count(10) + 1
            End Select
        Next t
        maxclasscount = 0
        maxclass = 0 '要預測的class值
        For m = 1 To 10
            If class_count(m) > maxclasscount Then '由小到大排序
                maxclasscount = class_count(m)
                maxclass = m
            End If
        Next m
Dim average As Single
        If class_name(maxclass) = fold5(n, 9) Then
            correct(5) = correct(5) + 1
        End If
        accuracy(5) = correct(5) / 296
    Next n
    List1.AddItem "5th fold : 296" & " " & "correct : " & correct(5) & " " & "accuracy : " & accuracy(5)
    average = (accuracy(1) + accuracy(2) + accuracy(3) + accuracy(4) + accuracy(5)) / 5
    List1.AddItem "average accuracy : " & average
End Sub
Private Sub five_nearst_Click()
List1.AddItem "K=5"
    Dim i As Integer, j As Integer, k As Integer, m As Integer, t As Integer, n As Integer, x As Integer, y As Integer
    Dim temp As Double, temp1 As Double, temp2 As Double, temp_index As Double, maxclass As Integer
    For n = 0 To 296
Dim maxDis(5) As Double
Dim maxInd(5) As Double
        For i = 0 To 4
            maxDis(i) = 1
        Next i
        For k = 0 To 1187
            If foldDis(k, n, 1) < maxDis(0) Then
                maxDis(4) = maxDis(3)
                maxDis(3) = maxDis(2)
                maxDis(2) = maxDis(1)
                maxDis(1) = maxDis(0)
                maxDis(0) = foldDis(k, n, 1)
                maxInd(4) = maxInd(3)
                maxInd(3) = maxInd(2)
                maxInd(2) = maxInd(1)
                maxInd(1) = maxInd(0)
                maxInd(0) = k
            ElseIf foldDis(k, n, 1) >= maxDis(0) And foldDis(k, n, 1) <= maxDis(1) Then
                maxDis(4) = maxDis(3)
                maxDis(3) = maxDis(2)
                maxDis(2) = maxDis(1)
                maxDis(1) = foldDis(k, n, 1)
                maxInd(4) = maxInd(3)
                maxInd(3) = maxInd(2)
                maxInd(2) = maxInd(1)
                maxInd(1) = k
            ElseIf foldDis(k, n, 1) >= maxDis(1) And foldDis(k, n, 1) <= maxDis(2) Then
                maxDis(4) = maxDis(3)
                maxDis(3) = maxDis(2)
                maxDis(2) = foldDis(k, n, 1)
                maxInd(4) = maxInd(3)
                maxInd(3) = maxInd(2)
                maxInd(2) = k
            ElseIf foldDis(k, n, 1) >= maxDis(2) And foldDis(k, n, 1) <= maxDis(3) Then
                maxDis(4) = maxDis(3)
                maxDis(3) = foldDis(k, n, 1)
                maxInd(4) = maxInd(3)
                maxInd(3) = k
            ElseIf foldDis(k, n, 1) >= maxDis(3) And foldDis(k, n, 1) <= maxDis(4) Then
                maxDis(4) = foldDis(k, n, 1)
                maxInd(4) = k
            End If
        Next k
        '取前三近的data
Dim class_count(11) As Integer
        For x = 0 To 10
            class_count(x) = 0
        Next x
        For t = 0 To 4
            Select Case trainingArray(maxInd(t), 9)
                Case "CYT"
                    class_count(1) = class_count(1) + 1
                Case "ERL"
                    class_count(2) = class_count(2) + 1
                Case "EXC"
                    class_count(3) = class_count(3) + 1
                Case "ME1"
                    class_count(4) = class_count(4) + 1
                Case "ME2"
                    class_count(5) = class_count(5) + 1
                Case "ME3"
                   class_count(6) = class_count(6) + 1
                Case "MIT"
                   class_count(7) = class_count(7) + 1
                Case "NUC"
                   class_count(8) = class_count(8) + 1
                Case "POX"
                    class_count(9) = class_count(9) + 1
                Case "VAC"
                    class_count(10) = class_count(10) + 1
            End Select
        Next t
Dim maxclasscount As Integer
        maxclasscount = 0
        maxclass = 0 '要預測的class值
        For m = 1 To 10
            If class_count(m) > maxclasscount Then '由小到大排序
                maxclasscount = class_count(m)
                maxclass = m
            End If
        Next m
Dim correct(6) As Integer
Dim accuracy(6) As Single
        If class_name(maxclass) = fold1(n, 9) Then
            correct(1) = correct(1) + 1
        End If
        accuracy(1) = correct(1) / 297
    Next n
    List1.AddItem "1st fold : 297" & " " & "correct : " & correct(1) & " " & "accuracy : " & accuracy(1)
'---------------------------------------------------------------------------------------------------------------
    For n = 0 To 296
    Erase maxDis
        For i = 0 To 4
            maxDis(i) = 1
        Next i
        For k = 0 To 1187
            If foldDis(k, n, 2) < maxDis(0) Then
                maxDis(4) = maxDis(3)
                maxDis(3) = maxDis(2)
                maxDis(2) = maxDis(1)
                maxDis(1) = maxDis(0)
                maxDis(0) = foldDis(k, n, 2)
                maxInd(4) = maxInd(3)
                maxInd(3) = maxInd(2)
                maxInd(2) = maxInd(1)
                maxInd(1) = maxInd(0)
                maxInd(0) = k
            ElseIf foldDis(k, n, 2) >= maxDis(0) And foldDis(k, n, 2) <= maxDis(1) Then
                maxDis(4) = maxDis(3)
                maxDis(3) = maxDis(2)
                maxDis(2) = maxDis(1)
                maxDis(1) = foldDis(k, n, 2)
                maxInd(4) = maxInd(3)
                maxInd(3) = maxInd(2)
                maxInd(2) = maxInd(1)
                maxInd(1) = k
            ElseIf foldDis(k, n, 2) >= maxDis(1) And foldDis(k, n, 2) <= maxDis(2) Then
                maxDis(4) = maxDis(3)
                maxDis(3) = maxDis(2)
                maxDis(2) = foldDis(k, n, 2)
                maxInd(4) = maxInd(3)
                maxInd(3) = maxInd(2)
                maxInd(2) = k
            ElseIf foldDis(k, n, 2) >= maxDis(2) And foldDis(k, n, 2) <= maxDis(3) Then
                maxDis(4) = maxDis(3)
                maxDis(3) = foldDis(k, n, 2)
                maxInd(4) = maxInd(3)
                maxInd(3) = k
            ElseIf foldDis(k, n, 2) >= maxDis(3) And foldDis(k, n, 2) <= maxDis(4) Then
                maxDis(4) = foldDis(k, n, 2)
                maxInd(4) = k
            End If
        Next k
        '取前三近的data
    Erase class_count
        For t = 0 To 4
            Select Case trainingArray2(maxInd(t), 9)
                Case "CYT"
                    class_count(1) = class_count(1) + 1
                Case "ERL"
                    class_count(2) = class_count(2) + 1
                Case "EXC"
                    class_count(3) = class_count(3) + 1
                Case "ME1"
                    class_count(4) = class_count(4) + 1
                Case "ME2"
                    class_count(5) = class_count(5) + 1
                Case "ME3"
                   class_count(6) = class_count(6) + 1
                Case "MIT"
                   class_count(7) = class_count(7) + 1
                Case "NUC"
                   class_count(8) = class_count(8) + 1
                Case "POX"
                    class_count(9) = class_count(9) + 1
                Case "VAC"
                    class_count(10) = class_count(10) + 1
            End Select
        Next t
        maxclasscount = 0
        maxclass = 0 '要預測的class值
        For m = 1 To 10
            If class_count(m) > maxclasscount Then '由小到大排序
                maxclasscount = class_count(m)
                maxclass = m
            End If
        Next m
        If class_name(maxclass) = fold2(n, 9) Then
            correct(2) = correct(2) + 1
        End If
        accuracy(2) = correct(2) / 297
    Next n
    List1.AddItem "2nd fold : 297" & " " & "correct : " & correct(2) & " " & "accuracy : " & accuracy(2)
'---------------------------------------------------------------------------------------------------------------
    For n = 0 To 296
    Erase maxDis
        For i = 0 To 4
            maxDis(i) = 1
        Next i
        For k = 0 To 1187
            If foldDis(k, n, 3) < maxDis(0) Then
                maxDis(4) = maxDis(3)
                maxDis(3) = maxDis(2)
                maxDis(2) = maxDis(1)
                maxDis(1) = maxDis(0)
                maxDis(0) = foldDis(k, n, 3)
                maxInd(4) = maxInd(3)
                maxInd(3) = maxInd(2)
                maxInd(2) = maxInd(1)
                maxInd(1) = maxInd(0)
                maxInd(0) = k
            ElseIf foldDis(k, n, 3) >= maxDis(0) And foldDis(k, n, 3) <= maxDis(1) Then
                maxDis(4) = maxDis(3)
                maxDis(3) = maxDis(2)
                maxDis(2) = maxDis(1)
                maxDis(1) = foldDis(k, n, 3)
                maxInd(4) = maxInd(3)
                maxInd(3) = maxInd(2)
                maxInd(2) = maxInd(1)
                maxInd(1) = k
            ElseIf foldDis(k, n, 3) >= maxDis(1) And foldDis(k, n, 3) <= maxDis(2) Then
                maxDis(4) = maxDis(3)
                maxDis(3) = maxDis(2)
                maxDis(2) = foldDis(k, n, 3)
                maxInd(4) = maxInd(3)
                maxInd(3) = maxInd(2)
                maxInd(2) = k
            ElseIf foldDis(k, n, 3) >= maxDis(2) And foldDis(k, n, 3) <= maxDis(3) Then
                maxDis(4) = maxDis(3)
                maxDis(3) = foldDis(k, n, 3)
                maxInd(4) = maxInd(3)
                maxInd(3) = k
            ElseIf foldDis(k, n, 3) >= maxDis(3) And foldDis(k, n, 3) <= maxDis(4) Then
                maxDis(4) = foldDis(k, n, 3)
                maxInd(4) = k
            End If
        Next k
        '取前三近的data
    Erase class_count
        For t = 0 To 4
            Select Case trainingArray3(maxInd(t), 9)
                Case "CYT"
                    class_count(1) = class_count(1) + 1
                Case "ERL"
                    class_count(2) = class_count(2) + 1
                Case "EXC"
                    class_count(3) = class_count(3) + 1
                Case "ME1"
                    class_count(4) = class_count(4) + 1
                Case "ME2"
                    class_count(5) = class_count(5) + 1
                Case "ME3"
                   class_count(6) = class_count(6) + 1
                Case "MIT"
                   class_count(7) = class_count(7) + 1
                Case "NUC"
                   class_count(8) = class_count(8) + 1
                Case "POX"
                    class_count(9) = class_count(9) + 1
                Case "VAC"
                    class_count(10) = class_count(10) + 1
            End Select
        Next t
        maxclasscount = 0
        maxclass = 0 '要預測的class值
        For m = 1 To 10
            If class_count(m) > maxclasscount Then '由小到大排序
                maxclasscount = class_count(m)
                maxclass = m
            End If
        Next m
        If class_name(maxclass) = fold3(n, 9) Then
            correct(3) = correct(3) + 1
        End If
        accuracy(3) = correct(3) / 297
    Next n
    List1.AddItem "3rd fold : 297" & " " & "correct : " & correct(3) & " " & "accuracy : " & accuracy(3)
'---------------------------------------------------------------------------------------------------------------
    For n = 0 To 296
    Erase maxDis
        For i = 0 To 4
            maxDis(i) = 1
        Next i
        For k = 0 To 1187
            If foldDis(k, n, 4) < maxDis(0) Then
                maxDis(4) = maxDis(3)
                maxDis(3) = maxDis(2)
                maxDis(2) = maxDis(1)
                maxDis(1) = maxDis(0)
                maxDis(0) = foldDis(k, n, 4)
                maxInd(4) = maxInd(3)
                maxInd(3) = maxInd(2)
                maxInd(2) = maxInd(1)
                maxInd(1) = maxInd(0)
                maxInd(0) = k
            ElseIf foldDis(k, n, 4) >= maxDis(0) And foldDis(k, n, 4) <= maxDis(1) Then
                maxDis(4) = maxDis(3)
                maxDis(3) = maxDis(2)
                maxDis(2) = maxDis(1)
                maxDis(1) = foldDis(k, n, 4)
                maxInd(4) = maxInd(3)
                maxInd(3) = maxInd(2)
                maxInd(2) = maxInd(1)
                maxInd(1) = k
            ElseIf foldDis(k, n, 4) >= maxDis(1) And foldDis(k, n, 4) <= maxDis(2) Then
                maxDis(4) = maxDis(3)
                maxDis(3) = maxDis(2)
                maxDis(2) = foldDis(k, n, 4)
                maxInd(4) = maxInd(3)
                maxInd(3) = maxInd(2)
                maxInd(2) = k
            ElseIf foldDis(k, n, 4) >= maxDis(2) And foldDis(k, n, 4) <= maxDis(3) Then
                maxDis(4) = maxDis(3)
                maxDis(3) = foldDis(k, n, 4)
                maxInd(4) = maxInd(3)
                maxInd(3) = k
            ElseIf foldDis(k, n, 4) >= maxDis(3) And foldDis(k, n, 4) <= maxDis(4) Then
                maxDis(4) = foldDis(k, n, 4)
                maxInd(4) = k
            End If
        Next k
        '取前三近的data
    Erase class_count
        For t = 0 To 4
            Select Case trainingArray4(maxInd(t), 9)
                Case "CYT"
                    class_count(1) = class_count(1) + 1
                Case "ERL"
                    class_count(2) = class_count(2) + 1
                Case "EXC"
                    class_count(3) = class_count(3) + 1
                Case "ME1"
                    class_count(4) = class_count(4) + 1
                Case "ME2"
                    class_count(5) = class_count(5) + 1
                Case "ME3"
                   class_count(6) = class_count(6) + 1
                Case "MIT"
                   class_count(7) = class_count(7) + 1
                Case "NUC"
                   class_count(8) = class_count(8) + 1
                Case "POX"
                    class_count(9) = class_count(9) + 1
                Case "VAC"
                    class_count(10) = class_count(10) + 1
            End Select
        Next t
        maxclasscount = 0
        maxclass = 0 '要預測的class值
        For m = 1 To 10
            If class_count(m) > maxclasscount Then '由小到大排序
                maxclasscount = class_count(m)
                maxclass = m
            End If
        Next m
        If class_name(maxclass) = fold4(n, 9) Then
            correct(4) = correct(4) + 1
        End If
        accuracy(4) = correct(4) / 297
    Next n
    List1.AddItem "4th fold : 297" & " " & "correct : " & correct(4) & " " & "accuracy : " & accuracy(4)
'---------------------------------------------------------------------------------------------------------------
    For n = 0 To 296
        Erase maxDis
        For i = 0 To 4
            maxDis(i) = 1
        Next i
        For k = 0 To 1187
            If foldDis(k, n, 5) < maxDis(0) Then
                maxDis(4) = maxDis(3)
                maxDis(3) = maxDis(2)
                maxDis(2) = maxDis(1)
                maxDis(1) = maxDis(0)
                maxDis(0) = foldDis(k, n, 5)
                maxInd(4) = maxInd(3)
                maxInd(3) = maxInd(2)
                maxInd(2) = maxInd(1)
                maxInd(1) = maxInd(0)
                maxInd(0) = k
            ElseIf foldDis(k, n, 5) >= maxDis(0) And foldDis(k, n, 5) <= maxDis(1) Then
                maxDis(4) = maxDis(3)
                maxDis(3) = maxDis(2)
                maxDis(2) = maxDis(1)
                maxDis(1) = foldDis(k, n, 5)
                maxInd(4) = maxInd(3)
                maxInd(3) = maxInd(2)
                maxInd(2) = maxInd(1)
                maxInd(1) = k
            ElseIf foldDis(k, n, 5) >= maxDis(1) And foldDis(k, n, 5) <= maxDis(2) Then
                maxDis(4) = maxDis(3)
                maxDis(3) = maxDis(2)
                maxDis(2) = foldDis(k, n, 5)
                maxInd(4) = maxInd(3)
                maxInd(3) = maxInd(2)
                maxInd(2) = k
            ElseIf foldDis(k, n, 5) >= maxDis(2) And foldDis(k, n, 5) <= maxDis(3) Then
                maxDis(4) = maxDis(3)
                maxDis(3) = foldDis(k, n, 5)
                maxInd(4) = maxInd(3)
                maxInd(3) = k
            ElseIf foldDis(k, n, 5) >= maxDis(3) And foldDis(k, n, 5) <= maxDis(4) Then
                maxDis(4) = foldDis(k, n, 5)
                maxInd(4) = k
            End If
        Next k
        '取前三近的data
    Erase class_count
        For t = 0 To 4
            Select Case trainingArray5(maxInd(t), 9)
                Case "CYT"
                    class_count(1) = class_count(1) + 1
                Case "ERL"
                    class_count(2) = class_count(2) + 1
                Case "EXC"
                    class_count(3) = class_count(3) + 1
                Case "ME1"
                    class_count(4) = class_count(4) + 1
                Case "ME2"
                    class_count(5) = class_count(5) + 1
                Case "ME3"
                   class_count(6) = class_count(6) + 1
                Case "MIT"
                   class_count(7) = class_count(7) + 1
                Case "NUC"
                   class_count(8) = class_count(8) + 1
                Case "POX"
                    class_count(9) = class_count(9) + 1
                Case "VAC"
                    class_count(10) = class_count(10) + 1
            End Select
        Next t
        maxclasscount = 0
        maxclass = 0 '要預測的class值
        For m = 1 To 10
            If class_count(m) > maxclasscount Then '由小到大排序
                maxclasscount = class_count(m)
                maxclass = m
            End If
        Next m
Dim average As Single
        If class_name(maxclass) = fold5(n, 9) Then
            correct(5) = correct(5) + 1
        End If
        accuracy(5) = correct(5) / 296
    Next n
    List1.AddItem "5th fold : 296" & " " & "correct : " & correct(5) & " " & "accuracy : " & accuracy(5)
    average = (accuracy(1) + accuracy(2) + accuracy(3) + accuracy(4) + accuracy(5)) / 5
    List1.AddItem "average accuracy : " & average
End Sub
Private Sub six_nearst_Click()
List1.AddItem "K=6"
    Dim i As Integer, j As Integer, k As Integer, m As Integer, t As Integer, n As Integer, x As Integer, y As Integer
    Dim temp As Double, temp1 As Double, temp2 As Double, temp_index As Double, maxclass As Integer
    For n = 0 To 296
Dim maxDis(6) As Double
Dim maxInd(6) As Double
        For i = 0 To 5
            maxDis(i) = 1
        Next i
        For k = 0 To 1187
            If foldDis(k, n, 1) < maxDis(0) Then
                maxDis(5) = maxDis(4)
                maxDis(4) = maxDis(3)
                maxDis(3) = maxDis(2)
                maxDis(2) = maxDis(1)
                maxDis(1) = maxDis(0)
                maxDis(0) = foldDis(k, n, 1)
                maxInd(5) = maxInd(4)
                maxInd(4) = maxInd(3)
                maxInd(3) = maxInd(2)
                maxInd(2) = maxInd(1)
                maxInd(1) = maxInd(0)
                maxInd(0) = k
            ElseIf foldDis(k, n, 1) >= maxDis(0) And foldDis(k, n, 1) <= maxDis(1) Then
                maxDis(5) = maxDis(4)
                maxDis(4) = maxDis(3)
                maxDis(3) = maxDis(2)
                maxDis(2) = maxDis(1)
                maxDis(1) = foldDis(k, n, 1)
                maxInd(5) = maxInd(4)
                maxInd(4) = maxInd(3)
                maxInd(3) = maxInd(2)
                maxInd(2) = maxInd(1)
                maxInd(1) = k
            ElseIf foldDis(k, n, 1) >= maxDis(1) And foldDis(k, n, 1) <= maxDis(2) Then
                maxDis(5) = maxDis(4)
                maxDis(4) = maxDis(3)
                maxDis(3) = maxDis(2)
                maxDis(2) = foldDis(k, n, 1)
                maxInd(5) = maxInd(4)
                maxInd(4) = maxInd(3)
                maxInd(3) = maxInd(2)
                maxInd(2) = k
            ElseIf foldDis(k, n, 1) >= maxDis(2) And foldDis(k, n, 1) <= maxDis(3) Then
                maxDis(5) = maxDis(4)
                maxDis(4) = maxDis(3)
                maxDis(3) = foldDis(k, n, 1)
                maxInd(5) = maxInd(4)
                maxInd(4) = maxInd(3)
                maxInd(3) = k
            ElseIf foldDis(k, n, 1) >= maxDis(3) And foldDis(k, n, 1) <= maxDis(4) Then
                maxDis(5) = maxDis(4)
                maxDis(4) = foldDis(k, n, 1)
                maxInd(5) = maxInd(4)
                maxInd(4) = k
            ElseIf foldDis(k, n, 1) >= maxDis(4) And foldDis(k, n, 1) <= maxDis(5) Then
                maxDis(5) = foldDis(k, n, 1)
                maxInd(5) = k
            End If
        Next k
        '取前三近的data
Dim class_count(11) As Integer
        For x = 0 To 10
            class_count(x) = 0
        Next x
        For t = 0 To 5
            Select Case trainingArray(maxInd(t), 9)
                Case "CYT"
                    class_count(1) = class_count(1) + 1
                Case "ERL"
                    class_count(2) = class_count(2) + 1
                Case "EXC"
                    class_count(3) = class_count(3) + 1
                Case "ME1"
                    class_count(4) = class_count(4) + 1
                Case "ME2"
                    class_count(5) = class_count(5) + 1
                Case "ME3"
                   class_count(6) = class_count(6) + 1
                Case "MIT"
                   class_count(7) = class_count(7) + 1
                Case "NUC"
                   class_count(8) = class_count(8) + 1
                Case "POX"
                    class_count(9) = class_count(9) + 1
                Case "VAC"
                    class_count(10) = class_count(10) + 1
            End Select
        Next t
Dim maxclasscount As Integer
        maxclasscount = 0
        maxclass = 0 '要預測的class值
        For m = 1 To 10
            If class_count(m) > maxclasscount Then '由小到大排序
                maxclasscount = class_count(m)
                maxclass = m
            End If
        Next m
Dim correct(6) As Integer
Dim accuracy(6) As Single
        If class_name(maxclass) = fold1(n, 9) Then
            correct(1) = correct(1) + 1
        End If
        accuracy(1) = correct(1) / 297
    Next n
    List1.AddItem "1st fold : 297" & " " & "correct : " & correct(1) & " " & "accuracy : " & accuracy(1)
'---------------------------------------------------------------------------------------------------------------
    For n = 0 To 296
    Erase maxDis
        For i = 0 To 5
            maxDis(i) = 1
        Next i
        For k = 0 To 1187
            If foldDis(k, n, 2) < maxDis(0) Then
                maxDis(5) = maxDis(4)
                maxDis(4) = maxDis(3)
                maxDis(3) = maxDis(2)
                maxDis(2) = maxDis(1)
                maxDis(1) = maxDis(0)
                maxDis(0) = foldDis(k, n, 2)
                maxInd(5) = maxInd(4)
                maxInd(4) = maxInd(3)
                maxInd(3) = maxInd(2)
                maxInd(2) = maxInd(1)
                maxInd(1) = maxInd(0)
                maxInd(0) = k
            ElseIf foldDis(k, n, 2) >= maxDis(0) And foldDis(k, n, 2) <= maxDis(1) Then
                maxDis(5) = maxDis(4)
                maxDis(4) = maxDis(3)
                maxDis(3) = maxDis(2)
                maxDis(2) = maxDis(1)
                maxDis(1) = foldDis(k, n, 2)
                maxInd(5) = maxInd(4)
                maxInd(4) = maxInd(3)
                maxInd(3) = maxInd(2)
                maxInd(2) = maxInd(1)
                maxInd(1) = k
            ElseIf foldDis(k, n, 2) >= maxDis(1) And foldDis(k, n, 2) <= maxDis(2) Then
                maxDis(5) = maxDis(4)
                maxDis(4) = maxDis(3)
                maxDis(3) = maxDis(2)
                maxDis(2) = foldDis(k, n, 2)
                maxInd(5) = maxInd(4)
                maxInd(4) = maxInd(3)
                maxInd(3) = maxInd(2)
                maxInd(2) = k
            ElseIf foldDis(k, n, 2) >= maxDis(2) And foldDis(k, n, 2) <= maxDis(3) Then
                maxDis(5) = maxDis(4)
                maxDis(4) = maxDis(3)
                maxDis(3) = foldDis(k, n, 2)
                maxInd(5) = maxInd(4)
                maxInd(4) = maxInd(3)
                maxInd(3) = k
            ElseIf foldDis(k, n, 2) >= maxDis(3) And foldDis(k, n, 2) <= maxDis(4) Then
                maxDis(5) = maxDis(4)
                maxDis(4) = foldDis(k, n, 2)
                maxInd(5) = maxInd(4)
                maxInd(4) = k
            ElseIf foldDis(k, n, 2) >= maxDis(4) And foldDis(k, n, 2) <= maxDis(5) Then
                maxDis(5) = foldDis(k, n, 2)
                maxInd(5) = k
            End If
        Next k
        '取前三近的data
    Erase class_count
        For t = 0 To 5
            Select Case trainingArray2(maxInd(t), 9)
                Case "CYT"
                    class_count(1) = class_count(1) + 1
                Case "ERL"
                    class_count(2) = class_count(2) + 1
                Case "EXC"
                    class_count(3) = class_count(3) + 1
                Case "ME1"
                    class_count(4) = class_count(4) + 1
                Case "ME2"
                    class_count(5) = class_count(5) + 1
                Case "ME3"
                   class_count(6) = class_count(6) + 1
                Case "MIT"
                   class_count(7) = class_count(7) + 1
                Case "NUC"
                   class_count(8) = class_count(8) + 1
                Case "POX"
                    class_count(9) = class_count(9) + 1
                Case "VAC"
                    class_count(10) = class_count(10) + 1
            End Select
        Next t
        maxclasscount = 0
        maxclass = 0 '要預測的class值
        For m = 1 To 10
            If class_count(m) > maxclasscount Then '由小到大排序
                maxclasscount = class_count(m)
                maxclass = m
            End If
        Next m
        If class_name(maxclass) = fold2(n, 9) Then
            correct(2) = correct(2) + 1
        End If
        accuracy(2) = correct(2) / 297
    Next n
    List1.AddItem "2nd fold : 297" & " " & "correct : " & correct(2) & " " & "accuracy : " & accuracy(2)
'---------------------------------------------------------------------------------------------------------------
    For n = 0 To 296
    Erase maxDis
        For i = 0 To 5
            maxDis(i) = 1
        Next i
        For k = 0 To 1187
            If foldDis(k, n, 3) < maxDis(0) Then
                maxDis(5) = maxDis(4)
                maxDis(4) = maxDis(3)
                maxDis(3) = maxDis(2)
                maxDis(2) = maxDis(1)
                maxDis(1) = maxDis(0)
                maxDis(0) = foldDis(k, n, 3)
                maxInd(5) = maxInd(4)
                maxInd(4) = maxInd(3)
                maxInd(3) = maxInd(2)
                maxInd(2) = maxInd(1)
                maxInd(1) = maxInd(0)
                maxInd(0) = k
            ElseIf foldDis(k, n, 3) >= maxDis(0) And foldDis(k, n, 3) <= maxDis(1) Then
                maxDis(5) = maxDis(4)
                maxDis(4) = maxDis(3)
                maxDis(3) = maxDis(2)
                maxDis(2) = maxDis(1)
                maxDis(1) = foldDis(k, n, 3)
                maxInd(5) = maxInd(4)
                maxInd(4) = maxInd(3)
                maxInd(3) = maxInd(2)
                maxInd(2) = maxInd(1)
                maxInd(1) = k
            ElseIf foldDis(k, n, 3) >= maxDis(1) And foldDis(k, n, 3) <= maxDis(2) Then
                maxDis(5) = maxDis(4)
                maxDis(4) = maxDis(3)
                maxDis(3) = maxDis(2)
                maxDis(2) = foldDis(k, n, 3)
                maxInd(5) = maxInd(4)
                maxInd(4) = maxInd(3)
                maxInd(3) = maxInd(2)
                maxInd(2) = k
            ElseIf foldDis(k, n, 3) >= maxDis(2) And foldDis(k, n, 3) <= maxDis(3) Then
                maxDis(5) = maxDis(4)
                maxDis(4) = maxDis(3)
                maxDis(3) = foldDis(k, n, 3)
                maxInd(5) = maxInd(4)
                maxInd(4) = maxInd(3)
                maxInd(3) = k
            ElseIf foldDis(k, n, 3) >= maxDis(3) And foldDis(k, n, 3) <= maxDis(4) Then
                maxDis(5) = maxDis(4)
                maxDis(4) = foldDis(k, n, 3)
                maxInd(5) = maxInd(4)
                maxInd(4) = k
            ElseIf foldDis(k, n, 3) >= maxDis(4) And foldDis(k, n, 3) <= maxDis(5) Then
                maxDis(5) = foldDis(k, n, 3)
                maxInd(5) = k
            End If
        Next k
        '取前三近的data
    Erase class_count
        For t = 0 To 5
            Select Case trainingArray3(maxInd(t), 9)
                Case "CYT"
                    class_count(1) = class_count(1) + 1
                Case "ERL"
                    class_count(2) = class_count(2) + 1
                Case "EXC"
                    class_count(3) = class_count(3) + 1
                Case "ME1"
                    class_count(4) = class_count(4) + 1
                Case "ME2"
                    class_count(5) = class_count(5) + 1
                Case "ME3"
                   class_count(6) = class_count(6) + 1
                Case "MIT"
                   class_count(7) = class_count(7) + 1
                Case "NUC"
                   class_count(8) = class_count(8) + 1
                Case "POX"
                    class_count(9) = class_count(9) + 1
                Case "VAC"
                    class_count(10) = class_count(10) + 1
            End Select
        Next t
        maxclasscount = 0
        maxclass = 0 '要預測的class值
        For m = 1 To 10
            If class_count(m) > maxclasscount Then '由小到大排序
                maxclasscount = class_count(m)
                maxclass = m
            End If
        Next m
        If class_name(maxclass) = fold3(n, 9) Then
            correct(3) = correct(3) + 1
        End If
        accuracy(3) = correct(3) / 297
    Next n
    List1.AddItem "3rd fold : 297" & " " & "correct : " & correct(3) & " " & "accuracy : " & accuracy(3)
'---------------------------------------------------------------------------------------------------------------
    For n = 0 To 296
    Erase maxDis
        For i = 0 To 5
            maxDis(i) = 1
        Next i
        For k = 0 To 1187
            If foldDis(k, n, 4) < maxDis(0) Then
                maxDis(5) = maxDis(4)
                maxDis(4) = maxDis(3)
                maxDis(3) = maxDis(2)
                maxDis(2) = maxDis(1)
                maxDis(1) = maxDis(0)
                maxDis(0) = foldDis(k, n, 4)
                maxInd(5) = maxInd(4)
                maxInd(4) = maxInd(3)
                maxInd(3) = maxInd(2)
                maxInd(2) = maxInd(1)
                maxInd(1) = maxInd(0)
                maxInd(0) = k
            ElseIf foldDis(k, n, 4) >= maxDis(0) And foldDis(k, n, 4) <= maxDis(1) Then
                maxDis(5) = maxDis(4)
                maxDis(4) = maxDis(3)
                maxDis(3) = maxDis(2)
                maxDis(2) = maxDis(1)
                maxDis(1) = foldDis(k, n, 4)
                maxInd(5) = maxInd(4)
                maxInd(4) = maxInd(3)
                maxInd(3) = maxInd(2)
                maxInd(2) = maxInd(1)
                maxInd(1) = k
            ElseIf foldDis(k, n, 4) >= maxDis(1) And foldDis(k, n, 4) <= maxDis(2) Then
                maxDis(5) = maxDis(4)
                maxDis(4) = maxDis(3)
                maxDis(3) = maxDis(2)
                maxDis(2) = foldDis(k, n, 4)
                maxInd(5) = maxInd(4)
                maxInd(4) = maxInd(3)
                maxInd(3) = maxInd(2)
                maxInd(2) = k
            ElseIf foldDis(k, n, 4) >= maxDis(2) And foldDis(k, n, 4) <= maxDis(3) Then
                maxDis(5) = maxDis(4)
                maxDis(4) = maxDis(3)
                maxDis(3) = foldDis(k, n, 4)
                maxInd(5) = maxInd(4)
                maxInd(4) = maxInd(3)
                maxInd(3) = k
            ElseIf foldDis(k, n, 4) >= maxDis(3) And foldDis(k, n, 4) <= maxDis(4) Then
                maxDis(5) = maxDis(4)
                maxDis(4) = foldDis(k, n, 4)
                maxInd(5) = maxInd(4)
                maxInd(4) = k
            ElseIf foldDis(k, n, 4) >= maxDis(4) And foldDis(k, n, 4) <= maxDis(5) Then
                maxDis(5) = foldDis(k, n, 4)
                maxInd(5) = k
            End If
        Next k
        '取前三近的data
    Erase class_count
        For t = 0 To 5
            Select Case trainingArray4(maxInd(t), 9)
                Case "CYT"
                    class_count(1) = class_count(1) + 1
                Case "ERL"
                    class_count(2) = class_count(2) + 1
                Case "EXC"
                    class_count(3) = class_count(3) + 1
                Case "ME1"
                    class_count(4) = class_count(4) + 1
                Case "ME2"
                    class_count(5) = class_count(5) + 1
                Case "ME3"
                   class_count(6) = class_count(6) + 1
                Case "MIT"
                   class_count(7) = class_count(7) + 1
                Case "NUC"
                   class_count(8) = class_count(8) + 1
                Case "POX"
                    class_count(9) = class_count(9) + 1
                Case "VAC"
                    class_count(10) = class_count(10) + 1
            End Select
        Next t
        maxclasscount = 0
        maxclass = 0 '要預測的class值
        For m = 1 To 10
            If class_count(m) > maxclasscount Then '由小到大排序
                maxclasscount = class_count(m)
                maxclass = m
            End If
        Next m
        If class_name(maxclass) = fold4(n, 9) Then
            correct(4) = correct(4) + 1
        End If
        accuracy(4) = correct(4) / 297
    Next n
    List1.AddItem "4th fold : 297" & " " & "correct : " & correct(4) & " " & "accuracy : " & accuracy(4)
'---------------------------------------------------------------------------------------------------------------
    For n = 0 To 296
        Erase maxDis
        For i = 0 To 5
            maxDis(i) = 1
        Next i
        For k = 0 To 1187
            If foldDis(k, n, 5) < maxDis(0) Then
                maxDis(5) = maxDis(4)
                maxDis(4) = maxDis(3)
                maxDis(3) = maxDis(2)
                maxDis(2) = maxDis(1)
                maxDis(1) = maxDis(0)
                maxDis(0) = foldDis(k, n, 5)
                maxInd(5) = maxInd(4)
                maxInd(4) = maxInd(3)
                maxInd(3) = maxInd(2)
                maxInd(2) = maxInd(1)
                maxInd(1) = maxInd(0)
                maxInd(0) = k
            ElseIf foldDis(k, n, 5) >= maxDis(0) And foldDis(k, n, 5) <= maxDis(1) Then
                maxDis(5) = maxDis(4)
                maxDis(4) = maxDis(3)
                maxDis(3) = maxDis(2)
                maxDis(2) = maxDis(1)
                maxDis(1) = foldDis(k, n, 5)
                maxInd(5) = maxInd(4)
                maxInd(4) = maxInd(3)
                maxInd(3) = maxInd(2)
                maxInd(2) = maxInd(1)
                maxInd(1) = k
            ElseIf foldDis(k, n, 5) >= maxDis(1) And foldDis(k, n, 5) <= maxDis(2) Then
                maxDis(5) = maxDis(4)
                maxDis(4) = maxDis(3)
                maxDis(3) = maxDis(2)
                maxDis(2) = foldDis(k, n, 5)
                maxInd(5) = maxInd(4)
                maxInd(4) = maxInd(3)
                maxInd(3) = maxInd(2)
                maxInd(2) = k
            ElseIf foldDis(k, n, 5) >= maxDis(2) And foldDis(k, n, 5) <= maxDis(3) Then
                maxDis(5) = maxDis(4)
                maxDis(4) = maxDis(3)
                maxDis(3) = foldDis(k, n, 5)
                maxInd(5) = maxInd(4)
                maxInd(4) = maxInd(3)
                maxInd(3) = k
            ElseIf foldDis(k, n, 5) >= maxDis(3) And foldDis(k, n, 5) <= maxDis(4) Then
                maxDis(5) = maxDis(4)
                maxDis(4) = foldDis(k, n, 5)
                maxInd(5) = maxInd(4)
                maxInd(4) = k
            ElseIf foldDis(k, n, 5) >= maxDis(4) And foldDis(k, n, 5) <= maxDis(5) Then
                maxDis(5) = foldDis(k, n, 5)
                maxInd(5) = k
            End If
        Next k
        '取前三近的data
    Erase class_count
        For t = 0 To 5
            Select Case trainingArray5(maxInd(t), 9)
                Case "CYT"
                    class_count(1) = class_count(1) + 1
                Case "ERL"
                    class_count(2) = class_count(2) + 1
                Case "EXC"
                    class_count(3) = class_count(3) + 1
                Case "ME1"
                    class_count(4) = class_count(4) + 1
                Case "ME2"
                    class_count(5) = class_count(5) + 1
                Case "ME3"
                   class_count(6) = class_count(6) + 1
                Case "MIT"
                   class_count(7) = class_count(7) + 1
                Case "NUC"
                   class_count(8) = class_count(8) + 1
                Case "POX"
                    class_count(9) = class_count(9) + 1
                Case "VAC"
                    class_count(10) = class_count(10) + 1
            End Select
        Next t
        maxclasscount = 0
        maxclass = 0 '要預測的class值
        For m = 1 To 10
            If class_count(m) > maxclasscount Then '由小到大排序
                maxclasscount = class_count(m)
                maxclass = m
            End If
        Next m
Dim average As Single
        If class_name(maxclass) = fold5(n, 9) Then
            correct(5) = correct(5) + 1
        End If
        accuracy(5) = correct(5) / 296
    Next n
    List1.AddItem "5th fold : 296" & " " & "correct : " & correct(5) & " " & "accuracy : " & accuracy(5)
    average = (accuracy(1) + accuracy(2) + accuracy(3) + accuracy(4) + accuracy(5)) / 5
    List1.AddItem "average accuracy : " & average
End Sub
Private Sub Partition_click()
    List1.Clear
'    forward.Enabled = False
'    backward.Enabled = False
    'check whether the file name is empty
    If infile.Text = "" Then
        MsgBox "Please input the file names!", , "File Name"
        infile.SetFocus
    Else
        in_file = App.Path & "\" & infile.Text
        'check whether the data file exists
        If Dir(in_file) = "" Then
            MsgBox "Input file not found!", , "File Name"
            infile.SetFocus
        Else
            Open in_file For Input As #1
            fileCount = 0
            '讀檔並存入二維陣列
            Do While Not EOF(1)
                Dim tmpline As String
                Dim inputdata() As String
                Dim s As Integer, i As Integer
                
                Line Input #1, tmpline
                tmpline = Replace(tmpline, "    ", " ")
                tmpline = Replace(tmpline, "   ", " ")
                tmpline = Replace(tmpline, "  ", " ")
                inputdata = Split(tmpline, " ")
                For i = 0 To UBound(inputdata)
                    fileArray(s, i) = inputdata(i)
                    learnArray(s, i) = inputdata(i)
                Next i
                s = s + 1 '用s紀錄1484筆，用i記錄每筆的10個資料
                fileCount = fileCount + 1
            Loop
            List1.AddItem "成功! " & " 共" & fileCount & "筆資料"
            Close #1
        End If
    End If
End Sub
