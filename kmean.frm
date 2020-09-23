VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "k mean clustering"
   ClientHeight    =   6045
   ClientLeft      =   5310
   ClientTop       =   4665
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   7155
   Begin VB.CommandButton cmdReset 
      Caption         =   "Clear Data"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txtNumCluster 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Text            =   "3"
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   5055
      Left            =   0
      ScaleHeight     =   333
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   477
      TabIndex        =   0
      Top             =   840
      Width           =   7215
      Begin VB.Label lblCentroid 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         Caption         =   "1"
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   5
         Top             =   240
         Width           =   90
      End
   End
   Begin VB.Label lblXYValue 
      AutoSize        =   -1  'True
      Caption         =   "X,Y"
      Height          =   195
      Left            =   4800
      TabIndex        =   7
      Top             =   360
      Width           =   255
   End
   Begin VB.Label lblXY 
      AutoSize        =   -1  'True
      Caption         =   "(X,Y)="
      Height          =   195
      Left            =   4200
      TabIndex        =   6
      Top             =   360
      Width           =   435
   End
   Begin VB.Label lblExplanation 
      AutoSize        =   -1  'True
      Caption         =   "Click data in the picture box below. The program will automatically cluster the data"
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Number of cluster"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   1245
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ###################################################################
' K-MEAN CLUSTERING TUTORIAL
' BY: Kardi Teknomo (email: kardi@plan.civil.tohoku.ac.jp)
' Update: 2001 Sept 5
' When User click picture box to input new data (X,Y), the program
' will make group/cluster the data by minimizing the sum of squares
' of distances between data and the corresponding cluster centroids.
' This algorithm is used for unsupervised learning of Neural network,
' Pattern recognitions, Classification analysis etc.
' ###################################################################


Private Data()                ' Row 0 = cluster, 1 =X, 2= Y; data in columns
Private Centroid() As Single  ' centroid (X and Y) of clusters; cluster number = column number
Private totalData As Integer  ' total number of data (total columns)
Private numCluster As Integer ' total number of cluster


' ###################################################################
' CONTROLS
' + Form_Load
' + cmdReset_Click
' + txtNumCluster_Change
' + Picture1_MouseDown
' + Picture1_MouseMove
' ###################################################################

Private Sub Form_Load()
Dim i As Integer

    Picture1.BackColor = &HFFFFFF   ' white
    Picture1.DrawWidth = 10         ' big dot
    Picture1.ScaleMode = 3          ' pixels
    lblExplanation.Caption = "Click data in the picture box below. The program will automatically cluster the data by color code"
    
    'take number of cluster
    numCluster = Int(txtNumCluster)
    ReDim Centroid(1 To 2, 1 To numCluster)
'    lblCentroid(0).Visible = False
'    lblCentroid(0).Caption = 1
    For i = 0 To numCluster - 1
        'create label
        If i > 0 Then Load lblCentroid(i)
        lblCentroid(i).Caption = i + 1
        lblCentroid(i).Visible = False
    Next i
End Sub


Private Sub cmdReset_Click()
' reset data
Dim i As Integer

    Picture1.Cls        ' clean picture
    Erase Data          ' remove data
    totalData = 0
    
    For i = 0 To numCluster - 1
        lblCentroid(i).Visible = False  ' don't show label
    Next i
    
    'enable to change the number of cluster
    txtNumCluster.Enabled = True
End Sub

Private Sub txtNumCluster_Change()
'change number of cluster and reset data
Dim i As Integer

    For i = 1 To numCluster - 1
        Unload lblCentroid(i)
    Next i
    numCluster = Int(txtNumCluster)
    ReDim Centroid(1 To 2, 1 To numCluster)
    'Call cmdReset_Click
    For i = 0 To numCluster - 1
        If i > 0 Then Load lblCentroid(i)
        lblCentroid(i).Caption = i + 1
        lblCentroid(i).Visible = False
    Next i
End Sub


Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'collecting data and showing result
Dim colorCluster As Integer
Dim i As Integer
    
    'disable to change the number of cluster
    txtNumCluster.Enabled = False
    
    ' take feature data
    totalData = totalData + 1
    ReDim Preserve Data(0 To 2, 1 To totalData)  ' notice: start with 0 for row
    Data(1, totalData) = X
    Data(2, totalData) = Y
    
    'do k-mean clustering
    Call kMeanCluster(Data, numCluster)
    
    'show the result
    Picture1.Cls
    For i = 1 To totalData
        colorCluster = Data(0, i) - 1
        If colorCluster = 7 Then colorCluster = 12   ' if white (similar to background change to other color)
        X = Data(1, i)
        Y = Data(2, i)
        Picture1.PSet (X, Y), QBColor(colorCluster)
    Next i
    
    'show centroid
    For i = 1 To min2(numCluster, totalData)
        lblCentroid(i - 1).Left = Centroid(1, i)
        lblCentroid(i - 1).Top = Centroid(2, i)
        lblCentroid(i - 1).Visible = True
    Next i
End Sub


Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblXYValue.Caption = X & "," & Y
End Sub


' ###################################################################
' FUNCTIONS
' + kMeanCluster:
' + dist: calculate distance
' + min2: return minimum value between two numbers
' ###################################################################

Sub kMeanCluster(Data() As Variant, numCluster As Integer)
' main function to cluster data into k number of Clusters
' input: + Data matrix (0 to 2, 1 to TotalData); Row 0 = cluster, 1 =X, 2= Y; data in columns
'        + numCluster: number of cluster user want the data to be clustered
'        + private variables: Centroid, TotalData
' ouput: o) update centroid
'        o) assign cluster number to the Data (= row 0 of Data)
Dim i As Integer
Dim j As Integer
Dim X As Single
Dim Y As Single
Dim min As Single
Dim cluster As Integer
Dim d As Single
Dim sumXY()
Dim isStillMoving As Boolean

isStillMoving = True

If totalData <= numCluster Then
    Data(0, totalData) = totalData               ' cluster No = total data
    Centroid(1, totalData) = Data(1, totalData)  ' X
    Centroid(2, totalData) = Data(2, totalData)  ' Y
Else
    'calculate minimum distance to assign the new data
    min = 10 ^ 10                                'big number
    X = Data(1, totalData)
    Y = Data(2, totalData)
    For i = 1 To numCluster
        d = dist(X, Y, Centroid(1, i), Centroid(2, i))
        If d < min Then
            min = d
            cluster = i
        End If
    Next i
    Data(0, totalData) = cluster
    
    Do While isStillMoving
    ' this loop will surely convergent
    
        'calculate new centroids
        ReDim sumXY(1 To 3, 1 To numCluster)    ' 1 =X, 2=Y, 3=count number of data
        For i = 1 To totalData
            sumXY(1, Data(0, i)) = Data(1, i) + sumXY(1, Data(0, i))
            sumXY(2, Data(0, i)) = Data(2, i) + sumXY(2, Data(0, i))
            sumXY(3, Data(0, i)) = 1 + sumXY(3, Data(0, i))
        Next i
        For i = 1 To numCluster
            Centroid(1, i) = sumXY(1, i) / sumXY(3, i)
            Centroid(2, i) = sumXY(2, i) / sumXY(3, i)
        Next i
        
        
        'assign all data to the new centroids
        isStillMoving = False
        For i = 1 To totalData
            min = 10 ^ 10                                'big number
            X = Data(1, i)
            Y = Data(2, i)
            For j = 1 To numCluster
                d = dist(X, Y, Centroid(1, j), Centroid(2, j))
                If d < min Then
                    min = d
                    cluster = j
                End If
            Next j
            If Data(0, i) <> cluster Then
                Data(0, i) = cluster
                isStillMoving = True
            End If
        Next i
    Loop
End If
End Sub


Function dist(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single) As Single
' calculate Euclidean distance
    dist = Sqr((Y2 - Y1) ^ 2 + (X2 - X1) ^ 2)
End Function


Private Function min2(num1, num2)
' return minimum value between two numbers
    If num1 < num2 Then
        min2 = num1
    Else
        min2 = num2
    End If
End Function
