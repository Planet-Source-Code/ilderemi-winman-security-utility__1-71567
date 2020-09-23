VERSION 5.00
Begin VB.Form frmGeoCalc 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6735
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox MyMaskDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0015160C&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   120
      Width           =   2655
   End
   Begin VB.TextBox MyLatitude 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0015160C&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   480
      Width           =   2655
   End
   Begin VB.TextBox MyLongitude 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0015160C&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox MyTimeDiff 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0015160C&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   1635
      Width           =   1335
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Left            =   1305
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   3
      Left            =   1305
      Top             =   110
      Width           =   2685
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   2
      Left            =   1310
      Top             =   470
      Width           =   2685
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   315
      Index           =   1
      Left            =   1310
      Top             =   830
      Width           =   2685
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   310
      Index           =   0
      Left            =   1310
      Top             =   1190
      Width           =   2680
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      Top             =   1635
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   375
      Left            =   2655
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   1815
      Left            =   4080
      TabIndex        =   8
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Latitude:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Longitude:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "TimeDiff:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0000FF00&
      Height          =   2055
      Left            =   0
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmGeoCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'|\\\\\\\\____________________////////|'
'||\\\\\\\\__________________////////||'
'|||\\\\\\\\________________////////|||'
'||||\\\\\\\\______________////////||||'
'|||||\\\\\\\\____________////////|||||'
'||||||\\\\\\\\__________////////||||||'
'|||||| \\\\\\\\________//////// ||||||'
'||||||  \\\\\\\\______////////  ||||||'
'||||||   \\\\\\\\____////////   ||||||'
'||||||    \\\\\\\\  ////////    ||||||'
'||||||     \\\\\\\\////////     ||||||'
'||||||                          ||||||'
'||||||                          ||||||'
'||||||    !!!!!!!!!!!!!!!!!!    ||||||'
'||||||    ! Amazing WinMan !    ||||||'
'||||||    !!!!!!!!!!!!!!!!!!    ||||||'
'||||||            !!            ||||||'
'||||||            !!            ||||||'
'||||||            !!            ||||||'
'||||||            !!            ||||||'
'||||||........:iLDEREMi:........||||||'
Const pi = 3.14159265358979
Public degrees, radians As Variant

Function GetRange(x)

    Dim temp1
    Dim temp2
    temp1 = x / (2 * pi)
    temp2 = (2 * pi) * (temp1 - Fix(temp1))


    If temp2 < 0 Then
        temp2 = (2 * pi) + temp2
    End If

    GetRange = temp2
End Function

Function GetMilitaryTime(DecimalTime, GMTOffset)

    
    Dim temp1
    Dim temp2
    ' Handle 24-hour time wrap
    If DecimalTime + GMTOffset < 0 Then DecimalTime = DecimalTime + 24
    If DecimalTime + GMTOffset > 24 Then DecimalTime = DecimalTime - 24
    temp1 = Abs(DecimalTime + GMTOffset)
    temp2 = Int(temp1)
    temp1 = 60 * (temp1 - temp2)
    temp1 = Right("0000" & CStr(Int(temp2 * 100 + temp1 + 0.5)), 4)
    GetMilitaryTime = Left(temp1, 2) & ":" & Right(temp1, 2)
End Function

Function GetSunRiseSet(latitude, ByVal longitude, ZoneRelativeGMT, RiseOrSet, Year, Month, Day)



    If Abs(latitude) > 63 Then
        GetSunRiseSet = "{invalid latitude}"
        Exit Function
    End If

    Y = Year
    m = Month
    d = Day
    altitude = -0.833


    Select Case UCase(RiseOrSet)
        Case "S"
        RS = -1
        Case Else
        RS = 1
    End Select

Ephem2000Day = 367 * Y - 7 * (Y + (m + 9) \ 12) \ 4 + 275 * m \ 9 + d - 730531.5
utold = pi
utnew = 0
sinalt = CDbl(Sin(altitude * radians)) ' solar altitude
sinphi = CDbl(Sin(latitude * radians)) ' viewer's latitude
cosphi = CDbl(Cos(latitude * radians)) '
longitude = CDbl(longitude * radians) ' viewer's longitude
Err.clear
On Error Resume Next


Do While (Abs(utold - utnew) > 0.001) And (ct < 35)
    ct = ct + 1
    utold = utnew
    days = Ephem2000Day + utold / (2 * pi)
    t = days / 36525
    ' These 'magic' numbers are orbital elem
    '     ents of the sun, and should not be chang
    '     ed
    l = GetRange(4.8949504201433 + 628.331969753199 * t)
    g = GetRange(6.2400408 + 628.3019501 * t)
    ec = 0.033423 * Sin(g) + 0.00034907 * Sin(2# * g)
    lambda = l + ec
    E = -1 * ec + 0.0430398 * Sin(2# * lambda) - 0.00092502 * Sin(4# * lambda)
    obl = 0.409093 - 0.0002269 * t
    ' Obtain ASIN of (SIN(obl) * SIN(lambda)
    '     )
    Delta = Sin(obl) * Sin(lambda)
    Delta = Atn(Delta / (Sqr(1 - Delta * Delta)))
    GHA = utold - pi + E
    cosc = (sinalt - sinphi * Sin(Delta)) / (cosphi * Cos(Delta))


    Select Case cosc
        Case cosc > 1
        correction = 0
        Case cosc < -1
        correction = pi
        Case Else
        correction = Atn((Sqr(1 - cosc * cosc)) / cosc)
    End Select

utnew = GetRange(utold - (GHA + longitude + RS * correction))
Loop



If Err = 0 Then
GetSunRiseSet = GetMilitaryTime(utnew * degrees / 15, ZoneRelativeGMT)
Else
GetSunRiseSet = "{err}"
End If

End Function



Private Sub Command1_Click()

    If IsDate(MyMaskDate) = False Then Exit Sub
    Y = Year(MyMaskDate)
    m = Month(MyMaskDate)
    d = Day(MyMaskDate)

    
    If MyTimeDiff = "" Then MyTimeDiff = 4.5

    If MyLatitude = "" Then 'This is the setting For brooklyn NY
        MyLatitude = 36.18
        MyLongitude = 59.36
    End If

    ' Set this to your offset from GMT (e.g.


    '     for Dallas is -6)
        ' NOTE: The routine does NOT handle swit
        '     ches to/from daylight savings
        'time, so beware!
        MyTimeZone = Val(MyTimeDiff) '-5
        ' Note:Set RiseOrSet to "R" for sunrise,
        '     "S" for sunset
        RiseOrSet = "R"
        Rise = GetSunRiseSet(MyLatitude, MyLongitude, MyTimeZone, _
        RiseOrSet, Y, m, d)
        SSEt = GetSunRiseSet(MyLatitude, MyLongitude, MyTimeZone, _
        "s", Y, m, d)
        Label1.Caption = WeekdayName(Weekday(Date), False) & ", " & Day(Date) & ", " & MonthName(Month(Date), False) & ", " & Year(Date) & vbNewLine & "Sunrise: " & Format(Rise, "H:nn AMPM") & vbNewLine & "Sunset: " & Format(SSEt, "H:nn AMPM")
    End Sub

Private Sub Form_Load()

SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
  Set myShadow = New clsShadow
  With myShadow
    If .Shadow(Me) Then
      .Depth = 10
      .Transparency = 128
    Else
      Set myShadow = Nothing
    End If
  End With

MyMaskDate = Date
    degrees = 180 / pi
    radians = pi / 180
Command1_Click
End Sub

Private Sub Label6_Click()
Unload Me
End Sub
