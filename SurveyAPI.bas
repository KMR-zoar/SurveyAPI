Attribute VB_Name = "SurveyAPI"
Function xy2blapi(X座標 As String, Y座標 As String, 系番号 As String, 出力データ As Long) As String
   Dim URL As String
   Dim CoordinateX As String
   Dim CoordinateY As String
   Dim Zone As String
   Dim ResText As Object
   Dim Err As Object
   Dim Lat As Object
   Dim Lon As Object
   Dim TrueNorth As Object
   Dim ScaleFactor As Object
   Dim Path2Lat As String
   Dim Path2Lon As String
   Dim Path2TN As String
   Dim Path2SF As String
   
   URL = "http://vldb.gsi.go.jp/sokuchi/surveycalc/surveycalc/xy2bl.pl"
   URL = URL & "?outputType=xml&refFrame=2"
   
   CoordinateX = X座標
   CoordinateY = Y座標
   Zone = 系番号
   
   URL = URL & "&zone=" & Zone & "&publicX=" & CoordinateX & "&publicY=" & CoordinateY
   
   Set ResText = CreateObject("MSXML2.DOMDocument")
   ResText.async = False
   
   If (ResText.Load(URL) = False) Then
      MsgBox "APIにアクセスできませんでした"
      Exit Function
   End If

   Path2Lat = "ExportData/OutputData/latitude"
   Path2Lon = "ExportData/OutputData/longitude"
   Path2TN = "ExportData/OutputData/gridConv"
   Path2SF = "ExportData/OutputData/scaleFactor"
   
   Set Lat = ResText.getElementsBytagName(Path2Lat).Item(0)
   Set Lon = ResText.getElementsBytagName(Path2Lon).Item(0)
   Set TrueNorth = ResText.getElementsBytagName(Path2TN).Item(0)
   Set ScaleFactor = ResText.getElementsBytagName(Path2SF).Item(0)
   
   On Error GoTo ErrorSet
   
   Select Case 出力データ
      Case 0
      xy2blapi = Lat.Text
      Case 1
      xy2blapi = Lon.Text
      Case 2
      xy2blapi = TrueNorth.Text
      Case 3
      xy2blapi = ScaleFactor.Text
      Case Else
      xy2blapi = "出力データの指定が正しくありません"
   End Select
   
   Exit Function
   
ErrorSet:
   
   Set Err = ResText.getElementsBytagName("ExportData/ErrMsg").Item(0)
   MsgBox Err.Text
   
End Function

Function bl2xyapi(緯度 As String, 経度 As String, 系番号 As String, 出力データ As Long) As String
   Dim URL As String
   Dim Lat As String
   Dim Lon As String
   Dim Zone As String
   Dim ResText As Object
   Dim Err As Object
   Dim CoordinateX As Object
   Dim CoordinateY As Object
   Dim TrueNorth As Object
   Dim ScaleFactor As Object
   Dim Path2X As String
   Dim Path2Y As String
   Dim Path2TN As String
   Dim Path2SF As String
   
   URL = "http://vldb.gsi.go.jp/sokuchi/surveycalc/surveycalc/bl2xy.pl"
   URL = URL & "?outputType=xml&refFrame=2"
   
   Lat = 緯度
   Lon = 経度
   Zone = 系番号
   
   URL = URL & "&zone=" & Zone & "&latitude=" & Lat & "&longitude=" & Lon
   
   Set ResText = CreateObject("MSXML2.DOMDocument")
   ResText.async = False
   
   If (ResText.Load(URL) = False) Then
      MsgBox "APIにアクセスできませんでした"
      Exit Function
   End If

   Path2X = "ExportData/OutputData/publicX"
   Path2Y = "ExportData/OutputData/publicY"
   Path2TN = "ExportData/OutputData/gridConv"
   Path2SF = "ExportData/OutputData/scaleFactor"
   
   Set CoordinateX = ResText.getElementsBytagName(Path2X).Item(0)
   Set CoordinateY = ResText.getElementsBytagName(Path2Y).Item(0)
   Set TrueNorth = ResText.getElementsBytagName(Path2TN).Item(0)
   Set ScaleFactor = ResText.getElementsBytagName(Path2SF).Item(0)
   
   On Error GoTo ErrorSet
   
   Select Case 出力データ
      Case 0
      bl2xyapi = CoordinateX.Text
      Case 1
      bl2xyapi = CoordinateY.Text
      Case 2
      bl2xyapi = TrueNorth.Text
      Case 3
      bl2xyapi = ScaleFactor.Text
      Case Else
      bl2xyapi = "出力データの指定が正しくありません"
   End Select
   
   Exit Function
   
ErrorSet:
   
   Set Err = ResText.getElementsBytagName("ExportData/ErrMsg").Item(0)
   MsgBox Err.Text
   
End Function

