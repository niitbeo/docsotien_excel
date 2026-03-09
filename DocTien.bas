Attribute VB_Name = "DocTienModule"
Option Explicit

' Chuyen doi so thanh chu tieng Viet CO DAU
' Su dung: =DocTien(A1) hoac =DocTien(12345)
Public Function DocTien(ByVal SoTien As Variant) As String
    On Error GoTo ErrorHandler
    
    If IsEmpty(SoTien) Or SoTien = "" Then
        DocTien = ""
        Exit Function
    End If
    
    Dim So As Double
    So = CDbl(SoTien)
    
    If So < 0 Then
        DocTien = ChrW(194) & "m " & DocSo(Abs(So))  ' Ã‚m
    ElseIf So = 0 Then
        DocTien = "Kh" & ChrW(244) & "ng " & ChrW(273) & ChrW(7891) & "ng"  ' KhÃ´ng Ä‘á»“ng
    Else
        DocTien = DocSo(So)
    End If
    
    Exit Function
    
ErrorHandler:
    DocTien = "#L" & ChrW(7895) & "I: Gi" & ChrW(225) & " tr" & ChrW(7883) & " kh" & ChrW(244) & "ng h" & ChrW(7907) & "p l" & ChrW(7879)
End Function

Private Function DocSo(ByVal So As Double) As String
    Dim strPhanNguyen As String
    Dim strPhanThapPhan As String
    Dim dblNguyen As Double
    Dim dblThapPhan As Double
    
    dblNguyen = Int(So)
    dblThapPhan = Round((So - dblNguyen) * 100, 0)
    
    If dblThapPhan >= 100 Then
        dblNguyen = dblNguyen + 1
        dblThapPhan = dblThapPhan - 100
    End If
    
    strPhanNguyen = DocPhanNguyen(dblNguyen)
    
    If dblThapPhan > 0 Then
        strPhanThapPhan = " ph" & ChrW(7849) & "y " & DocPhanNguyen(dblThapPhan)  ' pháº©y
        DocSo = strPhanNguyen & strPhanThapPhan & " " & ChrW(273) & ChrW(7891) & "ng"  ' Ä‘á»“ng
    Else
        DocSo = strPhanNguyen & " " & ChrW(273) & ChrW(7891) & "ng"  ' Ä‘á»“ng
    End If
End Function

Private Function DocPhanNguyen(ByVal So As Double) As String
    Dim Chuoi As String
    Dim Nhom(1 To 6) As Integer
    Dim Hang As Long
    Dim HangLonNhat As Long
    Dim strHang As String
    Dim CanDocDayDu As Boolean
    
    If So = 0 Then
        DocPhanNguyen = "Kh" & ChrW(244) & "ng"  ' KhÃ´ng
        Exit Function
    End If
    
    Chuoi = ""
    HangLonNhat = 0
    
    For Hang = 1 To 6
        Nhom(Hang) = So - (Int(So / 1000) * 1000)
        If Nhom(Hang) > 0 Then
            HangLonNhat = Hang
        End If
        
        So = Int(So / 1000)
        If So = 0 Then
            Exit For
        End If
    Next Hang
    
    For Hang = HangLonNhat To 1 Step -1
        If Nhom(Hang) > 0 Then
            CanDocDayDu = (Hang < HangLonNhat And Nhom(Hang) < 100)
            strHang = DocHangNguyen(Nhom(Hang), CanDocDayDu)
            
            If Chuoi <> "" Then
                Chuoi = Chuoi & " "
            End If
            
            If DonViHang(Hang) <> "" Then
                Chuoi = Chuoi & strHang & " " & DonViHang(Hang)
            Else
                Chuoi = Chuoi & strHang
            End If
        End If
    Next Hang
    
    If Len(Chuoi) > 0 Then
        Chuoi = UCase(Left(Chuoi, 1)) & Mid(Chuoi, 2)
    End If
    
    DocPhanNguyen = Chuoi
End Function

Private Function DocHangNguyen(ByVal So As Integer, Optional ByVal CanDocDayDu As Boolean = False) As String
    Dim Tram As Integer, Chuc As Integer, Donvi As Integer
    Dim strKetQua As String
    
    Tram = Int(So / 100)
    Chuc = Int((So Mod 100) / 10)
    Donvi = So Mod 10
    
    strKetQua = ""
    
    ' Hang tram
    If Tram > 0 Then
        strKetQua = DocChuSo(Tram) & " tr" & ChrW(259) & "m"  ' trÄƒm
    ElseIf CanDocDayDu And So > 0 Then
        strKetQua = DocChuSo(0) & " tr" & ChrW(259) & "m"
    End If
    
    ' Hang chuc
    If Chuc > 0 Then
        If strKetQua <> "" Then strKetQua = strKetQua & " "
        
        If Chuc = 1 Then
            strKetQua = strKetQua & "m" & ChrW(432) & ChrW(7901) & "i"  ' mÆ°á»i
        Else
            strKetQua = strKetQua & DocChuSo(Chuc) & " m" & ChrW(432) & ChrW(417) & "i"  ' mÆ°Æ¡i
        End If
    ElseIf Donvi > 0 And (Tram > 0 Or CanDocDayDu) Then
        If strKetQua <> "" Then strKetQua = strKetQua & " "
        strKetQua = strKetQua & "l" & ChrW(7867)  ' láº»
    End If
    
    ' Hang don vi
    If Donvi > 0 Then
        If strKetQua <> "" Then strKetQua = strKetQua & " "
        
        If Donvi = 5 And Chuc >= 1 Then
            strKetQua = strKetQua & "l" & ChrW(259) & "m"  ' lÄƒm
        ElseIf Donvi = 1 And Chuc >= 1 Then
            strKetQua = strKetQua & "m" & ChrW(7889) & "t"  ' má»‘t
        Else
            strKetQua = strKetQua & DocChuSo(Donvi)
        End If
    End If
    
    DocHangNguyen = strKetQua
End Function

Private Function DocChuSo(ByVal So As Integer) As String
    Select Case So
        Case 0: DocChuSo = "kh" & ChrW(244) & "ng"  ' khÃ´ng
        Case 1: DocChuSo = "m" & ChrW(7897) & "t"  ' má»™t
        Case 2: DocChuSo = "hai"
        Case 3: DocChuSo = "ba"
        Case 4: DocChuSo = "b" & ChrW(7889) & "n"  ' bá»‘n
        Case 5: DocChuSo = "n" & ChrW(259) & "m"  ' nÄƒm
        Case 6: DocChuSo = "s" & ChrW(225) & "u"  ' sÃ¡u
        Case 7: DocChuSo = "b" & ChrW(7843) & "y"  ' báº£y
        Case 8: DocChuSo = "t" & ChrW(225) & "m"  ' tÃ¡m
        Case 9: DocChuSo = "ch" & ChrW(237) & "n"  ' chÃ­n
    End Select
End Function

Private Function DonViHang(ByVal Hang As Long) As String
    Select Case Hang
        Case 1: DonViHang = ""
        Case 2: DonViHang = "ngh" & ChrW(236) & "n"  ' nghÃ¬n
        Case 3: DonViHang = "tri" & ChrW(7879) & "u"  ' triá»‡u
        Case 4: DonViHang = "t" & ChrW(7927)  ' tá»·
        Case 5: DonViHang = "ngh" & ChrW(236) & "n t" & ChrW(7927)  ' nghÃ¬n tá»·
        Case 6: DonViHang = "tri" & ChrW(7879) & "u t" & ChrW(7927)  ' triá»‡u tá»·
        Case Else: DonViHang = ""
    End Select
End Function

