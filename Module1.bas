Attribute VB_Name = "Module11"
Function geoazymut(xpocz As Double, ypocz As Double, xkonc As Double, ykonc As Double) As Double
Dim dx As Double, dy As Double
Pi = Atn(1) * 4
dx = xkonc - xpocz
dy = ykonc - ypocz
If (dx = 0) And (dy = 0) Then azym = 0 Else
            If (dx = 0) And (dy > 0) Then azym = Pi / 2 Else
            If (dx = 0) And (dy < 0) Then azym = Pi * 3 / 2 Else
            If (dx > 0) And (dy = 0) Then azym = 0 Else
            If (dx < 0) And (dy = 0) Then azym = Pi Else
            If (dx > 0) And (dy > 0) Then azym = Atn((dy ^ 2) ^ 0.5 / (dx ^ 2) ^ 0.5) Else
            If (dx < 0) And (dy > 0) Then azym = Pi - Atn((dy ^ 2) ^ 0.5 / (dx ^ 2) ^ 0.5) Else
            If (dx < 0) And (dy < 0) Then azym = Pi + Atn((dy ^ 2) ^ 0.5 / (dx ^ 2) ^ 0.5) Else
            If (dx > 0) And (dy < 0) Then azym = 2 * Pi - Atn((dy ^ 2) ^ 0.5 / (dx ^ 2) ^ 0.5)
          
          
          geoazymut = azym
        
            
End Function

Function odleglosc(xpocz As Double, ypocz As Double, xkonc As Double, ykonc As Double) As Double
Dim dx, dy As Double
dx = xkonc - xpocz
dy = ykonc - ypocz
odleglosc = (dx ^ 2 + dy ^ 2) ^ 0.5

End Function


Function przypisaniexy(obs As String, nr As String, xy As Double) As Double

przypisaniexy = xy
End Function


Function tworzenieA(Pkt As String, dx_dy As String, c As String, l As String, p As String, x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double) As Double
Dim rodz, x_y, A, Az, d12, d13, Al, Ap, Bl, Bp As Double

Pi = Atn(1) * 4
If (Pkt <> c) And (Pkt <> l) And (Pkt <> p) Then tworzenieA = 0 Else


If p = "" Then rodz = 1 Else rodz = 2 'dlugosc lub kat'
If (dx_dy = "dx") Or (dx_dy = "Dx") Then x_y = 1 Else If (dx_dy = "dy") Or (dx_dy = "Dx") Then x_y = 2 Else x_y = 3 'wspolrzedna dx lub dy
If rodz = 1 Then 'jesli obserwacja jest dlugoscia wtedy liczy tylko alfa'
Az = geoazymut(x1, y1, x2, y2)
Else 'a jesli to kat to liczy A B'
d12 = odleglosc(x1, y1, x2, y2)
d13 = odleglosc(x1, y1, x3, y3)
Al = 2000 / Pi * (x2 - x1) / d12 ^ 2
Bl = 2000 / Pi * (y2 - y1) / d12 ^ 2
Ap = 2000 / Pi * (x3 - x1) / d13 ^ 2
Bp = 2000 / Pi * (y3 - y1) / d13 ^ 2
End If


'no i trzeba umiejscowic odpowiednio w zaleznosci od tego czy dx czy dy'
 If (rodz = 2) And (x_y = 1) And (Pkt = c) Then tworzenieA = -(Bl - Bp) Else
    If (rodz = 2) And (x_y = 1) And (Pkt = l) Then tworzenieA = Bl Else
        If (rodz = 2) And (x_y = 1) And (Pkt = p) Then tworzenieA = -Bp Else
 If (rodz = 2) And (x_y = 2) And (Pkt = c) Then tworzenieA = (Al - Ap) Else
    If (rodz = 2) And (x_y = 2) And (Pkt = l) Then tworzenieA = -Al Else
        If (rodz = 2) And (x_y = 2) And (Pkt = p) Then tworzenieA = Ap Else
 If (rodz = 1) And (x_y = 1) And (Pkt = c) Then tworzenieA = -Cos(Az) Else
    If (rodz = 1) And (x_y = 1) And (Pkt = l) Then tworzenieA = Cos(Az) Else
 If (rodz = 1) And (x_y = 2) And (Pkt = c) Then tworzenieA = -Sin(Az) Else
    If (rodz = 1) And (x_y = 2) And (Pkt = l) Then tworzenieA = Sin(Az)
       
End Function

Public Function TworzenieL()

End Function


Function bledy_srednie(A As Range, tyt_z_A As Range, mac_wynik As Range)
Const ilosc_punktow = 120
     

ka = A.Columns.Count


Dim x(ilosc_punktow, 1) As Single

'znajdywanie mx i my w przekatnej macierzy ATA-1'
For m = 1 To ka
For i = 1 To ka
    If mac_wynik(m, 1) = tyt_z_A(2, i) Then
        If tyt_z_A(1, i) = "dx" Then
            x(m - 1, 0) = A(i, i) ^ 0.5
        Else
            x(m - 1, 1) = A(i, i) ^ 0.5
        End If
    End If
Next i
Next m
bledy_srednie = x
End Function

Function elipsy(A As Range, tyt_z_A As Range, mac_wynik As Range)
Const ilosc_punktow = 120
'liczenie A B fi z macierzy ATA-1'
Pi = Atn(1) * 4
ka = A.Columns.Count

Dim x(ilosc_punktow, 2) As Double

For m = 1 To ka / 2
For i = 1 To ka
    If mac_wynik(m, 1) = tyt_z_A(2, i) Then
        If tyt_z_A(1, i) = "dx" Then
            mx = A(i, i)
            q1 = i
        Else
            my = A(i, i)
            q2 = i
        End If
    End If
    
Next i
mxy = A(q1, q2)
If mx = 0 And my = 0 And mxy = 0 Then
    x(m - 1, 0) = 0
    x(m - 1, 1) = 0
    x(m - 1, 2) = 0
Else
    Ae = ((mx + my) / 2 + ((mx - my) ^ 2 / 4 + mxy ^ 2) ^ 0.5) ^ 0.5
    Be1 = ((mx + my) / 2 - ((mx - my) ^ 2 / 4 + (mxy) ^ 2) ^ 0.5)
   ' przypadek gdy B^2<0 to nie chce wyciagnac pierwiastka, kiedy cos sie takiego dzieje?'
   If Be1 < 0 Then Be = 0 Else Be = Be1 ^ 0.5
    Fi = geoazymut(0, 0, mx - my, 2 * mxy) / 2 * 200 / Pi
    x(m - 1, 0) = Ae
    x(m - 1, 1) = Be
    x(m - 1, 2) = Fi
End If
Next m
elipsy = x
End Function

Function Tworzenie_T(Dane As Range, t0 As Double)
Const maks_iloœæ_obserwacji = 50


Dim x(maks_iloœæ_obserwacji, maks_iloœæ_obserwacji) As Double

w = Dane.Rows.Count
For i = 1 To w
    x(i - 1, i - 1) = Dane(i, 1) + Dane(i, 2) / 60 - t0
Next i
Tworzenie_T = x
End Function
