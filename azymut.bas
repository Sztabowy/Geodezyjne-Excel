Attribute VB_Name = "Module1"
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
Dim dx As Double, dy As Double
dx = xkonc - xpocz
dy = ykonc - ypocz
odleglosc = (dx ^ 2 + dy ^ 2) ^ 0.5

End Function

