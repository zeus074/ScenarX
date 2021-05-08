Attribute VB_Name = "mOpenPassword"
Option Explicit
Dim X, Y, Z, b As Integer
Dim Blk(4, 1), Table(31), nTable(31) As Byte
Public PassOpen As String

Public Function OpenPwd(Seq1 As String, bPass As String) As String
Dim n, nS, ID(5), pSeq, pSeq2(10, 1), tID As Byte, Nine As Boolean
Dim TotBlock, Indice(31, 1) As Byte
Dim T1, T2 As Double
Dim S(6) As Integer
Dim Somma, Md, Seq2 As Double

S(0) = 0: S(1) = 25: S(2) = 28: S(3) = 29: S(4) = 1: S(5) = 5
S(6) = 12
Somma = 0: n = 0: Seq1 = Replace(Seq1, "*", ""): Seq1 = Replace(Seq1, "#", "")

'trovo i 9
For X = 1 To Len(Seq1)
If Mid(Seq1, X, 1) = 9 Then n = n + 1
Next X

If n > 0 Then
    If n Mod 2 = 0 Then Seq1 = Trim(Replace(Seq1, "9", "0")) Else Seq1 = Trim(Replace(Seq1, "9", "0")): Nine = True
End If

For X = 1 To Len(Seq1)
If Mid(Seq1, X, 1) = 7 Then GoTo F7
If Mid(Seq1, X, 1) = 8 Then GoTo F8
Somma = Somma + S(Mid(Seq1, X, 1))
Cont:
Next X

Md = Somma Mod 32
'creo la tabella partendo da seq2
Y = Md
For X = 0 To 31: Table(X) = Y: Y = Y + 1: If Y > 31 Then Y = 0
Next X

Md = 2 ^ Md: Seq2 = Md

If nS > 0 Then
'trovato un 7 o 8 cavoli amari
Y = 1
For X = nS - 1 To 0 Step -1
ID(Y) = 8 - (pSeq2(X, 0) Mod 8)
Y = Y + 1
Next X
TotBlock = Y - 1

For Z = 1 To TotBlock
'blocchi
tID = ID(Z): b = 1
Do
Blk(b, 0) = tID: Blk(b, 1) = tID + 7
tID = tID + 8: b = b + 1
Loop While tID + 7 <= 31
GoSub Distorci
Next Z

'analisi tabella e trovo indici
tID = nTable(0): Indice(1, 1) = tID: Y = 1
For X = 2 To 31
If 2 ^ (X - 1) > Val(bPass) Then Exit For
If tID + 1 <> nTable(X - 1) Then Y = Y + 1: Indice(Y, 0) = X - 1: Indice(Y, 1) = nTable(X - 1)
tID = nTable(X - 1)
Next X

Dim Resto As Double
Seq2 = 0: T1 = bPass

For X = Y To 1 Step -1
If T1 < (2 ^ Indice(X, 0)) Then GoTo StepOver
Md = T1 Mod (2 ^ Indice(X, 0)): Resto = Int(T1 / (2 ^ Indice(X, 0)))
T1 = Md
Seq2 = Seq2 + Resto * (2 ^ Indice(X, 1))
StepOver:
Next X

Fine:
If Nine Then Seq2 = (2 ^ 32) - 1 - Seq2
OpenPwd = "*#" & Seq2 & "##"

Else
'moltiplicazione
T1 = (Md * bPass) / (2 ^ 32)
T2 = (Md * bPass) - Int(T1) * (2 ^ 32)
T1 = CDbl(Md * bPass): T2 = CDbl(2 ^ 32)
Seq2 = MyMod(T1, T2) + Int(T1 / T2)

If Nine Then Seq2 = (2 ^ 32) - 1 - Seq2
OpenPwd = "*#" & Seq2 & "##"
End If
Exit Function


Distorci:

For X = 0 To 31: nTable(X) = Table(X): Next X

'distorsione uso pSeq2(X, 0)
b = nS - Z
        If pSeq2(b, 0) <= 7 Then
        If pSeq2(b, 1) = 8 Then Sposta 2, 3, 3, 2 Else Sposta 1, 2, 2, 1
    ElseIf pSeq2(b, 0) <= 15 Then
        If pSeq2(b, 1) = 8 Then Sposta 1, 2, 2, 1 Else Sposta 1, 2, 2, 3, 3, 1
    ElseIf pSeq2(b, 0) <= 23 Then
        If pSeq2(b, 1) = 8 Then Sposta 1, 2, 2, 3, 3, 1 Else Sposta 1, 3, 2, 1, 3, 2
    ElseIf pSeq2(b, 0) <= 31 Then
        If pSeq2(b, 1) = 8 Then Sposta 1, 3, 2, 1, 3, 2 Else Sposta 2, 3, 3, 2
    End If

For X = 0 To 31: Table(X) = nTable(X): Next X
Return


F7: ' trovato un 7
'facciamo subito un mod
Md = Somma Mod 32
If Md <= 7 Then
pSeq = 24
ElseIf Md <= 15 Then
pSeq = 0
ElseIf Md <= 23 Then
pSeq = 16
ElseIf Md <= 31 Then
pSeq = 24
End If
pSeq2(nS, 0) = Md: pSeq2(nS, 1) = 7: nS = nS + 1
Md = Md + pSeq
Somma = Md
GoTo Cont

F8: ' Trovato un 8
'facciamo subito un mod
Md = Somma Mod 32
If Md <= 15 Then
pSeq = 16
ElseIf Md <= 23 Then
pSeq = 24
ElseIf Md <= 31 Then
pSeq = 8
End If
pSeq2(nS, 0) = Md: pSeq2(nS, 1) = 8: nS = nS + 1
Md = Md + pSeq
Somma = Md
GoTo Cont
End Function

Function MyMod(ByVal a, ByVal b)
    a = Fix(CDbl(a))
    b = Fix(CDbl(b))
    MyMod = a - Fix(a / b) * b
End Function

Private Sub Sposta(a, b, c, d As Byte, Optional e As Byte, Optional f As Byte)
'a to b | c to d | e to f
Dim m As Byte
For m = 0 To Blk(a, 1) - Blk(a, 0): nTable(Blk(b, 0) + m) = Table(Blk(a, 0) + m): Next m
For m = 0 To Blk(c, 1) - Blk(c, 0): nTable(Blk(d, 0) + m) = Table(Blk(c, 0) + m): Next m
If e > 0 And f > 0 Then
For m = 0 To Blk(e, 1) - Blk(e, 0): nTable(Blk(f, 0) + m) = Table(Blk(e, 0) + m): Next m
End If
End Sub
