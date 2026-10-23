<%
'=====================================================================
'  PoliMusic :: SHA-1 en VBScript puro
'  Se usa para validar contrasenas almacenadas en formato hash
'  (la tabla TBL_USER contiene claves en texto plano y en SHA-1).
'  VBScript no permite operaciones bit a bit de 32 bits sin desbordar,
'  por eso se trabaja con dos mitades de 16 bits.
'=====================================================================
Function Hi16(x)
    Hi16 = Int(x / 65536)
End Function

Function Lo16(x)
    Lo16 = x - Int(x / 65536) * 65536
End Function

Function Xor32(a, b)
    Xor32 = (Hi16(a) Xor Hi16(b)) * 65536 + (Lo16(a) Xor Lo16(b))
End Function

Function And32(a, b)
    And32 = (Hi16(a) And Hi16(b)) * 65536 + (Lo16(a) And Lo16(b))
End Function

Function Or32(a, b)
    Or32 = (Hi16(a) Or Hi16(b)) * 65536 + (Lo16(a) Or Lo16(b))
End Function

Function Not32(a)
    Not32 = 4294967295 - a
End Function

Function Add32(a, b)
    Dim s
    s = a + b
    Add32 = s - Int(s / 4294967296) * 4294967296
End Function

Function Shl32(a, n)
    Dim m
    If n = 0 Then
        Shl32 = a
    Else
        m = 2 ^ (32 - n)
        Shl32 = (a - Int(a / m) * m) * (2 ^ n)
    End If
End Function

Function Shr32(a, n)
    If n = 0 Then
        Shr32 = a
    Else
        Shr32 = Int(a / (2 ^ n))
    End If
End Function

Function Rotl32(a, n)
    n = n Mod 32
    If n = 0 Then
        Rotl32 = a
    Else
        Rotl32 = Or32(Shl32(a, n), Shr32(a, 32 - n))
    End If
End Function

Function Hex8(x)
    Dim i, p, v, s
    s = ""
    For i = 7 To 0 Step -1
        p = 16 ^ i
        v = Int(x / p)
        x = x - v * p
        s = s & Mid("0123456789abcdef", v + 1, 1)
    Next
    Hex8 = s
End Function

Function Sha1Hex(txt)
    Dim bytes(), L, total, i, j, b, bitLen, blk, pos
    Dim w(), a, bb, c, d, e, f, k, tmp
    Dim h0, h1, h2, h3, h4

    ReDim w(79)

    L = Len(txt)

    ' longitud total del mensaje con relleno
    total = L + 1
    Do While ((total + 8) Mod 64) <> 0
        total = total + 1
    Loop
    total = total + 8

    ReDim bytes(0 To total - 1)

    For i = 1 To L
        b = Asc(Mid(txt, i, 1))
        If b < 0 Then b = b + 65536
        bytes(i - 1) = b And 255
    Next
    bytes(L) = 128
    For i = L + 1 To total - 9
        bytes(i) = 0
    Next

    ' longitud del mensaje en bits (64 bits big-endian, 4 altos en cero)
    bitLen = L * 8
    For i = 0 To 3
        bytes(total - 1 - i) = Int(bitLen / (256 ^ i)) Mod 256
    Next

    h0 = 1732584193
    h1 = 4023233417
    h2 = 2562383102
    h3 = 271733878
    h4 = 3285377520

    For blk = 0 To (total \ 64) - 1
        For j = 0 To 15
            pos = blk * 64 + j * 4
            w(j) = bytes(pos) * 16777216 + bytes(pos + 1) * 65536 + bytes(pos + 2) * 256 + bytes(pos + 3)
        Next
        For j = 16 To 79
            w(j) = Rotl32(Xor32(Xor32(Xor32(w(j - 3), w(j - 8)), w(j - 14)), w(j - 16)), 1)
        Next

        a = h0
        bb = h1
        c = h2
        d = h3
        e = h4

        For j = 0 To 79
            If j < 20 Then
                f = Or32(And32(bb, c), And32(Not32(bb), d))
                k = 1518500249
            ElseIf j < 40 Then
                f = Xor32(Xor32(bb, c), d)
                k = 1859775393
            ElseIf j < 60 Then
                f = Or32(Or32(And32(bb, c), And32(bb, d)), And32(c, d))
                k = 2400959708
            Else
                f = Xor32(Xor32(bb, c), d)
                k = 3395469782
            End If

            tmp = Add32(Add32(Add32(Add32(Rotl32(a, 5), f), e), k), w(j))
            e = d
            d = c
            c = Rotl32(bb, 30)
            bb = a
            a = tmp
        Next

        h0 = Add32(h0, a)
        h1 = Add32(h1, bb)
        h2 = Add32(h2, c)
        h3 = Add32(h3, d)
        h4 = Add32(h4, e)
    Next

    Sha1Hex = Hex8(h0) & Hex8(h1) & Hex8(h2) & Hex8(h3) & Hex8(h4)
End Function

' Compara una contrasena escrita por el usuario contra el valor guardado:
' acepta texto plano (como en los datos de ejemplo) o SHA-1.
Function PasswordMatches(plainText, storedValue)
    Dim s, h
    s = TrimStr(storedValue)
    h = TrimStr(plainText)
    PasswordMatches = False
    If Len(s) = 0 Then Exit Function
    If StrComp(s, h, vbTextCompare) = 0 Then
        PasswordMatches = True
        Exit Function
    End If
    If StrComp(s, Sha1Hex(h), vbTextCompare) = 0 Then
        PasswordMatches = True
    End If
End Function
%>
