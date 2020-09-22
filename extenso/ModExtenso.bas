Attribute VB_Name = "Module1"
Public Function Extenso(ByVal Valor As _
       Double, ByVal MoedaPlural As _
       String, ByVal MoedaSingular As _
       String) As String
  Dim StrValor As String, Negativo As Boolean
  Dim Buf As String, Parcial As Integer
  Dim Posicao As Integer, Unidades
  Dim Dezenas, Centenas, PotenciasSingular
  Dim PotenciasPlural

  Negativo = (Valor < 0)
  Valor = Abs(CDec(Valor))
  If Valor Then
    Unidades = Array(vbNullString, "Um", "Dois", _
               "Três", "Quatro", "Cinco", _
               "Seis", "Sete", "Oito", "Nove", _
               "Dez", "Onze", "Doze", "Treze", _
               "Catorze", "Quinze", "Dezasseis", _
               "Dezassete", "Dezoito", "Dezanove")
    Dezenas = Array(vbNullString, vbNullString, _
              "Vinte", "Trinta", "Quarenta", _
              "Cinquenta", "Sessenta", "Setenta", _
              "Oitenta", "Noventa")
    Centenas = Array(vbNullString, "Cento", _
               "Duzentos", "Trezentos", _
               "Quatrocentos", "Quinhentos", _
               "Seiscentos", "Setecentos", _
               "Oitocentos", "Novecentos")
    PotenciasSingular = Array(vbNullString, " Mil", _
                        " Milhão", " Bilião", _
                        " Trilião", " Quatrilião")
    PotenciasPlural = Array(vbNullString, " Mil", _
                      " Milhões", " Biliões", _
                      " Triliões", " Quatriliões")

    StrValor = Left(Format(Valor, String(18, "0") & _
               ".000"), 18)
    For Posicao = 1 To 18 Step 3
      Parcial = Val(Mid(StrValor, Posicao, 3))
     
     
        If Parcial Then
        
If Posicao = 13 And Val(Mid(StrValor, Posicao, 3)) = 1 Then
         Buf = "" & PotenciasSingular((18 - _
                Posicao) \ 3)
                
                
        ElseIf Parcial = 1 Then
          Buf = "Um" & PotenciasSingular((18 - _
                Posicao) \ 3)
        ElseIf Parcial = 100 Then
          Buf = "Cem" & PotenciasSingular((18 - _
                Posicao) \ 3)
        Else
          Buf = Centenas(Parcial \ 100)
          Parcial = Parcial Mod 100
          If Parcial <> 0 And Buf <> vbNullString Then
            Buf = Buf & " e "
          End If
          If Parcial < 20 Then
            Buf = Buf & Unidades(Parcial)
          Else
            Buf = Buf & Dezenas(Parcial \ 10)
            Parcial = Parcial Mod 10
            If Parcial <> 0 And Buf <> vbNullString Then
              Buf = Buf & " e "
            End If
            Buf = Buf & Unidades(Parcial)
          End If
          Buf = Buf & PotenciasPlural((18 - Posicao) \ 3)
        End If
        If Buf <> vbNullString Then
          If Extenso <> vbNullString Then
            Parcial = Val(Mid(StrValor, Posicao, 3))
            If Posicao = 16 And (Parcial < 100 Or _
                (Parcial Mod 100) = 0) Then
              Extenso = Extenso & " e "
            Else
              Extenso = Extenso & ", "
            End If
          End If
          Extenso = Extenso & Buf
        End If
      End If
    Next
    If Extenso <> vbNullString Then
      If Negativo Then
        Extenso = "Menos " & Extenso
      End If
      If Int(Valor) = 1 Then
        Extenso = Extenso & " " & MoedaSingular
      Else
        Extenso = Extenso & " " & MoedaPlural
      End If
    End If
    Parcial = Int((Valor - Int(Valor)) * _
              100 + 0.1)
    If Parcial Then
      Buf = Extenso(Parcial, "Cêntimos", _
            "Cêntimo")
      If Extenso <> vbNullString Then
        Extenso = Extenso & " e "
      End If
      Extenso = Extenso & Buf
    End If
  End If
End Function

