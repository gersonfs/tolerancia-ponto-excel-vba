Function CalcularHoraExtra(ent1 As Date, sai1 As Date, ent2 As Date, sai2 As Date, pe1 As Date, ps1 As Date, pe2 As Date, ps2 As Date)
    Dim diferencaDiariaMinima As Date
    
    If (pe1 = TimeValue("00:00") And ps1 = TimeValue("00:00")) Then
        CalcularHoraExtra = FormatarDecimalParaTempo(TimeValue("00:00"))
        Exit Function
    End If
    
    
    'Ajusta a data da saída quando passa da meia noite, adiciona 24 horas
    sai2 = AjustarSaida(ent2, sai2)
    ps2 = AjustarSaida(pe2, ps2)
    
    diferencaDiariaMinima = TimeValue("00:10")
    diferencaEntrada1 = ent1 - pe1
    diferencaSaida1 = ps1 - sai1
    diferencaEntrada2 = ent2 - pe2
    diferencaSaida2 = ps2 - sai2
    
    Total = diferencaEntrada1 + diferencaSaida1 + diferencaEntrada2 + diferencaSaida2
    
    
    If (Abs(Total) > diferencaDiariaMinima) Then
    
        If (diferencaEntrada2 + diferencaSaida2 > 0) Then
            pontoComHoraNoturna = (diferencaEntrada2 + diferencaSaida2) / 52.5 * 60
            Total = diferencaEntrada1 + diferencaSaida1 + pontoComHoraNoturna
        End If
        
        CalcularHoraExtra = FormatarDecimalParaTempo(Total)
        Exit Function
    End If
    
    
    If (Abs(diferencaEntrada1 + diferencaSaida1) > TimeValue("00:05:01")) Then
        CalcularHoraExtra = FormatarDecimalParaTempo(Total)
        Exit Function
    End If
    
    If (Abs(diferencaEntrada2 + diferencaSaida2) > TimeValue("00:05:01")) Then
        If (diferencaEntrada2 + diferencaSaida2 > 0) Then
            pontoComHoraNoturna = (diferencaEntrada2 + diferencaSaida2) / 52.5 * 60
            Total = diferencaEntrada1 + diferencaSaida1 + pontoComHoraNoturna
        End If
        CalcularHoraExtra = FormatarDecimalParaTempo(Total)
        Exit Function
    End If
    
    CalcularHoraExtra = FormatarDecimalParaTempo(TimeValue("00:00"))

End Function


Function FormatarDecimalParaTempo(valor)
    FormatarDecimalParaTempo = CDbl(valor)
    segundos = Format(CDate(Abs(valor)), "ss")
    If (segundos > "00") Then
       If (valor < 0) Then
          FormatarDecimalParaTempo = FormatarDecimalParaTempo + TimeValue("00:00:" & segundos)
        Else
          FormatarDecimalParaTempo = FormatarDecimalParaTempo - TimeValue("00:00:" & segundos)
       End If
    End If
End Function

Function AjustarSaida(entrada As Date, saida As Date)
    If (saida < entrada) Then
        saida = saida + TimeValue("23:59:59") + TimeValue("00:00:01")
    End If
    AjustarSaida = saida
End Function

Function FormatarHora(hora As Double)
    FormatarHora = Format(hora, "hh:mm:ss")
    If (hora < 0) Then
      FormatarHora = "-" & FormatarHora
    End If
End Function

Sub TestesAutomatizados()
    Debug.Assert FormatarHora(CalcularHoraExtra("16:15", "21:30", "22:30", "01:30", "16:15", "21:30", "22:30", "01:30")) = "00:00:00"
    Debug.Assert FormatarHora(CalcularHoraExtra("16:15", "21:30", "22:30", "01:30", "16:15", "21:30", "22:30", "01:41")) = "00:12:00"
    Debug.Assert FormatarHora(CalcularHoraExtra("16:15", "21:30", "22:30", "01:30", "16:30", "21:30", "22:30", "01:30")) = "-00:15:00"
    Debug.Assert FormatarHora(CalcularHoraExtra("16:15", "21:30", "22:30", "01:30", "16:00", "21:30", "22:30", "01:30")) = "00:15:00"
    Debug.Assert FormatarHora(CalcularHoraExtra("16:15", "21:30", "22:30", "01:30", "16:15", "21:41", "22:30", "01:30")) = "00:11:00"
    Debug.Assert FormatarHora(CalcularHoraExtra("16:15", "21:30", "22:30", "01:30", "16:14", "21:41", "22:31", "01:31")) = "00:12:00"
    Debug.Assert FormatarHora(CalcularHoraExtra("16:15", "21:30", "22:30", "01:30", "16:10", "21:30", "22:30", "01:30")) = "00:00:00"
    
    '5 minutos tolerancia saida1
    Debug.Assert FormatarHora(CalcularHoraExtra("16:15", "21:30", "22:30", "01:30", "16:15", "21:35", "22:30", "01:30")) = "00:00:00"
    
    '5 minutos tolerancia entrada1
    Debug.Assert FormatarHora(CalcularHoraExtra("16:15", "21:30", "22:30", "01:30", "16:15", "21:30", "22:35", "01:30")) = "00:00:00"
    
    '5 minutos tolerancia entrada2
    Debug.Assert FormatarHora(CalcularHoraExtra("16:15", "21:30", "22:30", "01:30", "16:15", "21:30", "22:35", "01:30")) = "00:00:00"
    
    '5 minutos tolerancia saida2
    Debug.Assert FormatarHora(CalcularHoraExtra("16:15", "21:30", "22:30", "01:30", "16:15", "21:30", "22:30", "01:35")) = "00:00:00"
    
    '6 minutos cobrados entrada1
    Debug.Assert FormatarHora(CalcularHoraExtra("16:15", "21:30", "22:30", "01:30", "16:12", "21:33", "22:30", "01:30")) = "00:06:00"
    
    '6 minutos cobrados entrada2
    Debug.Assert FormatarHora(CalcularHoraExtra("16:15", "21:30", "22:30", "01:30", "16:15", "21:30", "22:27", "01:33")) = "00:06:00"
    
    '60 minutos cobrados entrada2
    Debug.Assert FormatarHora(CalcularHoraExtra("16:15", "21:30", "22:30", "01:30", "16:15", "21:30", "22:30", "02:30")) = "01:08:00"
    
End Sub


