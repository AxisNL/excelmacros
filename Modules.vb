Option Explicit

'In de formules hieronder wordt het resultaat afgerond naar beneden op hele euro's.

'Box 1
'' Inkomstenbelasting en premie volksverzekeringen 2017

Function Calculate_IB_box1_2017(belastbaar_inkomen)
  Dim rest, belasting
  rest = belastbaar_inkomen
  belasting = 0
  'schijf 4
  If rest > 67072 Then
    Dim schijf4_basis, schijf4_belasting
    schijf4_basis = rest - 67072
    schijf4_belasting = 0.52 * schijf4_basis
    belasting = belasting + schijf4_belasting
    rest = rest - schijf4_basis
  End If

   'schijf 3
   If rest > 33792 Then
      Dim schijf3_basis, schijf3_belasting
      schijf3_basis = rest - 33792
      schijf3_belasting = 0.408 * schijf3_basis
      belasting = belasting + schijf3_belasting
      rest = rest - schijf3_basis
   End If

   'schijf 2
   If rest > 19983 Then
      Dim schijf2_basis, schijf2_belasting
      schijf2_basis = rest - 19983
      schijf2_belasting = 0.408 * schijf2_basis
      belasting = belasting + schijf2_belasting
      rest = rest - schijf2_basis
   End If

   'schijf 1
   If rest > 0 Then
      Dim schijf1_basis, schijf1_belasting
      schijf1_basis = rest
      schijf1_belasting = 0.3655 * schijf1_basis
      belasting = belasting + schijf1_belasting
   End If

   Calculate_IB_box1_2017 = Int(belasting)
End Function

'Algemene heffingskorting
'"De algemene heffingskorting is een korting op uw inkomstenbelasting en premie volksverzekeringen. U betaalt
' hierdoor minder belasting en premies. Iedereen heeft recht op de algemene heffingskorting. Maar of u volledig
' gebruik kunt maken van deze heffingskorting, hangt af van uw leeftijd en of u het hele jaar in Nederland hebt
' gewoond. De algemene heffingskorting is afhankelijk van de hoogte van uw inkomen. Dit betekent dat u minder
' algemene heffingskorting krijgt, als uw inkomen stijgt."
' bron: https://www.belastingdienst.nl/wps/wcm/connect/bldcontentnl/belastingdienst/prive/inkomstenbelasting/heffingskortingen_boxen_tarieven/heffingskortingen/algemene_heffingskorting/algemene_heffingskorting

'Deze functie wordt berekend over het 'belastbaar inkomen uit werk en woning', dus nadat de hypotheekrenteaftrek
'' en de winstvrijstelling e.d. zijn werk heeft gedaan.

Function Calculate_IB_algheffingskorting_2017(belastbaar_inkomen)
 If belastbaar_inkomen > 67068 Then
    Calculate_IB_algheffingskorting_2017 = 0
    Exit Function
 End If

 If belastbaar_inkomen > 19982 Then
    Dim basis
    basis = belastbaar_inkomen - 19982
    Calculate_IB_algheffingskorting_2017 = Int(2254 - (0.04787 * basis))
    Exit Function
 End If

 If belastbaar_inkomen < 2254 Then
    Calculate_IB_algheffingskorting_2017 = belastbaar_inkomen
    Exit Function
End If

 Calculate_IB_algheffingskorting_2017 = 2254
End Function

'Arbeidskorting
''"De arbeidskorting is de heffingskorting die u krijgt als u werkt. De arbeidskorting wordt berekend over het
' arbeidsinkomen."
' bron: https://www.belastingdienst.nl/wps/wcm/connect/bldcontentnl/belastingdienst/prive/inkomstenbelasting/heffingskortingen_boxen_tarieven/heffingskortingen/arbeidskorting/arbeidskorting
'
' Let dus op dat dit over het arbeidsinkomen is. Dit het bruto loon uit loondienst, en de winst uit onderneming
' voor de aftrekposten zoals winstvrijstelling en ondernemersaftrek e.d.


Function Calculate_IB_arbeidskorting_2017(arbeidsinkomen)
  Dim basis
   If arbeidsinkomen > 121972 Then
      Calculate_IB_arbeidskorting_2017 = 0
      Exit Function
   End If

   If arbeidsinkomen > 32444 Then
      basis = arbeidsinkomen - 32444
      Calculate_IB_arbeidskorting_2017 = Int(3223 - (0.036 * basis))
      Exit Function
   End If

   If arbeidsinkomen > 20108 Then
      Calculate_IB_arbeidskorting_2017 = 3223
      Exit Function
   End If

   If arbeidsinkomen > 9309 Then
      basis = arbeidsinkomen - 9309
      Calculate_IB_arbeidskorting_2017 = Int(165 + (0.28317 * basis))
      Exit Function
   End If

   Calculate_IB_arbeidskorting_2017 = Int(0.01772 * arbeidsinkomen)
End Function

' Inkomensafhankelijke bijdrage ZVW
' Als je werkt in loondienst, zal je werkgever over je loon al ZVW-premie betalen. Maar als je daarnaast
' (of slechts) in loondienst werkt, zul je over dat deel nog 'inkomensafhankelijke bijdrage ZVW' moeten
' betalen. Het totaal af te dragen bedrag is aan een maximum gebonden.
' bron: https://www.belastingdienst.nl/wps/wcm/connect/bldcontentnl/belastingdienst/prive/werk_en_inkomen/zorgverzekeringswet/bijdrage_zorgverzekeringswet/

Function Calculate_ZVW_2017(inkomen_loondienst, inkomen_freelance)
  Dim totaal, afdragen_totaal, afdragen_loondienst, afdragen_rest
  totaal = inkomen_loondienst + inkomen_freelance
  afdragen_totaal = 0.054 * totaal
  If afdragen_totaal > 2899 Then
    afdragen_totaal = 2899
  End If
  afdragen_loondienst = 0.054 * inkomen_loondienst
  If afdragen_loondienst > 2899 Then
    'De werkgever heeft alles al afgedragen!
    Calculate_ZVW_2017 = 0
    Exit Function
  End If
  afdragen_rest = afdragen_totaal - afdragen_loondienst
  Calculate_ZVW_2017 = Int(afdragen_rest)
End Function

' Box 3
' Vermogensrendementsheffing
'
' Heb je vermogen in box 3, komt de vermogensrendementsheffing om de hoek kijken. Eerst wordt het rendement
' volgens een vaste tabel berekend, het vermogensrendement, en daarna wordt er 30% heffing berekend.

Function Calculate_VRH_2017(vermogen)
  Dim rest, belasting, heffingsvrij
  rest = vermogen
  belasting = 0
  heffingsvrij = 25000
  rest = vermogen - heffingsvrij

  If rest < 0 Then
    Calculate_VRH_2017 = 0
    Exit Function
  End If

  'schijf 3
  If rest > 975001 Then
    Dim schijf3_basis, schijf3_belasting
    schijf3_basis = rest - 975001
    schijf3_belasting = 0.0539 * schijf3_basis
    belasting = belasting + schijf3_belasting
    rest = rest - schijf3_basis
  End If

  'schijf 2
  If rest > 75001 Then
    Dim schijf2_basis, schijf2_belasting
    schijf2_basis = rest - 75001
    schijf2_belasting = 0.046 * schijf2_basis
    belasting = belasting + schijf2_belasting
    rest = rest - schijf2_basis
  End If

  'schijf 1
  If rest > 0 Then
    Dim schijf1_basis, schijf1_belasting
    schijf1_basis = rest
    schijf1_belasting = 0.02871 * schijf1_basis
    belasting = belasting + schijf1_belasting
  End If

  Calculate_VRH_2017 = Int(belasting * 0.3)
End Function

' REST
' KOR
' De kleine ondernemersregeling. Let op, tel het KOR-bedrag wel op bij de inkomsten box 1!

Function Calculate_KOR_2017(btw_te_betalen)
  If btw_te_betalen >= 1883 Then
      Calculate_KOR_2017 = 0
      Exit Function
  End If

  If btw_te_betalen > 1345 Then
      Calculate_KOR_2017 = 2.5 * btw_te_betalen
      Exit Function
  End If

  If btw_te_betalen <= 0 Then
      Calculate_KOR_2017 = 0
      Exit Function
  End If
  Calculate_KOR_2017 = btw_te_betalen
End Function


' Huurtoeslag, 1 persoonshuishouden, onder aow leeftijd, geen aangepaste woning, etc.

Function Huurtoeslag_Simple_EP_2017(vermogen, kale_huur, servicekosten_energie, servicekosten_huismeester, servicekosten_schoonmaak, servicekosten_dienstruimten, rekeninkomen)
    '1persoons huishouden
    'geen woonwagen, geen bijzonderheden

    If vermogen > 25000 Then
        Huurtoeslag_Simple_EP_2017 = 0
        Exit Function
    End If

    'Rekenhuur
    If servicekosten_energie > 12 Then
    servicekosten_energie = 12
    End If
    If servicekosten_huismeester > 12 Then
    servicekosten_huismeester = 12
    End If
    If servicekosten_schoonmaak > 12 Then
    servicekosten_schoonmaak = 12
    End If
    If servicekosten_dienstruimten > 12 Then
    servicekosten_dienstruimten = 12
    End If
    Dim Rekenhuur
    Rekenhuur = kale_huur + servicekosten_energie + servicekosten_huismeester + servicekosten_schoonmaak + servicekosten_dienstruimten

    If Rekenhuur > 710.68 Then
        Huurtoeslag_Simple_EP_2017 = 0
        Exit Function
    End If

    If rekeninkomen < 0 Then
        rekeninkomen = 0
    End If

    If rekeninkomen > 22200 Then
        Huurtoeslag_Simple_EP_2017 = 0
        Exit Function
    End If

    'Stap 5: bereken de basishuur
    'De basishuur is het deel van de rekenhuur dat uw klant zelf moet betalen. Bij een inkomen tot
    'en met de minimuminkomensgrens geldt een minimumbasishuur. Bij een hoger inkomen wordt
    'de basishuur berekend met een formule:
    Dim Basishuur
    Dim Factor_a
    Dim Factor_b
    Dim taakstellingsbedrag

    If rekeninkomen < 15675 Then
        Basishuur = 223.42
    Else
        Factor_a = 7.02729E-07
        Factor_b = 0.002157297539
        taakstellingsbedrag = 16.94
        Basishuur = (Factor_a * (rekeninkomen ^ 2)) + (Factor_b * (rekeninkomen ^ 2)) + taakstellingsbedrag
    End If

    'Stap 6: bepaal de kwaliteitskortingsgrens
    'De kwaliteitskortingsgrens is vastgesteld op € 414,02. Voor het deel van de rekenhuur dat
    'boven de basishuur en onder de kwaliteitskortingsgrens ligt, krijgt uw klant 100% vergoeding.
    Dim kwaliteitskortingsgrens
    kwaliteitskortingsgrens = 414.02

    'Stap 7: bepaal de aftoppingsgrens
    'De a oppingsgrens wordt bepaald door de omvang van het huishouden:
    '• voor een huishouden van 1 of twee personen: € 592,55
    Dim aftoppingsgrens
    aftoppingsgrens = 592.55


    'Rekenen

    'Onderdeel A: het deel van de rekenhuur dat voor 100% vergoed wordt
    'Bepaal het laagste bedrag van rekenhuur en kwaliteitskortingsgrens
    Dim result
    Dim toeslag_deel_A
    result = WorksheetFunction.Min(Rekenhuur, kwaliteitskortingsgrens)
    toeslag_deel_A = result - Basishuur
    If toeslag_deel_A < 0 Then
        toeslag_deel_A = 0
    End If

    'Onderdeel B: het deel van de rekenhuur dat voor 65% vergoed wordt
    'Alleen invullen als de rekenhuur meer is dan € 414,02.
    If Rekenhuur > 414.02 Then
        Dim var1, var2
        'Bepaal het laagste bedrag van rekenhuur en a oppingsgrens
        var1 = WorksheetFunction.Min(Rekenhuur, aftoppingsgrens)
        'Bepaal het hoogste bedrag van basishuur en kwaliteitskortingsgrens
        var2 = WorksheetFunction.Max(Basishuur, kwaliteitskortingsgrens)
        Dim toeslag_deel_B
        toeslag_deel_B = 0.65 * (var1 - var2)
        If toeslag_deel_B < 0 Then
            toeslag_deel_B = 0
        End If

    Else
        toeslag_deel_B = 0
    End If

    Huurtoeslag_Simple_EP_2017 = toeslag_deel_A + toeslag_deel_B

End Function


Function Zorgtoeslag_Simple_EP_2017(vermogen, rekeninkomen)
    '1persoons huishouden, geen buitenland van toepassing

    If vermogen > 107752 Then
        Zorgtoeslag_Simple_EP_2017 = 0
        Exit Function
    End If

    'Stap 1: bepaal de standaardpremie
    'De standaardpremie is voor 2017 vastgesteld op € 1.530. Bij een aanvrager met een toeslagpartner wordt tweemaal de standaardpremie genomen (€ 3.060).
    Dim standaardpremie
    standaardpremie = 1530

    'Stap 2: bereken het gezamenlijke toetsingsinkomen
    Dim toetsingsinkomen
    toetsingsinkomen = rekeninkomen

    'Uw klant hee  geen recht op zorgtoeslag als het toetsingsinkomen hoger is dan:
    '• € 27.857 (aanvrager zonder toeslagpartner)
    '• € 35.116 (aanvrager met toeslagpartner)

    If toetsingsinkomen > 27857 Then
        Zorgtoeslag_Simple_EP_2017 = 0
        Exit Function
    End If

    'Stap 3: bereken de normpremie
    'De normpremie wordt berekend met het drempelinkomen en het gezamenlijke toetsingsinkomen.
    'Het drempelinkomen is voor 2017 vastgesteld op € 20.109.
    Dim drempelinkomen
    drempelinkomen = 20109

    'Voor een aanvrager zonder toeslagpartner:
    'Normpremie = 2,305% x drempelinkomen + 13,46% (toetsingsinkomen - drempelinkomen)
    'Leidt het tweede deel van de formule tot een negatief bedrag? Reken dan met 0.

    Dim normpremie, tmp

    tmp = 0.1346 * (toetsingsinkomen - drempelinkomen)
    If tmp < 0 Then
        tmp = 0
    End If
    normpremie = (0.02305 * drempelinkomen) + tmp

    'De maximale zorgtoeslag wordt uitgekeerd bij een toetsingsinkomen op of onder het
    'drempelinkomen. In 2017 is dat voor een aanvrager zonder toeslagpartner € 1.066
    'en voor een aanvrager met toeslagpartner € 2.043.
    '??

    'Stap 4: bereken de zorgtoeslag voor uw klant die in Nederland woont
    'Woont uw klant in het buitenland? Ga dan verder met stap 5. Voor een aanvrager zonder toeslagpartner:
    Dim zorgtoeslag
    zorgtoeslag = standaardpremie - normpremie


    Zorgtoeslag_Simple_EP_2017 = zorgtoeslag

End Function
