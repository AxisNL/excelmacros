# excelmacros
Excel macro's to calculate Dutch taxes and more.

### How to use

The following was tested on Excel 2016 for Mac. 
In Excel, go to Tools > Macro > Visual Basic Editor. In the editor, go to Insert > Module. In the default Module1, paste your functions and exit the editor.

After this is done, you can access the functions from cells, for example by using the formula ``=Calculate_IB_box1_2017(A1)``

The rest of this document is in Dutch, to explain the Dutch terminology.

### Opmerking

Ik heb deze formules geschreven voor eigen gebruik. Ik ben nog lang niet in de buurt van de AOW leeftijd, en ik woon alleen, zonder kinderen. Ik werk deels in loondienst, deels als ZZP'er. Deze formules zijn wellicht anders bij andere omstandigheden! Gebruik op eigen risico.

### Belastingen
In Nederland betaal je in 3 boxen belasting:

- box 1: belastbaar inkomen uit werk en woning
- box 2: belastbaar inkomen uit aanmerkelijk belang
- box 3: belastbaar inkomen uit sparen en beleggen

In box 1 zit al je werk en inkomen, en in box 3 je vermogen. Box 2 laat ik buiten beschouwing.

### VBA formules

In de formules hieronder wordt het resultaat afgerond naar beneden op hele euro's.

## Box 1
####Inkomstenbelasting en premie volksverzekeringen 2017
		Function Calculate_IB_box1_2017(belastbaar_inkomen)     rest = belastbaar_inkomen     belasting = 0     'schijf 4     If rest > 67072 Then        schijf4_basis = rest - 67072        schijf4_belasting = 0.52 * schijf4_basis        belasting = belasting + schijf4_belasting        rest = rest - schijf4_basis     End If         'schijf 3     If rest > 33792 Then        schijf3_basis = rest - 33792        schijf3_belasting = 0.408 * schijf3_basis        belasting = belasting + schijf3_belasting        rest = rest - schijf3_basis     End If     'schijf 2     If rest > 19983 Then        schijf2_basis = rest - 19983        schijf2_belasting = 0.408 * schijf2_basis        belasting = belasting + schijf2_belasting        rest = rest - schijf2_basis     End If         'schijf 1     If rest > 0 Then        schijf1_basis = rest        schijf1_belasting = 0.3655 * schijf1_basis        belasting = belasting + schijf1_belasting     End If         Calculate_IB_box1_2017 = Int(belasting)	End Function#### Algemene heffingskorting

"De algemene heffingskorting is een korting op uw inkomstenbelasting en premie volksverzekeringen. U betaalt hierdoor minder belasting en premies. Iedereen heeft recht op de algemene heffingskorting. Maar of u volledig gebruik kunt maken van deze heffingskorting, hangt af van uw leeftijd en of u het hele jaar in Nederland hebt gewoond. De algemene heffingskorting is afhankelijk van de hoogte van uw inkomen. Dit betekent dat u minder algemene heffingskorting krijgt, als uw inkomen stijgt."
[bron](https://www.belastingdienst.nl/wps/wcm/connect/bldcontentnl/belastingdienst/prive/inkomstenbelasting/heffingskortingen_boxen_tarieven/heffingskortingen/algemene_heffingskorting/algemene_heffingskorting)

Deze functie wordt berekend over het 'belastbaar inkomen uit werk en woning', dus nadat de hypotheekrenteaftrek en de winstvrijstelling e.d. zijn werk heeft gedaan.
	Function 	Calculate_IB_algheffingskorting_2017(belastbaar_inkomen)     If belastbaar_inkomen > 67068 Then        Calculate_IB_algheffingskorting_2017 = 0        Exit Function     End If         If belastbaar_inkomen > 19982 Then        basis = belastbaar_inkomen - 19982        Calculate_IB_algheffingskorting_2017 = Int(2254 - (0.04787 * basis))        Exit Function     End If         Calculate_IB_algheffingskorting_2017 = 2254	End Function#### Arbeidskorting
"De arbeidskorting is de heffingskorting die u krijgt als u werkt. De arbeidskorting wordt berekend over het arbeidsinkomen." [bron](https://www.belastingdienst.nl/wps/wcm/connect/bldcontentnl/belastingdienst/prive/inkomstenbelasting/heffingskortingen_boxen_tarieven/heffingskortingen/arbeidskorting/arbeidskorting)

Let dus op dat dit over het arbeidsinkomen is. Dit het bruto loon uit loondienst, en de winst uit onderneming voor de aftrekposten zoals winstvrijstelling en ondernemersaftrek e.d.

	Function Calculate_IB_arbeidskorting_2017(arbeidsinkomen)     If arbeidsinkomen > 121972 Then        Calculate_IB_arbeidskorting_2017 = 0        Exit Function     End If         If arbeidsinkomen > 32444 Then        basis = arbeidsinkomen - 32444        Calculate_IB_arbeidskorting_2017 = Int(3223 - (0.036 * basis))        Exit Function     End If         If arbeidsinkomen > 20108 Then        Calculate_IB_arbeidskorting_2017 = 3223        Exit Function     End If         If arbeidsinkomen > 9309 Then        basis = arbeidsinkomen - 9309        Calculate_IB_arbeidskorting_2017 = Int(165 + (0.28317 * basis))        Exit Function     End If         Calculate_IB_arbeidskorting_2017 = Int(0.01772 * arbeidsinkomen)	End Function
#### Inkomensafhankelijke bijdrage ZVW

Als je werkt in loondienst, zal je werkgever over je loon al ZVW-premie betalen. Maar als je daarnaast (of slechts) in loondienst werkt, zul je over dat deel nog 'inkomensafhankelijke bijdrage ZVW' moeten betalen. Het totaal af te dragen bedrag is aan een maximum gebonden. [bron](https://www.belastingdienst.nl/wps/wcm/connect/bldcontentnl/belastingdienst/prive/werk_en_inkomen/zorgverzekeringswet/bijdrage_zorgverzekeringswet/)			Function Calculate_ZVW_2017(inkomen_loondienst, inkomen_freelance)	   totaal = inkomen_loondienst + inkomen_freelance    afdragen_totaal = 0.054 * totaal      If afdragen_totaal > 2899 Then        afdragen_totaal = 2899      End If      afdragen_loondienst = 0.054 * inkomen_loondienst      If afdragen_loondienst > 2899 Then        'De werkgever heeft alles al afgedragen!        Calculate_ZVW_2017 = 0        Exit Function      End If      afdragen_rest = afdragen_totaal - afdragen_loondienst      Calculate_ZVW_2017 = Int(afdragen_rest)	End Function## Box 3#### VermogensrendementsheffingHeb je vermogen in box 3, komt de vermogensrendementsheffing om de hoek kijken. Eerst wordt het rendement volgens een vaste tabel berekend, het vermogensrendement, en daarna wordt er 30% heffing berekend.	Function Calculate_VRH_2017(vermogen)     rest = vermogen     belasting = 0     heffingsvrij = 25000     rest = vermogen - heffingsvrij          If rest < 0 Then        Calculate_VRH_2017 = 0        Exit Function     End If         'schijf 3     If rest > 975001 Then        schijf3_basis = rest - 975001        schijf3_belasting = 0.0539 * schijf3_basis        belasting = belasting + schijf3_belasting        rest = rest - schijf3_basis     End If     'schijf 2     If rest > 75001 Then        schijf2_basis = rest - 75001        schijf2_belasting = 0.046 * schijf2_basis        belasting = belasting + schijf2_belasting        rest = rest - schijf2_basis     End If         'schijf 1     If rest > 0 Then        schijf1_basis = rest        schijf1_belasting = 0.02871 * schijf1_basis        belasting = belasting + schijf1_belasting     End If         Calculate_VRH_2017 = Int(belasting * 0.3)	End Function---

This content is free to use for anyone, but if you like it, be sure to let me know! Angelo Hongens - angelo@hongens.nl - Nov 2017