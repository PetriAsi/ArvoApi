# Arvo2CSV
Pieni powershell scripti arvo-kyselyiden tulosten hakemiseen csv muodossa.

## Vaatimukset
Powershell 7.0 ja PowerHTML moduuli asennettuna

## valmistelu
### Yhteystunnus
tallenna arvojärjestelmssä käytetty tunnus seuraavasti tiedostoon
samaan kansioon missä scriptikin on
```powershell
get-credential | export-clixml opintopolku.xml
```
### tallenushakemisto
Scripti tallentaa kyselyiden tulokset Arvosta kansioon, jos haluat
muuttaa tallenuspaikkaa muokkaa scriptin alussa olevaa riviä:
```powershell
$saveTo = "Arvosta"
```

## Scriptin ajaminen
Suorita scripti powershell konsolista tai ajastetusti miten haluatkin.
Tuloksena pitäisi olla kansiorakenne joka jokainen sisältää kyselyn 
csv tiedostot.

Kyselyistä haetaan seuraavilla ajokerroilla vain ne joihin on tullut
uusia vastauksia.
```powershell
Get-ChildItem .\Arvosta\

    Directory: C:\...\Arvosta

Mode                 LastWriteTime         Length Name
----                 -------------         ------ ----
d----          12.11.2021     9.56                Amispalaute, ammatillisen tutkinnon osan tai osia suorittaneet
d----          12.11.2021     9.56                Amispalaute, ammatillisen tutkinnon suorittaneet
d----          12.11.2021     9.55                Amispalaute, ammatillisen tutkintokoulutuksen aloittaneet
d----          12.11.2021     9.53                Arvosta
d----          12.11.2021     9.56                Opiskelijapalaute - ammatillisen tutkinnon osan tai osia suorittaneet
                                                   2020-2021
d----          12.11.2021     9.56                Opiskelijapalaute - ammatillisen tutkinnon osan tai osia suorittaneet
                                                   2021-2022
d----          12.11.2021     9.56                Opiskelijapalaute - ammatillisen tutkinnon suorittaneet 2020-2021
d----          12.11.2021     9.56                Opiskelijapalaute - ammatillisen tutkinnon suorittaneet 2021-2022
d----          12.11.2021     9.56                Opiskelijapalaute - ammatillisen tutkintokoulutuksen aloittaneet 2020
                                                  -2021
d----          12.11.2021     9.56                Opiskelijapalaute - ammatillisen tutkintokoulutuksen aloittaneet 2021
                                                  -2022
d----          12.11.2021     9.56                TYÖPAIKKAOHJAAJAKYSELY
-a---          12.11.2021     9.57         262929 tallennetut.xml
```