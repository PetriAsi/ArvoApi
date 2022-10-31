# Arvo2CSV


Pieni powershell scripti arvo-kyselyiden tulosten hakemiseen csv muodossa.

Scripti autentikoi arvo-palveluun ja käyttää samoja api-päätepisteitä
kuin mikä ovat käytössä kun palvelua käyttää verkkoselaimen kanssa.

## Vaatimukset
Powershell 7.0 ja PowerHTML moduuli asennettuna

## valmistelu
### Yhteystunnus
tallenna arvojärjestelmässä käytetty tunnus seuraavasti tiedostoon
samaan kansioon missä scriptikin on. 

```powershell
get-credential | export-clixml opintopolku.xml
```
Kyseisellä tunnuksella pitää olla
vastuukäyttäjärooli arvoon , jolloin csv raportit näkyvät arvossa ja
ovat ladattavissa.

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

## Azure strorage ja Power bi
Milestäni helpoin tapa saada cvs tiedot Power Bi:n saataville on ajaa tämä scripti azuressa virtualikoneessa. Virtuaalikonelle kun
määrittää järjestelmän liitetyn identiteetin (system assigned identity) voi sitten Azure Data Lakessa antaa virtuallikoneelle suoran oikeuden
ladata tiedostoja haluamaansa tiedostosäilöön.

Tämän jälkeen ei tavitse kuin kopioida azcopy.exe samaan kansioon kuin missä scriptiä ajetettaan ja lisätä scriptin perään seuraavat rivit
tiedotojen kopioimiseksi azure data lakeen.
```powershell
.\azcopy.exe login --identity
.\azcopy.exe sync .\Arvosta 'https://[OMASTORAGENIMI].blob.core.windows.net/[SÄILÖNNIMI]'
```
