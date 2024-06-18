[CmdletBinding()]
param(
    [bool]$kaikkivastaustunnukset=$false
)

import-module PowerHTML
# tallenna arvojärjestelmssä käytetty tunnus seuraavasti tiedostoon
# get-credential | export-clixml opintopolku.xml
if (-not (Test-path -Path "opintopolku.xml" -PathType leaf)) {
    $virhe = @"
tallenna arvojärjestelmssä käytetty tunnus seuraavasti tiedostoon
"get-credential | export-clixml opintopolku.xml"
ja yritä scriptin ajoa uudelleen
"@
    throw $virhe           
}

$cred = import-clixml "opintopolku.xml"

#kansio mihin kyselyt tallennetaan
$saveTo = "Arvosta"
$tallennetut = Join-Path -Path $saveTo -ChildPath "tallennetut.xml"


$loginUrl = "https://virkailija.opintopolku.fi/cas/login?service=https%3A%2F%2Fvirkailija.opintopolku.fi%2Fvirkailijan-tyopoyta%2Fauthenticate"
$arvourl = "https://arvo.csc.fi"

#varmistetaan tallennuskansio
if (-not (Test-Path $saveTo -PathType Container) ) {
    try {
        $parent = Split-path -path $saveTo -Parent
        $leaf = Split-path -path $saveTo -Leaf
        
        if($parent -eq ""){
            $parent = "."
        }
        
        New-Item -Path $saveTo -Name $leaf  -ItemType "directory"
        }
    catch{
        throw "Tallennushakemiston $saveTo luonti ei onnistunut"
    }
}

try {
    $logininfo = Invoke-WebRequest -Uri $loginUrl -SessionVariable OpSession 

    $execution = ($logininfo.Content | ConvertFrom-HTML).SelectSingleNode('//*[@id="fm1"]/input[1]' ).GetAttributeValue('Value','')

    $loginParms=@{
        'username' = $cred.GetNetworkCredential().username
        'password' = $cred.GetNetworkCredential().password
        'execution' = $execution
        '_eventId' = 'submit'
        'geolocation' = ''       
    }

    $headers = @{
        'referer' = $loginUrl
        'origin'  = 'https://virkailija.opintopolku.fi'
        'accept'  = 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9'
    }

    $arvoh = @{
        'referer' =  $arvourl
        'accept'  = 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9'
    }


    $loginResult = Invoke-WebRequest -Method POST -Headers $headers -Body $loginParms -Uri $loginUrl -WebSession $OpSession 
    $arvo = Invoke-WebRequest -Uri $arvourl -Headers $arvoh -WebSession $OpSession 

    $arvoh['X-XSRF-TOKEN'] = $Opsession.Cookies.GetCookies($arvourl)['XSRF-TOKEN'].Value
}
catch {
    $ErrorMessage = $_.Exception.Message
    throw "Yhdistäminen arvoon ei onnistunut: $ErrorMessage"    
}

$haettavat = @{
    'kysely' = "$arvourl/api/csv/kysely/KYSID?lang=fi"
    'vastauksittain' = "$arvourl/api/csv/kysely/vastauksittain/KYSID?lang=fi"
    'kohteet' = "$arvourl/api/csv/kysely/kohteet/KYSID?lang=fi"
    'vastanneet' = "$arvourl/api/csv/kysely/vastaajat/KYSID?lang=fi"
    'vastaustunnukset' = "$arvourl/api/csv/vastaajat/KYSID?lang=fi"
}



$arvoh['accept'] = 	'application/json, text/plain, */*'
$arvoh['angular-ajax-request'] = 'true'
$arvoh['Origin'] = $arvourl
$arvoh['Sec-Fetch-Site'] =	'same-origin'
$arvoh['Sec-Fetch-Mode'] =	'cors'
$arvoh['Sec-Fetch-Dest'] =	'empty'

try {
    $kyselyt = Invoke-webrequest -uri 'https://arvo.csc.fi/api/kysely' -WebSession $OpSession -Headers $arvoh -Method Get 

    #ladataan delliset kyselyt, vain muuttuneiden tietojen hakemiseen
    if ((test-path $tallennetut -PathType Leaf)) {
        $edelliset = import-clixml $tallennetut        
    } else {
        $edelliset = @{}
    }
    
} catch {
    $ErrorMessage = $_.Exception.Message
    throw "Kyselyiden hakeminen epäonnistui: $ErrorMessage"  
}

$toget = $kyselyt.Content | ConvertFrom-Json

$toget | ForEach-Object {
    $get = $_
    $kid =  $get.kyselyid
    $edellinen = $edelliset |Where-Object {$_.kyselyid -eq $get.kyselyid }
    $nimi = Join-Path  -path $saveTo -ChildPath ($_.nimi.fi -replace ':','-')
    Write-Verbose "Tallennetaan kid: $($kid) polkuun: $($nimi)"        
    if ($null -ne $get.viimeisin_vastaus -and $get.viimeisin_vastaus -ne $edellinen.viimeisin_vastaus) {
        if ( -not (Test-Path  $nimi)) { new-item -Name $nimi -ItemType Directory}
    
        foreach ($key in ('kysely','kohteet','vastauksittain','vastanneet')){
            $response = Invoke-WebRequest -headers $arvoh -Uri ($haettavat[$key] -replace 'KYSID',$kid) -WebSession $OpSession 
            $tallennnustiedosto = Join-Path -Path $nimi -ChildPath ( $key + '-' + $kid + '.csv')
            [System.IO.StreamReader]::new($response.RawContentStream).ReadToEnd()| Out-File $tallennnustiedosto -Encoding utf8BOM
            if ( (Get-Content $tallennnustiedosto | measure-object).count -eq 1) {
                Remove-Item $tallennnustiedosto -Force
            }
            #$response.headers | fl   
        }
    }

    #vastustunnukset
    if ($null -ne $get.kyselykerrat) {
      #avoinnaolevat tai tulevat
      foreach ( $kk in ($get.kyselykerrat | where-object { ($_.kaytettavissa -eq $true) -or ( $_.kaytettavissa -eq $false -and [datetime]$_.voimassa_alkupvm -gt (get-date)) -or $kaikkivastaustunnukset})){
            if ( -not (Test-Path  $nimi)) { new-item -Name $nimi -ItemType Directory}
    
            $kkid = $kk.kyselykertaid
            Write-verbose "Vastaajatunnukset kysely $($kid) - kerta $($kkid))"
            #hae vain jos on tullut uusia vastaajatunnuksia
            
            if (($kk.vastaajatunnuksia -ne ($edellinen.kyselykerrat | where-object {$_.kyselykertaid -eq $kkid}).vastaajatunnuksia ) -or $kaikkivastaustunnukset) {
                    $lähde = ($haettavat['vastaustunnukset'] -replace 'KYSID',$kid) 
                    $kohde = ( 'vastaustunnukset-' + $kid + '-' + $kkid +'.csv')
                    $response = Invoke-WebRequest -headers $arvoh -Uri $lähde -WebSession $OpSession 
                    $tallennnustiedosto = Join-Path -Path $nimi -ChildPath $kohde
                    [System.IO.StreamReader]::new($response.RawContentStream).ReadToEnd()| Out-File $tallennnustiedosto -Encoding utf8BOM
                    if ( (Get-Content $tallennnustiedosto | measure-object).count -eq 1) {
                        Write-Error "Tiedosto tyhjä $($tallennnustiedosto)"
                        Remove-Item $tallennnustiedosto -Force
                    } else {
                        #kaikki tunnukset palautuvat nyt yhdellä kyselyllä, voimme poistua
                        Write-Verbose "Tallennettu tiedostoon $($tallennnustiedosto)"
                        break
                    }
                        
            }
    
        }
    }
}

#Tallennetaan kyselyiden tiedot
$toget | export-clixml $tallennetut -Force