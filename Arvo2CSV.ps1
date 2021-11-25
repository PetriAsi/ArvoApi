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
    $edellinen = $edelliset |Where-Object {$_.kyselyid -eq $get.kyselyid }
    if ($get.viimeisin_vastaus -ne $edellinen.viimeisin_vastaus) {
        $nimi = Join-Path  -path $saveTo -ChildPath ($_.nimi_fi -replace ':','-')
        $kid =  $_.kyselyid
        
        if ( -not (Test-Path  $nimi)) { new-item -Name $nimi -ItemType Directory}
        foreach ($key in $haettavat.Keys){
            $response = Invoke-WebRequest -headers $arvoh -Uri ($haettavat[$key] -replace 'KYSID',$kid) -WebSession $OpSession 
            [System.IO.StreamReader]::new($response.RawContentStream).ReadToEnd()| Out-File (Join-Path -Path $nimi -ChildPath ( $key + '-' + $kid + '.csv')) -Encoding utf8BOM
            #$response.headers | fl   
        }
    }
}

#Tallennetaan kyselyiden tiedot
$toget | export-clixml $tallennetut -Force


