#Requires -modules TervisWCSSybase,TervisPasswordstatePowerShell,TervisMicrosoft.PowerShell.Utility

function Set-TervisShipEngineEnvironment {
    param (
        [Parameter(Mandatory)]
        [ValidateSet("Delta", "Production")]
        $Name
    )

    $GUID = switch ($Name) {
        "Delta" { "46a1e774-6777-44f4-9d1e-4bdc2ba347e0"; break }
        "Production" { "152d3024-c204-455d-8b76-1b279b44e3e3"; break }
        Default {}
    }

    $Script:APICredential = Get-TervisPasswordstatePassword -Guid $GUID -AsCredential
}

function Invoke-TervisShipEngineShipWarrantyOrder {    
    param (
        [Parameter(Mandatory)]$Name,
        $Phone,
        $CompanyName,
        [Parameter(Mandatory)]$AddressLine1,
        $AddressLine2,
        $AddressLine3,
        [Parameter(Mandatory)]$CityLocality,
        [Parameter(Mandatory)]$StateProvince,
        [Parameter(Mandatory)]$PostalCode,
        $CountryCode = "US",
        [Parameter(Mandatory)]$WeightInLB,
        $OriginLocation = "Osprey_Returns"
    )
    try {
        $WarehouseId = Get-TervisShipEngineWarehouseId -Location $OriginLocation
    
        $ShipTo = New-TervisShipEngineAddress `
            -Name $Name `
            -Phone $Phone `
            -AddressLine1 $AddressLine1 `
            -AddressLine2 $AddressLine2 `
            -AddressLine3 $AddressLine3 `
            -CityLocality $CityLocality `
            -StateProvince $StateProvince `
            -PostalCode $PostalCode `
            -CountryCode $CountryCode
    
        $ServiceCode = Get-TervisShipEngineWarrantyOrderService -WeightInLB $WeightInLB
    
        $Carrier = (Get-TervisShipEngineCarriers).Content.carriers.services | 
        Where-Object service_code -eq $ServiceCode
    
        [Array]$Packages = New-TervisShipEnginePackage -WeightInLB $WeightInLB
    
        New-TervisShipEngineLabel `
            -CarrierId $Carrier.carrier_id `
            -ServiceCode $Carrier.service_code `
            -ShipTo $ShipTo `
            -WarehouseId $WarehouseId `
            -Packages $Packages
    } catch {
        "$(Get-Date -Format o)`nInvoke-TervisShipEngineShipWarrantyOrder Error`n$($_.InvocationInfo.PositionMessage)`n$($_.ScriptStackTrace)`n" | Out-File -FilePath "C:\Log\TervisWarrantyFormInternal\log.txt" -Append

    }

}

function Invoke-TervisShipEngineAPI {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]$Endpoint,
        $APIVersion = "v1",
        $Method = "GET",
        $Resource,
        $Subresource,
        $Body
    )

    $ShipEngineAPIHost = "api.shipengine.com"

    $Headers = @{
        Host      = $ShipEngineAPIHost
        "API-Key" = $Script:APICredential.GetNetworkCredential().Password
    }

    $Uri = (
        $ShipEngineAPIHost, $APIVersion, $Endpoint, $Resource, $Subresource |
        Where-Object { $_ -notlike "" }
    ) -join "/"

    Write-Verbose -Message (@{
            Uri     = $Uri
            Headers = $Headers
            Body    = $Body
        } | ConvertTo-Json)

    "$(Get-Date -Format o)`n
    Invoke-TervisShipEngineAPI Error`n
    Uri:`n
    $Uri`n
    Headers`n
    $($Headers | ConvertTo-Json)`n
    Body`n
    $($Body | ConvertTo-Json)`n
    " | 
    Out-File -FilePath "C:\Log\TervisWarrantyFormInternal\log.txt" -Append
    
    $Response = try {
        Invoke-WebRequest `
            -Uri $Uri `
            -Headers $Headers `
            -Method $Method `
            -ContentType "application/json" `
            -Body $Body
    }
    catch {
        # To get the actual content of the error. 
        # https://stackoverflow.com/questions/25057721/how-do-you-get-the-response-from-a-404-page-requested-from-powershell
        $ContentStream = $_.Exception.Response.GetResponseStream()
        $StreamReader = New-Object System.IO.StreamReader($ContentStream)
        $Content = $StreamReader.ReadToEnd()

        [PSCustomObject]@{
            StatusCode = $_.Exception.Response.StatusCode.value__
            Content    = $Content
        }
    }

    if ($Response.StatusCode -ne 200) {
        
        throw "$($Responses.StatusCode) $()"
    }

    return [PSCustomObject]@{
        Status  = $Response.StatusCode
        Content = $Response.Content | ConvertFrom-Json
    }
}

function Get-TervisShipEngineCarriers {
    Invoke-TervisShipEngineAPI -Endpoint "carriers"
}

function New-TervisShipEngineLabel {
    param (
        [Parameter(Mandatory)]$CarrierId,
        [Parameter(Mandatory)]$ServiceCode,
        [Parameter(Mandatory)]$ShipTo,
        [Parameter(Mandatory, ParameterSetName = "ShipFrom")]$ShipFrom,
        [Parameter(Mandatory, ParameterSetName = "WarehouseId")]$WarehouseId,
        [Array]$Packages
    )

    $Payload = @{
        shipment = @{
            carrier_id   = $CarrierId
            service_code = $ServiceCode
            ship_to      = $ShipTo
            ship_from    = $ShipFrom
            warehouse_id = $WarehouseId
            packages     = $Packages
        } | Remove-HashtableKeysWithEmptyOrNullValues
    } | ConvertTo-Json -Depth 10 -Compress

    Invoke-TervisShipEngineAPI -Endpoint "labels" -Method "POST" -Body $Payload
}

function Get-TervisShipEngineLabel {
    param (
        [Parameter(Mandatory)]$LabelId
    )

    $Response = Invoke-TervisShipEngineAPI -Endpoint labels -Resource $LabelId

    try {
        $ZPLRaw = Invoke-WebRequest -Uri $Response.Content.label_download.zpl
        return [System.Text.Encoding]::ASCII.GetString($ZPLRaw.Content)
    }
    catch {
        Throw "Could not retrieve label $LabelId" 
    }
}

function New-TervisShipEngineAddress {
    param (
        [Parameter(ValueFromPipelineByPropertyName, Mandatory)]$Name,
        [Parameter(ValueFromPipelineByPropertyName)]$Phone,
        [Parameter(ValueFromPipelineByPropertyName)]$CompanyName,
        [Parameter(ValueFromPipelineByPropertyName, Mandatory)]$AddressLine1,
        [Parameter(ValueFromPipelineByPropertyName)]$AddressLine2,
        [Parameter(ValueFromPipelineByPropertyName)]$AddressLine3,
        [Parameter(ValueFromPipelineByPropertyName, Mandatory)]$CityLocality,
        [Parameter(ValueFromPipelineByPropertyName, Mandatory)]$StateProvince,
        [Parameter(ValueFromPipelineByPropertyName, Mandatory)]$PostalCode,
        [Parameter(ValueFromPipelineByPropertyName, Mandatory)]$CountryCode,
        [ValidateSet("unknown", "yes", "no")]$AddressResidentialIndicator = "unknown"
    )
 
    [PSCustomObject]@{
        name                          = $Name
        phone                         = $Phone
        company_name                  = $CompanyName
        address_line1                 = $AddressLine1
        address_line2                 = $AddressLine2
        address_line3                 = $AddressLine3
        city_locality                 = $CityLocality
        state_province                = $StateProvince
        postal_code                   = $PostalCode
        country_code                  = $CountryCode
        address_residential_indicator = $AddressResidentialIndicator
    }
}

function New-TervisShipEnginePackage {
    param (
        $WeightInLB
    )

    [PSCustomObject]@{
        weight = [PSCustomObject]@{
            value = $WeightInLB
            unit  = "pound"
        }
    }
}

function Remove-TervisShipEngineLabel {
    param (
        [Parameter(Mandatory)]$LabelId
    )

    Invoke-TervisShipEngineAPI -Endpoint labels -Method Put -Resource $LabelId -Subresource void
}

function Get-TervisShipEngineShipments {
    Invoke-TervisShipEngineAPI -Endpoint shipments -Method Get
}

function Get-TervisShipEngineWarrantyOrderService {
    param (
        $WeightInLB
    )

    switch ([system.decimal]::Parse($WeightInLB)) {
        { $_ -lt 1 } { return "usps_first_class_mail" }
        { $_ -le 10 } { return "ups_surepost_1_lb_or_greater" }
        { $_ -gt 10 } { return "ups_ground" }
        Default { throw "Weight input error" }
    }
}

function Get-TervisShipEngineWarehouseId {
    param (
        [Parameter(Mandatory)]
        [ValidateSet("NorthPort_Returns","Osprey_Returns")]
        $Location
    )

    $Response = Invoke-TervisShipEngineAPI -Endpoint warehouses
    if ($Response.Status -eq 200) {
        return $Response.Content.warehouses | 
        Where-Object name -eq $Location |
        Select-Object -ExpandProperty warehouse_id
    }
    else {
        Throw "Could not retrieve warehouse location."
    }
}