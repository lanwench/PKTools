
#requires -Version 4
function Get-PKWeather {
<#
.SYNOPSIS
    Retrieves current weather conditions for a specified location using the OpenWeatherMap API.

.DESCRIPTION
    The 'Get-PKWeather' function fetches weather data from OpenWeatherMap API 2.5 for a given location. 
    This was written mainly as a learning exercise to understand how to use the OpenWeatherMap API and PowerShell's Invoke-WebRequest cmdlet.
    It supports fetching weather information by postal code and country, or by using the current computer's geographic location 
    The function uses the OpenWeatherMap API to retrieve weather details such as temperature, humidity, wind speed, and more.
    The optional -GetForecast parameter allows for retrieving a 3-day weather forecast with the selection of the time of day for the forecast.

.NOTES
    Name    : Function_Get-PKWeather.ps1
    Created : 2025-04-24
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2025-04-24 - Created script

.PARAMETER Country
    Specifies the 2-character country code/abbreviation when not using auto-discover for location; if -Postcode is not provided, uses the default location or capital

.PARAMETER PostCode
    If -Country is specified, specifies the postal/zip code of the location for which to retrieve weather data 

.PARAMETER ApiToken
    Specifies the API token for accessing the OpenWeatherMap API (v 2.5 only). 
    If not provided, a default test token is used. 
    The token must be at least 20 characters long and must not contain whitespace.

.PARAMETER Unit
    Specifies the unit of measurement for temperature. Acceptable values are: Fahrenheit, Celsius, Kelvin, Metric, Imperial.
    
.OUTPUTS
    [PSCustomObject]
        Outputs a custom object containing the following properties:
        - Location: The name of the location (city or coordinates)
        - Country: The 2-character country code
        - Date: The local date and time of the report
        - Summary: A summary of the current weather conditions
        - TimeZone: The timezone of the location
        - Temperature: The current temperature (in specified unit)
        - FeelsLike: The "feels like" temperature
        - High: The maximum temperature
        - Low: The minimum temperature
        - Humidity: The humidity percentage
        - Cloudiness: The cloudiness percentage
        - Condition: The weather condition and description
        - WindSpeed: The wind speed
        - WindDirection: The wind direction in degrees
        - Forecast: The 3-day weather forecast (if -GetForecast is specified)
        - Coordinates: The latitude and longitude of the location
        - Sunrise: The time of sunrise (local time)
        - Sunset: The time of sunset (local time)

.NOTES
    - Requires PowerShell version 4 or higher.
    - The function uses the OpenWeatherMap API and requires an API token for access.
    - The default API token provided is for testing purposes and may have limited functionality.

.LINK
    https://www.reddit.com/r/PowerShell/comments/j7z7ox/get_current_weather_conditions_with_powershell/

.LINK
    https://openweathermap.org/api/one-call-api

.EXAMPLE
    PS C:\> Get-PKWeather
    Retrieves the current weather for the computer's current geographic location, returing the temperature in Fahrenheit

        Location      : Greenwood
        Country       : US
        Date          : 2025-04-28T20:34:41
        TimeZone      : (UTC-4)
        Temperature   : 75 F
        FeelsLike     : 74 F
        High          : 78 F
        Low           : 74 F
        Humidity      : 31%
        Cloudiness    : 5%
        Condition     : Clear (clear sky)
        WindSpeed     : 4
        WindDirection : 4
        Forecast      : n/a
        Coordinates   : 38.0415/-78.7833
        Sunrise       : 2025-04-28T06:22:04
        Sunset        : 2025-04-28T20:02:51


.EXAMPLE
    PS C:\> Get-PKWeather -Country GB -Postcode "DT6 3AA" -Unit C
    Retrieves the current weather for the specified postal code in Great Britain, returning the temperature in Celsius                                                                 
    
        Location      : DT6 3AA
        Country       : GB
        Date          : 2025-04-28T20:40:07
        TimeZone      : (UTC1)
        Temperature   : 12 C
        FeelsLike     : 11 C
        High          : 13 C
        Low           : 11 C
        Humidity      : 72%
        Cloudiness    : 62%
        Condition     : Clouds (broken clouds)
        WindSpeed     : 2.47
        WindDirection : 1
        Forecast      : n/a
        Coordinates   : 50.7344/-2.7596
        Sunrise       : 2025-04-28T05:50:29
        Sunset        : 2025-04-28T20:26:17


#>
[CmdletBinding(DefaultParameterSetName = "Location")]
param (
    
    [Parameter(
        ParameterSetName = "Named",
        Position = 0,
        Mandatory,
        HelpMessage = "2-character ISO-3116 country code"
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({If ($_.Length -eq 2) {$True} Else {Throw "Please provide the two-character ISO-3116 country code"}})]
    [string]$Country,

    
    [Parameter(
        ParameterSetName = "Named",
        Position = 1,
        HelpMessage = "Postal code for location; requires -Country (if not specified, defaults to country capital)"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("ZipCode","Zip")]
    [string]$Postcode,

    [Parameter(
        HelpMessage = "API token for OpenWeatherMap 2.5 (default is a test token)"
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({If (-not ($_.Length -gt 20 -and $_ -notmatch '\s')) {Throw "Invalid API token"}})]
    [string]$ApiToken="814bf0eb3a9db3d11be051039669b7d7",

    [Parameter(
        HelpMessage = "Temperature unit (default is Fahrenheit/Imperial)"
    )]
    [ValidateSet("Imperial","Fahrenheit","F","C","Celsius","Centigrade","Kelvin","Metric")]
    [string]$Unit = "Fahrenheit",

    [Parameter(
        HelpMessage = "Include 1-7 day weather forecast (default is 3 days)"
    )]
    [switch]$GetForecast,

    [Parameter(
        HelpMessage = "Number of days if -GetForecast is specified (between 1 and 5; default is 3)"
    )]
    [ValidateRange(1 ,5)]
    [int]$ForecastDays = 3,

    [Parameter(
        HelpMessage = "Time of day for forecast 'snapshot' if -GetForecast is specified; provides specific weather description but doesn't affect high/low for that day (default is 12 noon local time)"
    )]
    [ValidateSet("00","03","06","09","12","15","18","21","Midnight","3AM","6AM","9AM","Noon","12PM","3PM","6PM","9PM")]
    [string]$ForecastTime = "Noon",

    [Parameter(
        HelpMessage = "If -GetForecast is specified, this switch consolidates the forecast into a single summary string within the main output object. If this switch is NOT used with -GetForecast, current weather is output, followed by separate objects for each forecast day."
    )]
    [switch]$SummarizeForecast

)
Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # How did we get here
    $Source = $PSCmdlet.ParameterSetName
    $ScriptName = $MyInvocation.MyCommand.Name
    [switch]$PipelineInput = $MyInvocation.ExpectingInput

    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where-Object { $CurrentParams.keys -notContains $_ } | 
    Where-Object { Test-Path variable:$_ } | Foreach-Object {
        $CurrentParams.Add($_, (Get-Variable $_).value)
    }
    $CurrentParams.Add("PipelineInput", $PipelineInput)
    $CurrentParams.Add("ParameterSetName", $Source)
    $CurrentParams.Add("ScriptName", $ScriptName)
    $CurrentParams.Add("ScriptVersion", $Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"
    

    #region normalize/standardize some variables/params
    
    # Temperature Units
    
    # GEmini
    # Temperature Units
    Switch -Regex ($Unit) {
        '^Kelvin$' { # Matches "Kelvin" exactly
            $APIStr = "standard"
            $UnitSymbol = "K"
            $UnitStr = "Kelvin"
            $WindUnit = "m/s"
        }
        '^(Imperial|Fahrenheit|F)$' { # Matches "Imperial" OR "Fahrenheit" OR "F"
            $APIStr = "imperial"
            $UnitSymbol = "F"
            $UnitStr = "Fahrenheit/Imperial" # General string for this group
            $WindUnit = "mph"
        }
        '^(C|Celsius|Centigrade|Metric)$' { # Matches "C" OR "Celsius" OR "Centigrade" OR "Metric"
            $APIStr = "metric"
            $UnitSymbol = "C"
            $UnitStr = "Celsius/Centigrade/Metric" 
            $WindUnit = "m/s"
        }
    }

    <#
    Switch ($Unit) {
        Kelvin {
            $APIStr = "standard"
            $UnitSymbol = "K"
            $UnitStr = "Kelvin"
            $WindUnit = "m/s"
        }
        Imperial {
            $APIStr = "imperial"
            $UnitSymbol = "F"
            $UnitStr = "Fahrenheit/Imperial"
            $WindUnit = "mph"
        }
        Fahrenheit {
            $APIStr = "imperial"
            $UnitSymbol = "F"
            $UnitStr = "Fahrenheit/Imperial"
            $WindUnit = "mph"
        }
        F {
            $APIStr = "imperial"
            $UnitSymbol = "F"
            $UnitStr = "Fahrenheit/Imperial"
            $WindUnit = "mph"
        }
        C {
            $APIStr = "metric"
            $UnitSymbol = "C"
            $UnitStr = "Celsius/Centigrade/Metric"
            $WindUnit = "m/s"
        }
        Celsius {
            $APIStr = "metric"
            $UnitSymbol = "C"
            $UnitStr = "Celsius/Centigrade/Metric"
            $WindUnit = "m/s"
        }
        Centigrade {
            $APIStr = "metric"
            $UnitSymbol = "C"
            $UnitStr = "Celsius/Centigrade/Metric"
            $WindUnit = "m/s"
        }
        Metric {
            $APIStr = "metric"
            $UnitSymbol = "C"
            $UnitStr = "Celsius/Centigrade/Metric"
            $WindUnit = "m/s"
        }
    }
#>

    # Time for forecast (out of the 8 available times returned)
    
    If ($GetForecast.IsPresent) {
        Switch -Regex ($ForecastTime) {
            "00|midnight"   { $FTime = 00 }
            "03|3AM"        { $FTime = 03 }
            "06|6AM"        { $FTime = 06 }
            "09|9AM"        { $FTime = 09 }
            "12|noon|12PM"  { $FTime = 12 }
            "15|3PM"        { $FTime = 15 }
            "18|6PM"        { $FTime = 18 }
            "21|9PM"        { $FTime = 21 }
        }

        Switch -Regex ($FTime) {
            00        {$FTimeStr = "midnight" }
            03        {$FTimeStr = "3AM" }
            06        {$FTimeStr = "6AM" }
            09        {$FTimeStr = "9AM" }
            12        {$FTimeStr = "noon"}
            15        {$FTimeStr = "3PM" }
            18        {$FTimeStr = "6PM" }
            21        {$FTimeStr = "9PM" }
        }
    }
    #endregion

    #region Inner functions
    Function _GetLocation {
        <#
        .LINK
            # https://stackoverflow.com/questions/46287792/powershell-getting-gps-coordinates-in-windows-10-using-windows-location-api#46287884
        #>
        [Cmdletbinding()]
        Param()
        Try {
            Write-Verbose "Adding System.Device assembly (required to access Location namespace)"
            Add-Type -AssemblyName System.Device -ErrorAction Stop  
            Write-Verbose "Creating GeoCoordinateWatcher object"
            $GeoWatcher = New-Object System.Device.Location.GeoCoordinateWatcher 
            Write-Verbose "Starting GeoCoordinateWatcher and waiting for discovery to complete within 100 ms"
            $GeoWatcher.Start() 
            While (($GeoWatcher.Status -ne 'Ready') -and ($GeoWatcher.Permission -ne 'Denied')) {Start-Sleep -Milliseconds 100 }  
            If ($GeoWatcher.Permission -eq 'Denied') {Throw "Access to location data is denied; Please check your privacy settings"} 
            Else {
                Write-Output $Geowatcher.Position.Location | Select-Object Latitude,
                Longitude,
                Altitude,@{N="Label";E={"$([math]::Round($_.Latitude, 2)) Lat/$([math]::Round($_.Longitude, 2)) Lon"}}
            }
        }
        Catch {Throw "Failed to get current location! $($_.Exception.Message))"}
        Finally {$GeoWatcher.Dispose()}
    }

    Function _ValidateCountryCode {
        <#
        .LINK
            https://www.apicountries.com/countries
        #>
        [Cmdletbinding()]
        Param(
            [Parameter(Mandatory,ValueFromPipeline,Position=0)]
            [string]$Country
        )
        $Msg = "Invoking API call to validate country code"
        Write-Verbose $Msg
        Try {
            $CountryURI = "https://www.apicountries.com/countries"
            $AllCountries  = Invoke-webrequest -uri $CountryURI -ErrorAction Stop -Verbose:$False -ProgressAction SilentlyContinue
            $LookupTable = $AllCountries.Content | ConvertFrom-JSON -ErrorAction Stop | Select-Object Name,Capital,Region,Alpha2Code,Alpha3Code,LatLng | Sort-Object Name
            If (-not ($CountryLookup = $LookupTable | Where-Object { $_.Alpha2Code -eq $Country.ToUpper() -or $_.Alpha3Code -eq $Country.ToUpper() } | Select-Object -First 1)) {
                $Msg = "Invalid country code '$($Country.ToUpper())' - please provide a valid two-character ISO-3166 country code"
                Throw $Msg
            }
            Else {
                $Msg = "Valid country code '$($CountryLookup.Alpha2Code)' ($($CountryLookup.Name))"
                Write-Verbose $Msg
                Write-Output $CountryLookup
            }
        }
        Catch {
            $Msg = "Failed to get country list from $CountryURI - $($_.Exception.Message)"
            Throw $Msg
        }
    } # end _validateCountryCode
    Function _FormatTemp{
        Param([Parameter(ValueFromPipeline,Position=0)]$Temp)
        "$([math]::Round($Temp)) $UnitSymbol"
    }

    function _FormatTime {
        param(
            [Parameter(Mandatory,ValueFromPipeline,Position=0)]
            [int]$DT
        )
        # Convert the Unix timestamp to a DateTimeOffset object (which defaults to UTC)
        $dateTimeOffsetUtc = [DateTimeOffset]::FromUnixTimeSeconds($DT)
        # Apply the timezone offset.  We create a TimeSpan from the offset.
        $localDateTimeOffset = $dateTimeOffsetUtc.ToOffset([TimeSpan]::FromSeconds($JSON.timezone))
        # Convert the DateTimeOffset to a regular DateTime object.
        $localDateTime = $localDateTimeOffset.DateTime
        return $localDateTime
    }

    Function _FormatTimeZone {
        [CmdletBinding()]
        Param (
            [Parameter(Mandatory,ValueFromPipeline,ValueFromPipelineByPropertyName,Position=0)]        
            [int]$Timezone
        )
        $OffsetInHours = $Timezone / 3600
        $TZOffsetStr = "(UTC$OffsetInHours)"
        Write-Verbose "The UTC offset is: $TZOffsetStr"
        Write-Output $TZOffsetStr
    }

    Function _ConvertDateTime {
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory, ValueFromPipeline, Position = 0)]
            $WeatherObject
        )
    
        $TimeZoneOffsetSeconds = $WeatherObject.timezone
        $TimeZoneOffset = [TimeSpan]::FromSeconds($TimeZoneOffsetSeconds)
    
        $TimeUnix = $WeatherObject.dt
        $TimeUTC = [datetime]::UnixEpoch.AddSeconds($TimeUnix)
        $TimeLocation = $TimeUTC.Add($TimeZoneOffset)
        $FormattedTimeLocation = $TimeLocation.ToString("dddd MMMM dd yyyy, HH:mm:ss")
    
        $SunriseTimeUnix = $WeatherObject.sys.sunrise
        $SunriseTimeUTC = [datetime]::UnixEpoch.AddSeconds($SunriseTimeUnix)
        $SunriseTimeLocation = $SunriseTimeUTC.Add($TimeZoneOffset)
        $FormattedSunriseTimeLocation = $SunriseTimeLocation.ToString("dddd MMMM dd yyyy, HH:mm:ss")
    
        $SunsetTimeUnix = $WeatherObject.sys.sunset
        $SunsetTimeUTC = [datetime]::UnixEpoch.AddSeconds($SunsetTimeUnix)
        $SunsetTimeLocation = $SunsetTimeUTC.Add($TimeZoneOffset)
        $FormattedSunsetTimeLocation = $SunsetTimeLocation.ToString("dddd MMMM dd yyyy, HH:mm:ss")
    
        [PSCustomObject]@{
            Location                = $WeatherObject.name
            TimeZone                = $WeatherObject.timezone
            #OffsetSeconds           = $TimeZoneOffsetSeconds
            TimeUnix                = $TimeUnix
            TimeUTC                 = $TimeUTC.ToString("yyyy-MM-dd HH:mm:ss")
            TimeLocation            = $TimeLocation.ToString("yyyy-MM-dd HH:mm:ss")
            FormattedTimeLocation   = $FormattedTimeLocation
            SunriseTimeUnix         = $SunriseTimeUnix
            SunriseTimeUTC          = $SunriseTimeUTC.ToString("yyyy-MM-dd HH:mm:ss")
            SunriseTimeLocation     = $SunriseTimeLocation.ToString("yyyy-MM-dd HH:mm:ss")
            FormattedSunriseTimeLocation = $FormattedSunriseTimeLocation
            SunsetTimeUnix          = $SunsetTimeUnix
            SunsetTimeUTC           = $SunsetTimeUTC.ToString("yyyy-MM-dd HH:mm:ss")
            SunsetTimeLocation      = $SunsetTimeLocation.ToString("yyyy-MM-dd HH:mm:ss")
            FormattedSunsetTimeLocation = $FormattedSunsetTimeLocation
        }
    }

    # not currently using this but it might be fun later so I'm keeping it
    Function _SaveImage {
        [CmdletBinding(SupportsShouldProcess,ConfirmImpact = "High")]
        Param([Parameter(ValueFromPipeline,Position=0)]$JSON,[string]$OutPath = $Env:Temp)
        Try {
            $URI = "http://openweathermap.org/img/w/$($JSON.weather.icon).png"
            Write-Verbose "Invoking API call to $URI"
            $response = Invoke-WebRequest -Uri $URI
            $imageBytes = $response.Content
            $imageObject = New-Object PSObject -Property @{ImageBytes = $imageBytes}
            $OutFile = "$OutPath\openweather_$($JSON.weather.icon).png"
            Write-Verbose "Saving image to $OutFile"
            If ($PSCmdlet.ShouldProcess($Outfile,"Save OpenWeather icon to file")) {
                [System.IO.File]::WriteAllBytes($OutFile, $imageBytes)
                Get-Item $OutFile -Force
            }
            Else {
                Write-Verbose "Operation cancelled by user"
                $ImageObject
            }
        }
        Catch {
            Throw $_.Exception.Message
        }
    }

    #endregion Functions

    Switch ($Source) {
        Zip {$Msg = "Getting weather for zip code $Postcode $($Country.ToUpper())"}
        Location {$Msg = "Getting weather for current location"}                       
    }
    $Msg += " in $UnitStr units"
    Write-Verbose "[BEGIN: $ScriptName] $Msg"
}
Process {

    Try{
        
        Switch ($Source) {
            Named {
                $Msg = "Validating country code '$($Country.ToUpper())'"
                Write-Verbose $Msg
                If ($CountryLookup = _ValidateCountryCode -Country $Country.ToUpper() -Verbose:$False) {
                    $Msg = "Valid country code '$($CountryLookup.Alpha2Code)' ($($CountryLookup.Name))"
                    Write-Verbose $Msg

                    If (-not $Postcode) {
                        $Lat = $CountryLookup.LatLng[0]
                        $Lon = $CountryLookup.LatLng[1]
                        $Msg = "No postal code provided; using default country latitude and longitude ($Lat, $Lon)"
                        Write-Verbose $Msg
                        $Msg = "Creating URI string"
                        Write-Verbose $Msg
                        [string]$WeatherURI="https://api.openweathermap.org/data/2.5/weather?lat=$Lat&lon=$Lon&appid=$($ApiToken)&units=$APIStr"
                        If ($GetForecast.IsPresent) {
                            [string]$ForecastURI = "https://api.openweathermap.org/data/2.5/forecast?lat=$Lat&lon=$Lon&appid=$($ApiToken)&units=$APIStr"
                        }
                    }
                    Else {
                        $Msg = "Creating URI string"
                        Write-Verbose $Msg
                        [string]$WeatherURI="https://api.openweathermap.org/data/2.5/weather?zip=$($Postcode),$($Country)&appid=$($ApiToken)&units=$APIStr"
                        If ($GetForecast.IsPresent) {
                            [string]$ForecastURI = "https://api.openweathermap.org/data/2.5/forecast?zip=$($Postcode),$($Country)&appid=$($ApiToken)&units=$APIStr"
                        }
                    }
                }
                Else {
                    $Msg = "Invalid country code '$($Country.ToUpper())' - please provide a valid two-character ISO-3166 country code"
                    Write-Error $Msg
                } 
            }
            Location {
                $Msg = "Getting current location coordinates"
                Write-Verbose $Msg
                $Location = _GetLocation -Verbose:$False
                $Msg = "Creating URI string"
                Write-Verbose $Msg
                [string]$WeatherURI="https://api.openweathermap.org/data/2.5/weather?lat=$($Location.Latitude)&lon=$($Location.Longitude)&appid=$($ApiToken)&units=$APIStr"
                If ($GetForecast.IsPresent) {
                    [string]$ForecastURI = "https://api.openweathermap.org/data/2.5/forecast?lat=$($Location.Latitude)&lon=$($Location.Longitude)&appid=$($ApiToken)&units=$APIStr"
                }
            }
        }

        If ($WeatherURI) {

            $Msg = "Invoking API call to $WeatherURI"
            Write-Verbose $Msg
            $Webrequest = Invoke-WebRequest -Uri $WeatherURI -Method Get -Erroraction Stop -Verbose:$False

            If ($Webrequest.StatusCode -eq 200){
                $WeatherObject = ConvertFrom-Json -InputObject $($Webrequest.Content)
                $Date = $WeatherObject | _ConvertDateTime
                $CurrentWeatherStr = "As of $($Date.FormattedTimeLocation) local time, the temperature in $($WeatherObject.Name) is $($WeatherObject.main.temp | _FormatTemp) and conditions are '$(($WeatherObject.weather.main)) ($($WeatherObject.weather.description))'"
                Write-Verbose $CurrentWeatherStr

                If ($GetForecast.IsPresent) {

                    $Msg = "Fetching and processing $ForecastDays`-day forecast data "
                    Write-Verbose $Msg
                    $ForecastRequest = Invoke-WebRequest -Uri $ForecastURI -ErrorAction Stop -Verbose:$False
                    $ForecastObj = $ForecastRequest.Content | ConvertFrom-Json -ErrorAction Stop -Verbose:$False
                    
                    If ($null -ne $ForecastObj.list) {  # Get the timezone offset for the forecast's city (in seconds from UTC)

                        $ForecastArray = @()    
                        #$CityTimezoneOffsetSeconds = $ForecastObj.city.timezone
                        $CityTimezoneOffset = [TimeSpan]::FromSeconds($ForecastObj.city.timezone)
                        #$CurrentUniversalTime = (Get-Date).ToUniversalTime()  # Determine "today's date" accurately at the forecast location
                        $CurrentTimeAtForecastLocation = ((Get-Date).ToUniversalTime()).Add($CityTimezoneOffset)
                        $TodayAtForecastLocation = $CurrentTimeAtForecastLocation.Date # This is a [datetime] object representing today at 00:00

                        # Add local date information to each forecast item - dt_text is UTC, so we convert to local using offset
                        $AllForecastItemsWithLocalDate = $ForecastObj.List | ForEach-Object {
                            $UtcDateTime = [datetime]$_.dt_txt 
                            $LocalDateTimeAtSlot = $UtcDateTime.Add($CityTimezoneOffset) # Convert slot time to local
                            $_ | Add-Member -MemberType NoteProperty -Name LocalDateTime -Value $LocalDateTimeAtSlot -PassThru |
                                Add-Member -MemberType NoteProperty -Name LocalDate -Value $LocalDateTimeAtSlot.Date -PassThru
                        }

                        $GroupedByFutureDate = $AllForecastItemsWithLocalDate | 
                            Where-Object { $_.LocalDate -gt $TodayAtForecastLocation } | # Only take dates strictly AFTER today
                            Group-Object LocalDate | Sort-Object Name
                        $UniqueFutureDates = $GroupedByFutureDate | Select-Object -First $ForecastDays 

                        ForEach ($DateGroup In $UniqueFutureDates) {

                            $DailyForecastSlots = $DateGroup.Group
                            $RepresentativeSlot = $DailyForecastSlots | Sort-Object {[Math]::Abs($_.LocalDateTime.Hour - $FTime)} | Select-Object -First 1
                            If ($RepresentativeSlot -and $RepresentativeSlot.weather.Count -gt 0) {
                                $Description = $RepresentativeSlot.weather[0].description.Substring(0, 1).ToUpper() + $RepresentativeSlot.weather[0].description.Substring(1).ToLower()
                                $HumidityValue = "$($RepresentativeSlot.main.humidity)%"
                                $WindSpeedValue = $RepresentativeSlot.wind.speed
                            }

                            $Result = [PSCustomObject]@{
                                Date         = Get-Date $DateGroup.Name -f "ddd dd MMM yy"
                                SnapshotTime = $FTime
                                Low          = ($DailyForecastSlots | Measure-Object -Property {$_.main.temp_min} -Minimum).Minimum  | _FormatTemp -Verbose:$False
                                High         = ($DailyForecastSlots | Measure-Object -Property {$_.main.temp_max} -Maximum).Maximum  | _FormatTemp -Verbose:$False
                                Humidity     = $HumidityValue ?? "Error"
                                Windspeed    = $WindSpeedValue ?? "Error"
                                Description  = $Description ?? "Error"
                                OneLiner     = $Null 
                            }

                            $Result.OneLiner = "$($Result.Date): High $($Result.High), low $($Result.Low). At $FTimeStr`: $($Result.Description.ToLower()), humidity $($Result.Humidity), wind $($Result.Windspeed) $WindUnit"
                            $ForecastArray += $Result
                        }
                    }
                    If ($ForecastArray.Count -gt 0) {
                        $Forecast = $ForecastArray
                        If ($SummarizeForecast.IsPresent) {
                            $Forecast = $Forecast.OneLiner -join "`n"
                        }
                        #Else {
                        #    $Forecast = $Forecast | Select-Object Date, SnapshotTime, Low, High, Humidity, Windspeed, Description, OneLiner
                        #}
                    } 
                    Else {
                        $Forecast = "$ForecastDays`-day forecast data (for future dates) not available or could not be processed."
                        $Msg = "Could not retrieve or process detailed $ForecastDays-day forecast for future dates!"
                        Write-Warning $Msg
                    }
                }
                Else {$Forecast = "n/a"}
                
<#                
                # Gemini 2 - this is perfect! Trying to squish it down more in the next try.

                If ($GetForecast.IsPresent) {
                    $Msg = "Fetching and processing $ForecastDays`-day forecast data "
                    Write-Verbose $Msg
                    $ForecastRequest = Invoke-WebRequest -Uri $ForecastURI -ErrorAction Stop
                    $ForecastObj = $ForecastRequest.Content | ConvertFrom-Json -ErrorAction Stop
                    $ForecastOutputArray = @()

                    If ($null -ne $ForecastObj.list) {  # Get the timezone offset for the forecast's city (in seconds from UTC)
                        
                        $CityTimezoneOffsetSeconds = $ForecastObj.city.timezone
                        $CityTimezoneOffset = [TimeSpan]::FromSeconds($CityTimezoneOffsetSeconds)
                        $CurrentUniversalTime = (Get-Date).ToUniversalTime()  # Determine "today's date" accurately at the forecast location
                        $CurrentTimeAtForecastLocation = $CurrentUniversalTime.Add($CityTimezoneOffset)
                        $TodayAtForecastLocation = $CurrentTimeAtForecastLocation.Date # This is a [datetime] object representing today at 00:00

                        #Write-Verbose "Current date at forecast location is $($TodayAtForecastLocation.ToString('yyyy-MM-dd')). Filtering out this date for the forecast."

                        # Add local date information to each forecast item - dt_text is UTC, so we convert to local using offset
                        $AllForecastItemsWithLocalDate = $ForecastObj.List | ForEach-Object {
                            $UtcDateTime = [datetime]$_.dt_txt 
                            $LocalDateTimeAtSlot = $UtcDateTime.Add($CityTimezoneOffset) # Convert slot time to local
                            $_ | Add-Member -MemberType NoteProperty -Name LocalDateTime -Value $LocalDateTimeAtSlot -PassThru |
                                Add-Member -MemberType NoteProperty -Name LocalDate -Value $LocalDateTimeAtSlot.Date -PassThru
                        }

                        $GroupedByFutureDate = $AllForecastItemsWithLocalDate | 
                            Where-Object { $_.LocalDate -gt $TodayAtForecastLocation } | # Only take dates strictly AFTER today
                            Group-Object LocalDate | Sort-Object Name

                        # Select up to the first $ForecastDays unique *future* days
                        $UniqueFutureDates = $GroupedByFutureDate | Select-Object -First $ForecastDays

                        ForEach ($DateGroup In $UniqueFutureDates) {
                            $DailyForecastSlots = $DateGroup.Group
                            $CurrentDayDate = $DateGroup.Name # This is a [datetime] object (the date part of a future day)

                            # Calculate the actual min and max temperatures for the entire day
                            $DailyLow = ($DailyForecastSlots | Measure-Object -Property {$_.main.temp_min} -Minimum).Minimum
                            $DailyHigh = ($DailyForecastSlots | Measure-Object -Property {$_.main.temp_max} -Maximum).Maximum

                            # Find the forecast slot closest to the user-specified $FTime
                            $RepresentativeSlot = $DailyForecastSlots | Sort-Object {[Math]::Abs($_.LocalDateTime.Hour - $FTime)} | Select-Object -First 1
                            
                            $Description = "Conditions vary"
                            $HumidityValue = "N/A"
                            $WindSpeedValue = "N/A"

                            If ($RepresentativeSlot -and $RepresentativeSlot.weather.Count -gt 0) {
                                $Description = $RepresentativeSlot.weather[0].description.Substring(0, 1).ToUpper() + $RepresentativeSlot.weather[0].description.Substring(1).ToLower()
                                $HumidityValue = "$($RepresentativeSlot.main.humidity)%"
                                $WindSpeedValue = $RepresentativeSlot.wind.speed
                            }

                            $FormattedDailyHigh = $DailyHigh | _FormatTemp -Verbose:$False
                            $FormattedDailyLow  = $DailyLow  | _FormatTemp -Verbose:$False
                            $ForecastDateStr = Get-Date $CurrentDayDate -f "ddd dd MMM yy"
                            $ForecastOutputArray += "$ForecastDateStr`: High of $FormattedDailyHigh, low of $FormattedDailyLow. At $($ForecastTime.Substring(0, 1).ToLower() + $ForecastTime.Substring(1)), $($Description.ToLower()), humidity $HumidityValue, wind $WindSpeedValue $WindUnit"
                        }
                    }
                    If ($ForecastOutputArray.Count -gt 0) {
                        $Forecast = $ForecastOutputArray
                    } Else {
                        $Forecast = "$ForecastDays`-day forecast data (for future dates) not available or could not be processed."
                        $Msg = "Could not retrieve or process detailed $ForecastDays-day forecast for future dates!"
                        Write-Warning $Msg
                    }
                }
                Else {$Forecast = "n/a"}

#>                

# ... (rest of the Process block)

<#
                # Gemini 1
                If ($GetForecast.IsPresent) {
                    $Msg = "Fetching and processing $ForecastDays`-day forecast data"
                    Write-Verbose $Msg
                    $ForecastRequest = Invoke-WebRequest -Uri $ForecastURI -ErrorAction Stop
                    $ForecastObj = $ForecastRequest.Content | ConvertFrom-Json -ErrorAction Stop
                    $ForecastOutputArray = @() # Use a new array name

                    If ($null -ne $ForecastObj.list) {
                        # Get the timezone offset from the forecast city data (in seconds)
                        $CityTimezoneOffsetSeconds = $ForecastObj.city.timezone
                        $CityTimezoneOffset = [TimeSpan]::FromSeconds($CityTimezoneOffsetSeconds)

                        # Add local date information to each forecast item and group by that local date
                        $GroupedByLocalDate = $ForecastObj.List | ForEach-Object {
                            # dt_txt is UTC. Convert to local time using the city's timezone offset.
                            $UtcDateTime = [datetime]$_.dt_txt
                            $LocalDateTime = $UtcDateTime.Add($CityTimezoneOffset)
                            $_ | Add-Member -MemberType NoteProperty -Name LocalDateTime -Value $LocalDateTime -PassThru |
                                Add-Member -MemberType NoteProperty -Name LocalDate -Value $LocalDateTime.Date -PassThru
                        } | Group-Object LocalDate

                        # Select up to the first 3 unique days from the forecast
                        $UniqueDates = $GroupedByLocalDate | Sort-Object Name | Select-Object -First $ForecastDays

                        ForEach ($DateGroup In $UniqueDates) {
                            $DailyForecastSlots = $DateGroup.Group
                            $CurrentDayDate = $DateGroup.Name # This is a DateTime object (the date part)

                            # Calculate the actual min and max temperatures for the entire day
                            $DailyLow = ($DailyForecastSlots | Measure-Object -Property {$_.main.temp_min} -Minimum).Minimum
                            $DailyHigh = ($DailyForecastSlots | Measure-Object -Property {$_.main.temp_max} -Maximum).Maximum

                            # Find the forecast slot closest to the user-specified $FTime for description, humidity, wind
                            # $FTime is set in the Begin block (e.g., 0, 3, 6, 12 for Noon, etc.)
                            $RepresentativeSlot = $DailyForecastSlots | Sort-Object {[Math]::Abs($_.LocalDateTime.Hour - $FTime)} | Select-Object -First 1
                            
                            $Description = "Conditions vary" # Default description
                            $HumidityValue = "N/A"
                            $WindSpeedValue = "N/A"

                            If ($RepresentativeSlot -and $RepresentativeSlot.weather.Count -gt 0) {
                                $Description = $RepresentativeSlot.weather[0].description.Substring(0, 1).ToUpper() + $RepresentativeSlot.weather[0].description.Substring(1).ToLower()
                                $HumidityValue = "$($RepresentativeSlot.main.humidity)%"
                                $WindSpeedValue = $RepresentativeSlot.wind.speed # This is a number
                            }

                            $FormattedDailyHigh = $DailyHigh | _FormatTemp -Verbose:$False
                            $FormattedDailyLow  = $DailyLow  | _FormatTemp -Verbose:$False
                            
                            $ForecastDateStr = Get-Date $CurrentDayDate -f "ddd dd MMM yyyy"

                            $ForecastOutputArray += "$ForecastDateStr`: $Description; daily high of $FormattedDailyHigh, low of $FormattedDailyLow. Humidity: $HumidityValue; Wind: $WindSpeedValue $WindUnit"
                        }
                    }

                    
                    If ($ForecastOutputArray.Count -gt 0) {
                        $Forecast = $ForecastOutputArray
                        #$Forecast = $ForecastOutputArray -join "`n`t`t" # `"`n`t`t"` for nice formatting in verbose, or use `"`n`"` for the actual property
                    } Else {
                        $Forecast = "$ForecastDays`-day forecast data not available or could not be processed."
                        Write-Warning "Could not retrieve or process detailed $ForecastDays-day forecast."
                    }
                }
                Else {$Forecast = "n/a"}

                #>

<#

# Gemini is proposing an alternative to this original - 

                If ($GetForecast.IsPresent) {
                    $ForecastRequest = Invoke-WebRequest -Uri $ForecastURI -Verbose:$False
                    $ForecastObj = $ForecastRequest.Content | ConvertFrom-Json -ErrorAction Stop
                    $Forecast = @()
                    $ForecastObj.List | Where-Object {(Get-Date $_.dt_txt).Hour -eq $FTime} | Select-Object -First 3 | Foreach-Object {
                        $Day = $_
                        $Description = $Day.weather.description.Substring(0, 1).ToUpper() + $Day.weather.description.Substring(1).ToLower()
                        $Temp        = $Day.main.temp | _FormatTemp -Verbose:$False
                        $High        = $Day.main.temp_max | _FormatTemp -Verbose:$False
                        $Low         = $Day.main.temp_min | _FormatTemp -Verbose:$False
                        $Humidity    = "$($Day.main.humidity)%"
                        #$Forecast    += "$(Get-Date $Day.dt_txt -Format "ddd dd MMM yyyy, HH:mm")`: $Description, with an expected high of $High and low of $Low, $Humidity humidity, and projected wind speed of $($Day.wind.speed) mph"
                        $Forecast    += "$(Get-Date $Day.dt_txt -Format "ddd dd MMM yyyy, HH:mm")`: $Description, with an expected temperature of $Temp, $Humidity humidity, and projected wind speed of $($Day.wind.speed) mph"
                    }
                    $Forecast = $Forecast -join "`n"
                }
                Else {$Forecast = "n/a"}

#>                

                [PSCustomObject]@{
                    Location     = $($WeatherObject.name)
                    Country       = $WeatherObject.sys.country.ToUpper()
                    Date          = $Date.TimeLocation
                    Summary       = $CurrentWeatherStr
                    TimeZone      = &{_FormatTimeZone -Timezone $WeatherObject.timezone}
                    Temperature   = $WeatherObject.main.temp | _FormatTemp -Verbose:$False
                    FeelsLike     = $WeatherObject.main.feels_like | _FormatTemp -Verbose:$False
                    High          = $WeatherObject.main.temp_max | _FormatTemp -Verbose:$False
                    Low           = $WeatherObject.main.temp_min | _FormatTemp -Verbose:$False
                    Humidity      = "$($WeatherObject.main.humidity)%"
                    Cloudiness    = "$($WeatherObject.Clouds.All)%"
                    Condition     = "$($WeatherObject.weather.main) ($($WeatherObject.weather.description))"
                    WindSpeed     = $WeatherObject.wind.speed
                    WindDirection = [Math]::Round($WeatherObject.wind.deg / 45)
                    Forecast      = $Forecast
                    Coordinates   = "$($WeatherObject.coord.lat),$($WeatherObject.coord.lon)"
                    Sunrise       = $Date.SunriseTimeLocation
                    Sunset        = $Date.SunsetTimeLocation
                }
            }
            Else {
                $Msg = "HTTP CODE : $($Webrequest.StatusCode)"
                Write-Warning $Msg
            }
        } #end if uri
    }
    Catch{
        $Msg = "Operation failed! $($_)"
        Write-Error $Msg  
    }
} #end process
} # end Get-PKWeather
