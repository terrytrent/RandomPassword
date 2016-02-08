#requires -Version 3
function Get-RandomPassword()
{
    <#
            .SYNOPSIS
                Generates Random Passwords based on a Password Type selection
            .DESCRIPTION
                Generates Random Passwords based on a Password Type selection:
                
                Random Characters creates a password of specified length, or if none is specfied 16 characters in length.
                Switches can be used to modify the characters in the password.

                Random Words creates a password from a list of words provided by http://www.freescrabbledictionary.com/

                Diceware Passphrase creates a password based on the methods described at http://world.std.com/~reinhold/diceware.html
            .EXAMPLE
                C:\PS>Get-RandomPassword -PasswordType 'Random Characters'
                I}k8MmX2.'F]7L7D
            .EXAMPLE
                C:\PS> Get-RandomPassword -PasswordType 'Random Characters' -Length 8
                jSOEdZ0>
            .EXAMPLE
                C:\PS> Get-RandomPassword -PasswordType 'Random Characters' -Length 8 -UpperCase -LowerCase
                RsqCWkrq
            .EXAMPLE
                C:\PS> Get-RandomPassword -PasswordType 'Word List'
                OdderAlbuminoidsArreeding
            .EXAMPLE
                C:\PS> Get-RandomPassword -PasswordType 'Diceware Passphrase'
                season; quay belt~ our% waist jewish9 tg aide4
            .OUTPUTS
                System.String
            .INPUTS
                None
            .LINK
                https://github.com/terrytrent/RandomPassword
            .LINK
                http://world.std.com/~reinhold/diceware.html
            #>    
    Param(
        
        [parameter(Mandatory = $true,Position = 0)]
        [ValidateSet('Random Characters','Word List','Diceware Passphrase')]
        [string]$PasswordType,
        
        [int]$Length,
        
        [alias('U')]
        [Switch]$UpperCase,

        [alias('L')]
        [Switch]$LowerCase,

        [alias('N')]
        [Switch]$Numbers,

        [alias('S')]
        [Switch]$Symbols
        
    )

    Setup
    
    if($PasswordType -eq 'Random Characters')
    {
        Generate-Password
    }
    elseif($PasswordType -eq 'Word List')
    {
        Get-WordListPassword
    }
    elseif($PasswordType -eq 'Diceware Passphrase')
    {
        Get-DicewarePassphrase
    }

    return $RandomPassword
}
function Setup()
{
    Declare-Variables
    Test-ComplexitySwitches
    Validate-Parameters
    Set-AsciiCodes
}
function Validate-Parameters()
{
    if((($PasswordType -eq 'Word List') -or ($PasswordType -eq 'Diceware Passphrase')) -and ($ComplexitySwitchesEnabled -eq $true))
    {
        $message = "Cannot use Complexity Switches with '$PasswordType' Passwords."
        throw($message)
        break
    }
    
    if(($PasswordType -eq 'Word List') -or ($PasswordType -eq 'Diceware Passphrase'))
    {
        write-warning -Message "You have selected a '$PasswordType' Password - this will take a few moments to generate"
    }
    
    if((($PasswordType -eq 'Word List') -or ($PasswordType -eq 'Diceware Passphrase')) -and ($Length -ne $null))
    {
        write-warning -Message "Cannot use Length Parameter with '$PasswordType' Password Type. Ignoring Length Parameter."
        Set-Variable -Name Length -Scope 2 -Value $null
    }
        
    if(($PasswordType -eq 'Random Characters') -and ($Length -eq $null))
    {
        Write-Warning -Message "Length was not set.  Cannot use Null Length with 'Random Characters' Password Type.`nDefaulting to 16 Characters."
        Set-Variable -Name Length -Scope 2 -Value 16
    }
        
    if(($PasswordType -eq 'Random Characters') -and (($Length -ne $null) -or ($Length -lt 8)))
    {
        $lengthType = ($Length.GetType()).name
        if(($lengthType -ne 'Int32') -or ($Length -lt 8))
        {
            $message = "The variable cannot be validated because the value '$Length' is not a valid value for the Length Parameter.`nPlease choose a number between 8 and 2048."
            throw("$message")
            break
        }
    }
}
function Declare-Variables()
{
    New-Variable -Name RandomPassword -Value $null -Scope 2
    New-Variable -Name CharacterTypeTotals -Value @{} -Scope 2
    New-Variable -Name AsciiTotals -Value @{} -Scope 2
    
    $thisScript = $MyInvocation.MyCommand
    try
    {
        New-Variable -Name scriptLocation -Scope 2 -Value $(Split-Path -Path $thisScript.path -Parent)
    }
    catch
    {
        New-Variable -Name scriptLocation -Scope 2 -Value 'C:\temp\Scripts\Get-WordsList'
        if(-Not (Test-Path $scriptLocation))
        {
            $null = New-Item -Path $scriptLocation -ItemType directory
        }
    }
    
    New-Variable -Name wordsUri -Value 'http://www.freescrabbledictionary.com/sowpods/download/sowpods.txt' -Scope 2
    New-Variable -Name lastModifiedFile -Value "$scriptLocation\lastmodified.txt" -Scope 2
    New-Variable -Name wordsContentFileFull -Value "$scriptLocation\WordList.csv" -Scope 2
    New-Variable -Name DicewareWordsFileFullPath -Value "$scriptLocation\DicewareWordList.csv" -Scope 2
    New-Variable -Name NumberOfWordsInDicewarePassphrase -Value 8 -Scope 2
}
function Test-ComplexitySwitches()
{
    $ComplexitySwitches = @($UpperCase, $LowerCase, $Numbers, $Symbols)
        
    if($ComplexitySwitches -contains $true)
    {
        Set-Variable -Name ComplexitySwitchesEnabled -Value $true -Scope 2
    }
    else
    {
        Set-Variable -Name ComplexitySwitchesEnabled -Value $false -Scope 2
    }
}
function Set-AsciiCodes()
{    
    $LowerCaseAsciiCodes = Set-LowerCaseAsciiCodes
    $UpperCaseAsciiCodes = Set-UpperCaseAsciiCodes
    $SymbolsAsciiCodes = Set-SymbolsAsciiCodes
    $NumbersAsciiCodes = Set-NumbersAsciiCodes
    New-Variable -Name ASCIICodes -Scope 2 -Value $(foreach($_ in ($LowerCaseAsciiCodes, $UpperCaseAsciiCodes, $SymbolsAsciiCodes, $NumbersAsciiCodes))
        {
            $_
        }
    )
}
function Set-AsciiCodesIfComplexitySpecified()
{
    Param(
        $Switch,
        $AsciiCodes
    )
    if($ComplexitySwitchesEnabled)
    {
        if($Switch)
        {
            return $AsciiCodes
        }
    }
    else
    {
        return $AsciiCodes
    }
}
function Set-LowerCaseAsciiCodes()
{
    $LowerCaseValues = $(97..122)
    $LowerCaseAsciiCodes = Set-AsciiCodesIfComplexitySpecified -Switch $LowerCase -AsciiCodes $LowerCaseValues
    return $LowerCaseAsciiCodes
}
function Set-UpperCaseAsciiCodes()
{
    $UpperCaseValues = $(65..90)
    $UpperCaseAsciiCodes = Set-AsciiCodesIfComplexitySpecified -Switch $UpperCase -AsciiCodes $UpperCaseValues
    return $UpperCaseAsciiCodes
}
function Set-SymbolsAsciiCodes()
{
    $SymbolsValues = foreach($_ in ($(33..47), $(58..64), $(91..96), $(123..126)))
        {
            $_
        }
    $SymbolsAsciiCodes = Set-AsciiCodesIfComplexitySpecified -Switch $Symbols -AsciiCodes $SymbolsValues
    return $SymbolsAsciiCodes
}
function Set-NumbersAsciiCodes()
{
    $NumbersValues = $(48..57)
    $NumbersAsciiCodes = Set-AsciiCodesIfComplexitySpecified -Switch $Numbers -AsciiCodes $NumbersValues
    return $NumbersAsciiCodes
}
function Generate-Password()
{
    do
    {
        Declare-GeneratePasswordLoopVariables
        $RandomAsciiCodes = Get-RandomCodesFromAsciiCodes
        Set-Variable -Name RandomPassword -Value $(Convert-AsciiCodesToChar -RandomAsciiCodes $RandomAsciiCodes) -Scope 1
        $rerun = Test-PasswordComplexity
    }
    while($rerun -eq $true)
}
function Declare-GeneratePasswordLoopVariables()
{
    Set-Variable -Name RandomAsciiCodes -Value $null -Scope 1
    Set-Variable -Name rerun -Value $false -Scope 1
}
function Get-RandomCodesFromAsciiCodes()
{
    $TotalLength = $(Get-Variable -Name length -ValueOnly -scope 2)
    $AllAsciiCodes = $(Get-Variable -Name AsciiCodes -ValueOnly -scope 2)
    $RandomAsciiCodes = @()
    foreach($value in (1..$TotalLength))
    {
        $Seed = Get-CryptoSeed
        $RandomAsciiCodes += $(Get-Random -InputObject $AllAsciiCodes -Count 1 -SetSeed $Seed)
    }
    
    return $RandomAsciiCodes
}
function Get-CryptoSeed()
{
    $bytes = New-Object -TypeName byte[] -ArgumentList (4)
    $RNG = [Security.Cryptography.RNGCryptoServiceProvider]::Create()
    $RNG.GetBytes($bytes)
    
    $totalInt = 1
    foreach($int in $bytes)
    {
        $totalInt = $totalInt * $int
    }
    $RandomSeed = ($totalInt / 2) -1
    return $RandomSeed
}
function Convert-AsciiCodesToChar()
{
    Param(
        $RandomAsciiCodes
    )
    foreach($char in $(0..$RandomAsciiCodes.Length))
    {
        $passwordAsciiArray += @([char]($RandomAsciiCodes[$char]))
    }
        
    $passwordAsciiJoined = ($passwordAsciiArray -join '').substring(0,$RandomAsciiCodes.count)

    return $passwordAsciiJoined
}
function Test-PasswordComplexity()
{
    Declare-PasswordComplexityVariables
    
    foreach($asciiCode in $RandomPasswordAsArray)
    {
        Generate-AsciiTotals -value $asciiCode
    }

    Generate-CharacterTypeTotals
    
    foreach($value in $CharacterTypeTotals)
    {
        if($value -lt 1 -and $value -ne $null)
        {
            $rerun = $true
        }
    }
    return $rerun
}
function Declare-PasswordComplexityVariables()
{
    Set-Variable -Name CharacterTypeTotals -Value @() -Scope 1
    Set-Variable -Name AsciiTotals -Scope 3 -Value $(New-Object -TypeName psobject -Property @{
            LowerCase = 0
            UpperCase = 0
            Symbols   = 0
            Numbers   = 0
    })
    Set-Variable -Name RandomPasswordAsArray -Value @([char[]]$RandomPassword) -Scope 1
}
function Generate-AsciiTotals()
{
    Param(
        $value
    )
        
    if($ComplexitySwitchesEnabled)
    {
        Test-AsciiCodesOfValueBasedOnSwitches -value $value
    }
    else
    {
        Test-AsciiCodesOfValue -value $value
    }
}
function Test-AsciiCodesOfValueBasedOnSwitches()
{
    Param(
        $value
    )
                
    if($LowerCase)
    {
        $LowerCaseTotal = $AsciiTotals.LowerCase
        $LowerCasePlus = Test-IfValueIsLowerCase -value $value
        $AsciiTotals.LowerCase = $LowerCaseTotal + $LowerCasePlus
    }
        
    if($UpperCase)
    {
        $UpperCaseTotal = $AsciiTotals.UpperCase
        $UpperCasePlus = Test-IfValueIsUpperCase -value $value
        $AsciiTotals.UpperCase = $UpperCaseTotal + $UpperCasePlus
    }
        
    if($Symbols)
    {
        $SymbolTotal = $AsciiTotals.Symbols
        $SymbolPlus = Test-IfValueIsSymbol -value $value
        $AsciiTotals.Symbols = $SymbolTotal + $SymbolPlus
    }
        
    if($Numbers)
    {
        $NumbersTotal = $AsciiTotals.Numbers
        $NumbersPlus = Test-IfValueIsNumber -value $value
        $AsciiTotals.Numbers = $NumbersTotal + $NumbersPlus
    }
}
function Test-AsciiCodesOfValue()
{
    Param(
        $value
    )
    $LowerCaseTotal = $AsciiTotals.LowerCase
    $UpperCaseTotal = $AsciiTotals.UpperCase
    $SymbolTotal = $AsciiTotals.Symbols
    $NumbersTotal = $AsciiTotals.Numbers
                        
    $LowerCasePlus = Test-IfValueIsLowerCase -value $value
    $UpperCasePlus = Test-IfValueIsUpperCase -value $value
    $SymbolPlus = Test-IfValueIsSymbol -value $value
    $NumbersPlus = Test-IfValueIsNumber -value $value
        
    $AsciiTotals.LowerCase = $LowerCaseTotal + $LowerCasePlus
    $AsciiTotals.UpperCase = $UpperCaseTotal + $UpperCasePlus
    $AsciiTotals.Symbols = $SymbolTotal + $SymbolPlus
    $AsciiTotals.Numbers = $NumbersTotal + $NumbersPlus
}
function Test-IfValueIsLowerCase()
{
    Param(
        $value
    )
    
    $LowerCaseAsciiCodes = Set-LowerCaseAsciiCodes
    $asciiCodeForValue = [int][char]($value)
    
    if($LowerCaseAsciiCodes -contains $asciiCodeForValue)
    {
        return 1
    }
    else
    {
        return 0
    }
}
function Test-IfValueIsUpperCase()
{
    Param(
        $value
    )
        
    $UpperCaseAsciiCodes = Set-UpperCaseAsciiCodes
    $asciiCodeForValue = [int][char]($value)
    
    if($UpperCaseAsciiCodes -contains $asciiCodeForValue)
    {
        return 1
    }
    else
    {
        return 0
    }
}
function Test-IfValueIsSymbol()
{
    Param(
        $value
    )
        
    $SymbolAsciiCodes = Set-SymbolsAsciiCodes
    $asciiCodeForValue = [int][char]($value)
    
    if($SymbolAsciiCodes -contains $asciiCodeForValue)
    {
        return 1
    }
    else
    {
        return 0
    }
}
function Test-IfValueIsNumber()
{
    Param(
        $value
    )
        
    $NumberAsciiCodes = Set-NumbersAsciiCodes
    $asciiCodeForValue = [int][char]($value)
    
    if($NumberAsciiCodes -contains $asciiCodeForValue)
    {
        return 1
    }
    else
    {
        return 0
    }
}
function Generate-CharacterTypeTotals()
{
    if($ComplexitySwitchesEnabled)
    {
        if($LowerCase)
        {
            Set-Variable -Name CharacterTypeTotals -Value @($AsciiTotals.LowerCase) -Scope 1
        }
            
        if($UpperCase)
        {
            Set-Variable -Name CharacterTypeTotals -Value @($CharacterTypeTotals[0], $AsciiTotals.UpperCase) -Scope 1
        }
            
        if($Symbols)
        {
            Set-Variable -Name CharacterTypeTotals -Value @($CharacterTypeTotals[0], $CharacterTypeTotals[1], $AsciiTotals.Symbols) -Scope 1
        }
            
        if($Numbers)
        {
            Set-Variable -Name CharacterTypeTotals -Value @($CharacterTypeTotals[0], $CharacterTypeTotals[1], $CharacterTypeTotals[2], $AsciiTotals.Numbers) -Scope 1
        }
    }
    else
    {
        Set-Variable -Name CharacterTypeTotals -Value @($AsciiTotals.LowerCase, $AsciiTotals.UpperCase, $AsciiTotals.Symbols, $AsciiTotals.Numbers) -Scope 1
    }
}
function Get-WordListPassword()
{
    Get-WordFileUIRHeaders
    Get-LastFileModifiedContent
    $wordsLastModified = $wordsRequest.Headers.'Last-Modified'

    if(($wordsLastModified -ne $localLastModified) -or (-Not (Test-Path -Path $wordsContentFileFull)))
    {
        Generate-WordList
        Export-WordList
        Export-LastModifiedFile
    }
    else
    {
        Import-WordListFromCSV
    }

    Set-Variable -Name RandomPassword -Value $(Get-RandomWords) -Scope 1    
}
function Get-WordFileUIRHeaders()
{
    try
    {
        New-Variable -Name wordsRequest -Value $(Invoke-WebRequest -Uri $wordsUri -Method head -ErrorAction SilentlyContinue -UserAgent ([Microsoft.PowerShell.Commands.PSUserAgent]::InternetExplorer)) -Scope 1
    }
    catch
    {
        try
        {
            $wordsContent = Import-Csv -Path $wordsContentFileFull
        }
        catch
        {
            break
        }
        return $wordsContent
        break
    }
}
function Get-LastFileModifiedContent()
{
    New-Variable -Name localLastModified -Value $(Get-Content -Path $lastModifiedFile -ErrorAction SilentlyContinue) -Scope 1
}
function Generate-WordList()
{
    New-Variable -Name wordsContent -Scope 1 -Value $(Get-WordListFromURI)
}
function Get-WordListFromURI()
{
    $list = (Invoke-WebRequest -Uri $wordsUri -UserAgent ([Microsoft.PowerShell.Commands.PSUserAgent]::InternetExplorer)).content -split '\n' | Select-Object -Skip 2 -Property @{
        label      = 'WordList'
        expression = {
            $_
        }
    }
    return $list
}
function Export-WordList()
{
    $wordsContent | Export-Csv -NoTypeInformation -Path $wordsContentFileFull
}
function Export-LastModifiedFile()
{
    $wordsLastModified | Out-File -FilePath $lastModifiedFile
}
function Import-WordListFromCSV()
{
    try
    {
        New-Variable -Name wordsContent -Scope 1 -Value $(Import-Csv -Path $wordsContentFileFull)
    }
    catch
    {
        break
    }
}
function Get-RandomWords()
{
    $RandomNumbers = Get-RandomNumbers
    $PasswordWords = ''
    foreach($value in $RandomNumbers)
    {
        [string]$Word = ($wordsContent[$value]).WordList
        [string]$WordFirstLetter = $Word[0]
        [string]$WordRestOfWord = $Word.substring(1)
        $PasswordWords += ($WordFirstLetter).ToUpper()
        $PasswordWords += $WordRestOfWord.ToLower()
    }
    return $PasswordWords
}
function Get-RandomNumbers()
{
    $Seed = Get-CryptoSeed
    $RandomNumbers = Get-Random -InputObject (0..$($wordsContent.count)) -Count 3 -SetSeed $Seed
    return $RandomNumbers
}
function Get-DicewarePassphrase()
{
    do{
        $PassphraseArray=@()
        foreach($number in (1..$NumberOfWordsInDicewarePassphrase))
        {
            $DicewareWord=Get-DicewareWord
            $PassphraseArray+=$DicewareWord
        }
        $Passphrase=$PassphraseArray -join ' '
        
        $PassphraseIsLongEnough=Test-PassphraseLength -Passphrase $Passphrase
    }
    while($PassphraseIsLongEnough -eq $false )
    
    Set-Variable -Name RandomPassword -Value $Passphrase -Scope 1
}
function Test-PassphraseLength()
{
    Param(
        $Passphrase
    )
    $PassphraseWithoutSpaces=$Passphrase.replace(' ','')
    $ExpectedMinimumPassphraseLength=17
    $PassphraseLength=($PassphraseWithoutSpaces | Measure-Object -Character).Characters
    
    if($PassphraseLength -lt $ExpectedMinimumPassphraseLength)
    {
        return $false
    }
    else
    {
        return $true
    }
}
function Get-DicewareWord()
{
    $SidesOfDice=6
    $NumberOfRoles=5
    $WordList=Get-DicewareWordList
    $Number=Get-DicewareNumber
    $RandomBool=Get-RandomBool
    
    $Word=$($WordList | where {$_.Number -eq $Number}).Word
    if($RandomBool)
    {
        $SpecialorDigit=Get-SpecialOrDigit
        $WordWithSpecialOrDigit="$Word$SpecialOrDigit"
        return $WordWithSpecialOrDigit
    }
    else
    {
        return $Word
    }
}
function Get-RandomBool()
{
    $Bool=Get-RandomDiceRoll -min 0 -max 2
    switch($Bool)
    {
        0 {return $false}
        1 {return $true}
    }
}
function Get-SpecialOrDigit()
{
    $SidesOfDice=6
    $Number1=Get-RandomDiceRoll -min 0 -max $SidesOfDice
    $Number2=Get-RandomDiceRoll -min 0 -max $SidesOfDice
    $DicewareSpecialOrDigitTable=Generate-DicewareSpecialOrDigitTable
    
    $Character=$DicewareSpecialOrDigitTable[$number1][$number2]
    return $Character
}
function Generate-DicewareSpecialOrDigitTable()
{
    $DicewareSpecialOrDigitTable=@(
        @('~','!','#','$','%','^'),
        @('&','*','(',')','-','='),
        @('+','[',']','\','{','}'),
        @(':',';','"',"'",'<','>'),
        @('?','/','0','1','2','3'),
        @('4','5','6','7','8','9')
    )
    return $DicewareSpecialOrDigitTable
}
function Get-DicewareWordList()
{
    if(test-path $DicewareWordsFileFullPath)
    {
        $DicewareWordList=Import-DicewareWordList
        $WordListTestsGood=Test-ImportedDicewareWordList -WordList $DicewareWordList
        
        if($WordListTestsGood -eq $false)
        {
            $DicewareWordList=DownloadAndSave-DicewareWordList
        }
    }
    else
    {
        $DicewareWordList=DownloadAndSave-DicewareWordList
    }
    return $DicewareWordList
}
function Test-ImportedDicewareWordList()
{
    Param(
        $WordList
    )
    
    $WordListNoteProperties=($WordList | Get-Member -MemberType NoteProperty).Name
    
    $WordListContainsExpectedProperties=(($WordListNoteProperties -contains 'Word') -and ($WordListNoteProperties -contains 'Number'))
    $WordListContainsExpectedCount=($WordList.count -eq 7776)
    
    if($WordListContainsExpectedProperties -and $WordListContainsExpectedCount)
    {
        return $true
    }
    else
    {
        return $false
    }
}
function DownloadAndSave-DicewareWordList()
{
    $DicewareWordList=Download-DicewareWordList
    Export-DicewareWordListToCSV -WordList $DicewareWordList
    return $DicewareWordList
}
function Download-DicewareWordList()
{
    $WordList=Invoke-WebRequest -Uri http://world.std.com/~reinhold/diceware.wordlist.asc -UserAgent ([Microsoft.PowerShell.Commands.PSUserAgent]::InternetExplorer)
    $WordListContent=$WordList.Content -split "`n" | select -skip 2 | select -first 7776
    
    $NumberAndWordListArray=@()
    foreach($Line in $WordListContent)
    {
        $Number=$Line.substring(0,5)
        $Word=$Line.substring(6)
        $NumberAndWord = New-Object -TypeName psobject
        $NumberAndWord | Add-Member -MemberType NoteProperty -Name Number -Value $Number
        $NumberAndWord | Add-Member -MemberType NoteProperty -Name Word -Value $Word
        
        $NumberAndWordListArray+=$NumberAndWord
    }
    
    return $NumberAndWordListArray
}
function Export-DicewareWordListToCSV()
{
    Param(
        $WordList
    )
    $WordList | Export-Csv -Path $DicewareWordsFileFullPath -NoTypeInformation
}
function Import-DicewareWordList()
{
    $NumberAndWordListArray=Import-Csv -Path $DicewareWordsFileFullPath
    return $NumberAndWordListArray
}
function Get-DicewareNumber()
{
    Set-Variable -Name DicewareNumber -Value @()
    foreach($number in (1..$NumberOfRoles))
    {
        $DicewareNumber+=Get-RandomDiceRoll -min 1 -max $($SidesOfDice + 1)
    }
    $DicewareNumber=$DicewareNumber -join ''
    
    return $DicewareNumber
}
function Get-RandomDiceRoll()
{
    Param(
        $min,
        $max
    )
    $Seed=Get-CryptoSeed
    
    $Roll=Get-Random -SetSeed $Seed -Minimum $min -Maximum $max
    
    return $Roll
}
