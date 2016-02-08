# Random Password Generator
First and foremost this script allows you to create three kinds of passwords:

1. Random Characters
2. Random Words
3. Diceware Passphrase

## Random Characters
The First kind of password, Random Characters, can be generated with the following command and switches, and any combination thereof:

    Get-RandomPassword -PasswordType 'Random Characters'
    Get-RandomPassword -PasswordType 'Random Characters' -Length 8
    Get-RandomPassword -PasswordType 'Random Characters' -LowerCase -UpperCase -Symbols -Numbers

For example, if I want to generate a password that is 24 characters long with only Lower Case letters, Symbols, and Numbers:

    Get-RandomPassword -PasswordType 'Random Characters' -Length 24 -Lowercase -Symbols -Numbers

Could Generate:

    j#c}fy/@&#-_fo>"&]9&,qvg

This is the fastest method to generate a password using this script.

## Random Words
The second kind of password, Random Words, can be generated with the following command:

    Get-RandomPassword -PasswordType 'Word List'

This command generates a password containing 3 words from the Free Scrabble Dictionary at http://www.freescrabbledictionary.com/

For Example:

    SixainesLaryngectomyWearer

This can take several minutes because the script needs to download the word list (over 20000 words), verify it is up to date, and then pull random words from this list.

## Diceware Passphrase
The third kind of password, Diceware Passphrase, can be generated with the following command:

    Get-RandomPassword -PasswordType 'DiceWare Passphrase'

This command generates a password (passphrase) containing 8 words and randomly assigned special characters or digits, as specified on http://world.std.com/~reinhold/diceware.html.

The word list used to generate the passphrase can be found at http://world.std.com/~reinhold/diceware.wordlist.asc.

For Example:

    tonic wavy0 brew{ noon8 echo adopt< hood

# Clean Code
This script is also a test in Clean Code, based on Uncle Bob's Clean Code (see http://blog.cleancoder.com/ and http://www.amazon.com/s/ref=nb_sb_noss_2?url=search-alias%3Dstripbooks&field-keywords=Clean+Code+Robert+C+Martin)

The book(s) referenced are written more towards actual coding languages, however I wanted to see if it would also apply to scripting languages such as PowerShell.

So far I have found that it makes things easier to write, and follow, at least as long as it is my own code.

Writing this way also allows me to drop in replacements if I find a way to do a part of the script better without making drastic changes.

For instance, if I find a better way to create a seed for randomness I just have to modify the function Get-CryptoSeed to return a different value for $RandomSeed.  It's easy to find in the code and allows for quick drop-ins.

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

I would love to get any input regarding using Clean Code in PowerShell, good or bad!  I want to find a way to make PowerShell easier to write, keep track of, and read for other authors!
