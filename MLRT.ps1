#You may need to tweak some of the commands in this function to work in your environment (file paths, office version/installation locations, etc)
function Get-SomeInput {
    $input = read-host "

    
    Type in the letter that corresponds with the action you want to run then press ENTER:

    NOTE: Commands prefixed with '*' will fail if this script isnt being run as admin.
          Commands prefixed with '^' will fail if the user you're logged in as is a not local admin.
  
  ###############################################################################################
  #|                                                                                           |#
  #|  [A.] Kill Skype                                                                          |#            
  #| ^[B.] Remove Microsoft Office Related Credentials From the Windows Credential Manager     |#
  #| ^[C.] Remove ALL Credentials From the Windows Credential Manager                          |#
  #|  [D.] List Stored Credentials in the Windows Credential Manager                           |#            
  #| *[E.] Remove ALL Office 365 License Keys                                                  |#
  #|  [F.] Re-Run Dstatus                                                                      |#
  #| *[G.] Remove a specific o365 license                                                      |#
  #|  [H.] Launch Microsoft word                                                               |#
  #|  [I.] Launch Microsoft word     **[SAFE MODE]**                                           |#
  #|  [J.] Launch Outlook                                                                      |#
  #|  [K.] Launch Outlook            **[SAFE MODE]**                                           |#
  #|  [L.] Launch o365 Diagnostics tool                                                        |#
  #| *[M.] Repair Office 365                                                                   |#
  #| *[N.] Uninstall Office 365                                                                |#
  #| *[O.] Re-Insall Office 365                                                                |#
  #|  [x.] EXIT                                                                                |#
  #|                                                                                           |#
  ###############################################################################################
  Enter a letter"

    switch ($input) `
    {
        'A' {
            $ProcessIsRunning = { Get-Process Lync* -ErrorAction SilentlyContinue }
            #This will run the get-process command whenever you call it by using the Invoke() method.
            if(!$ProcessIsRunning.Invoke()) {
            write-host "Skype is not running, let's continue." -ForegroundColor Green
            } else {
                Write-Host "Skpe is running" -ForegroundColor Red
                Write-Host "Give me a few second while I close Skype" -ForegroundColor Yellow
                Get-Process Lync* | Stop-Process
            }
            Get-SomeInput
        }
        'B' {
            #This will launch a pop-up to confirm whether or not to remove the o365 related credential fron the credential manager
            $shell = new-object -comobject "WScript.Shell"
            $choice = $shell.popup("Would you like to remove stored o365 credentials?",0,"Proceed?",4+32)

            if ($choice -eq 6)
            {
            

            Write-Host "Removing o365 credentials stored in windows credential manager." -ForegroundColor Yellow
            #This will Remove and stored credentials related to o365
            cmdkey /list | ForEach-Object{if($_ -like "*Target:*" -and $_ -like "*microsoft*"){cmdkey /del:($_ -replace " ","" -replace "Target:","")}} 
            cmdkey /list | ForEach-Object{if($_ -like "*Target:*" -and $_ -like "*outlook*"){cmdkey /del:($_ -replace " ","" -replace "Target:","")}}
            cmdkey /list | ForEach-Object{if($_ -like "*Target:*" -and $_ -like "*office*"){cmdkey /del:($_ -replace " ","" -replace "Target:","")}}
            cmdkey /list | ForEach-Object{if($_ -like "*Target:*" -and $_ -like "*onedrive*"){cmdkey /del:($_ -replace " ","" -replace "Target:","")}} 
            cmdkey /list | ForEach-Object{if($_ -like "*Target:*" -and $_ -like "*exchange*"){cmdkey /del:($_ -replace " ","" -replace "Target:","")}} 
            cmdkey /list | ForEach-Object{if($_ -like "*Target:*" -and $_ -like "*msteams*"){cmdkey /del:($_ -replace " ","" -replace "Target:","")}} 
            }

            if ($choice -eq 7)
            {
                Write-Host "You've chosen not to remove the o365 related windows credentials." -ForegroundColor Red
            }
            Get-SomeInput
        }
        'C' {
            #This will launch a pop-up to confirm whether or not to remove the o365 related credential fron the credential manager
            $shell = new-object -comobject "WScript.Shell"
            $choice = $shell.popup("Would you like to remove stored o365 credentials?",0,"Proceed?",4+32)

            if ($choice -eq 6)
            {

            Write-Host "Removing o365 credentials stored in windows credential manager." -ForegroundColor Yellow
            #This will Remove and stored credentials related to o365
            cmdkey /list | ForEach-Object{if($_ -like "*Target:*"){cmdkey /del:($_ -replace " ","" -replace "Target:","")}} 

            }

            if ($choice -eq 7)
            {
                Write-Host "You've chosen not to remove ALL stored credentials." -ForegroundColor Red
            }
            Get-SomeInput
        }
        'D' {
            CMDKEY /list
            $shell = new-object -comobject "WScript.Shell"
            $choice = $shell.popup("Would you like to launch Windows Credential Manager?",0,"Proceed?",4+32)

            if ($choice -eq 6)
            {
                rundll32.exe keymgr.dll,KRShowKeyMgr
            }
            
            if ($choice -eq 7)
            {
                Write-Host "You've chosen not to launch Windows Credential Manager." -ForegroundColor Red
            }
            Get-SomeInput
        }
        'E' {
            # store the license info into an array
            $license = cscript "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" /dstatus
            $key1 = "Office 16, Office16O365ProPlusR_Grace edition"

            Write-Output $license 

            #loop till the end of the array searching for the $o365 string
            for ($i=0; $i -lt $license.Length; $i++)
            {
            if ($license[$i] -match $key1)
            {

            $i += 5 #jumping five lines from "Product name" to get to the product key line in the array, check output of dstatus and adjust as needed for the product you are removing
            $keyline = $license[$i] # extra step but i would rather deal with the variable as a string than an array, could be removed i guess, efficiency is not my concern
            $prodkey1 = $keyline.substring($keyline.length - 5, 5) # getting the last 5 characters of the line (prodkey)

            #Displays License and key ofthe currently licensed o365 install
            Write-Host "PRODUCT: $key1" -ForegroundColor Green
            Write-Host "KEY: $prodkey1" -ForegroundColor Green

            #This will launch a pop-up to confirm whether or not to remove the office 365 license key
            $shell = new-object -comobject "WScript.Shell"
            $choice = $shell.popup("Would you like to remove $key1 : $prodkey1 ?",0,"Proceed?",4+32)

            if ($choice -eq 6)
            {  
                #removing the inactive key from the workstation
                cscript "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" /unpkey:$prodkey1
                Write-Host cscript "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" /dstatus -ForegroundColor Green
            }
            if ($choice -eq 7)
            {
                Write-Host "You've chosen not to remove the active office 365 license key." -ForegroundColor Red
            }
            }
            }

            #Removing active o365 key
            $Key2 = "Office 16, Office16O365ProPlusR_Subscription1 edition"

            #loop till the end of the array searching for the $o365 string
            for ($i=0; $i -lt $license.Length; $i++)
            {
            if ($license[$i] -match $key2)
            {
            $i += 6 #jumping six lines to get to the product key line in the array, check output of dstatus and adjust as needed for the product you are removing
            $keyline = $license[$i] # extra step but i would rather deal with the variable as a string than an array, could be removed i guess, efficiency is not my concern
            $prodkey = $keyline.substring($keyline.length - 5, 5) # getting the last 5 characters of the line (prodkey)

            #Displays License and key ofthe currently licensed o365 install
            Write-Host "PRODUCT: $key2" -ForegroundColor Green
            Write-Host "KEY: $prodkey" -ForegroundColor Green

            #This will launch a pop-up to confirm whether or not to remove the office 365 license key
            $shell = new-object -comobject "WScript.Shell"
            $choice = $shell.popup("Would you like to $key2 : $prodkey ?",0,"Proceed?",4+32)

            if ($choice -eq 6)
            {   
                #removing the key from the workstation
                cscript "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" /unpkey:$prodkey
            }

            if ($choice -eq 7)
            {
                Write-Host "You chose not to remove the active office 365 license key." -ForegroundColor Red
            }
            }
            }
            Get-SomeInput
            }
        'F' {
            cscript "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" /dstatus
            Get-SomeInput
            }
        'G' {
            $InKey = Read-Host -Prompt 'Which License do you want to remove?'
            cscript "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" /unpkey:$InKey
            Get-SomeInput
            }
        'H'{
            Start-Process winword.exe
            Get-SomeInput
            }
        'I'{
            Start-Process winword.exe /safe
            Get-SomeInput
           }
        'J'{
            Start-Process outlook.exe
            Get-SomeInput
            }
        'K'{
            Start-Process outlook.exe /safe
            Get-SomeInput
           }
        'L'{
            Start-Process https://outlookdiagnostics.azureedge.net/sarasetup/SetupProd.exe
            Get-SomeInput
           }

        #You can replace these with the corresponding .exe's for options M, N, and O and store the locally or on a network share if you prefer.
        'M'{
            & "C:\Program Files\common files\microsoft shared\Clicktorun\OfficeClickToRun.exe" scenario=Repair platform=x86 culture=en-us
            Get-SomeInput
           }
        'N'{
            #This links to microsofts tool for a FULL removal.
            Start-Process https://aka.ms/diag_officeuninstall

            # Uncomment the comman below for regular uninstall
            # & "C:\Program Files\common files\microsoft shared\Clicktorun\OfficeClickToRun.exe" scenario=install scenariosubtype=ARP sourcetype=none productstoremove=O365ProPlusRetail.16_en-us_x-none culture=en-us version.16=16.0
            Get-SomeInput
           }
        'O'{
            $shell = new-object -comobject "WScript.Shell"
            $choice = $shell.popup("Would you like to Re-Install Office 365?",0,"Proceed?",4+32)

            if ($choice -eq 6)
            {
                start-process https://portal.office.com/OLS/MySoftware.aspx#
            }
            if ($choice -eq 7)
            {
                Write-Host "You've chosen not to Re-Install Office 365." -ForegroundColor Red
            }
            Get-SomeInput
            }
        'X'{
            Read-Host -Prompt "Press Enter again to exit"
           }

        default {
            write-host 'You have entered and invalid response [, please try again.'
            Get-SomeInput
        }
    }
}

Get-SomeInput