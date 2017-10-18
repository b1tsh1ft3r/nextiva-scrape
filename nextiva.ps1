# Nextiva Page Scrape Utility
#
# This powershell script is purely for scraping the data from the nextiva phone system web login.
# Its purely for showing phone status on your own basic html page that can be displayed on a monitor
# or TV and gets around having their licensing to have exportable stats/api access. Its purely
# intended as a solution for low budget or no budget situations. IE/Powershell were chosen purely because
# its what was available when i wrote this tiny solution in our environment.

Import-Module .\WASP.DLL            # Import the WASP Powershell DLL for Send-Keys functions

$UPDATE_FREQUENCY = 3               # TIME IN SECONDS WE WILL UPDATE WITH "FRESH' DATA
$TIMEOUT_REFRESH = ((10*60)/$UPDATE_FREQUENCY) # 200=10Mins

$ie = New-Object -com InternetExplorer.Application
$ie.visible=$true

ECHO "[OPENING INTERNET EXPLODER]"
$ie.navigate("https://cp4.nextiva.com/callcenter")                  # Open Nextiva login page in IE

ECHO "[WAITING FOR PAGE TO LOAD...]"
while($ie.Busy -eq $true) { Start-Sleep -seconds 5; }               # Wait until IE Loads the page

$IE_WINDOW = Select-Window IEXPLORE | Set-WindowActive              # focus on IE Window

while($ie.Busy -eq $true) { Start-Sleep -seconds 2; }               # Wait until login is successful and page renders

ECHO "LOGIN TO NEXTIVA FIRST AND THEN PRESS ENTER."
pause
CLS

ECHO "[ENTERING MAIN LOOP OF PROGRAM]"

$TEMP = $TIMEOUT_REFRESH

Do {
     $a = $ie.document.body
     $b = $a -replace "<.*?>"                                       # remove some html tags
     $c = $b -replace "noWrap>Agents","JUNK"                        # replace with the phrase "noWrap>Agents" so we have a good pointer for "Agents" later since there are two total (eliminate one)
     $d = $c -split "Agents"                                        # now split where we find "Agents" (should be the only one in the file now...)
     $e = $d[2] -split "Speed"                                      # other split is at "Speed"

     # NOW REMOVE ALL SPACES, THEN ADJUST STATUSES AND OUTPUT TO FILE (make it look nice) (the long way to do it...)

     # html display method
     $a = $e[0]
     $b = $a -creplace " ",""
     $c = $b -creplace "Unavailable" , " - <img src=""images/unavailable.jpg""> <br>"
     $d = $c -creplace "Available" ,   " - <img src=""images/available.jpg""> <br>"
     $e = $d -creplace "Wrap-Up",      " - <img src=""images/wrapup.jpg""> <br>"
     $f = $e -creplace "unknown",      " - <img src=""images/signout.jpg""> <br>"
     $1 = $f -creplace "Sign-Out",     " - <img src=""images/signout.jpg""> <br>"

     # INSERT HEADER at top of output html file (banner of sorts)
     # ----------------------------------------------------------
         $HEADER = "<center><head> <meta http-equiv=""refresh"" content=""10""><hr><img src=""images/cd.png""></head><hr><b><u>Phone Status</u></b><hr>"
         $HEADER | Out-File .\TEMP.TXT
     # ----------------------------------------------------------

     $1 | Out-File -Append .\TEMP.TXT               # Modified to include "-Append" for header stuff above
     Get-Content .\TEMP.TXT | where {$_ -ne ""} > .\INDEX.HTML

     # INSERT FOOTER at bottom of output html file (banner of sorts)
     # with MOTD text to display.
     # ----------------------------------------------------------
         $FOOTER = "</center><hr><center><b><u>Messages of the day</u></b><hr><embed src=""message.txt"" width=1024><hr></center>"
         $FOOTER | Out-File -Append .\INDEX.HTML
     # ----------------------------------------------------------

     # move page into directory to serve via webserver
     move-item -force .\INDEX.HTML .\htdocs

     echo "[EXPORTED PHONE STATUS DATA...]"

     Start-Sleep $UPDATE_FREQUENCY

     $TEMP -=1;                                     # DECREMENT REFRESH_TIME

         if($TEMP -le 0)
         {
            ECHO "[REFRESHING PAGE TO STAY LOGGED IN...]"
            $TEMP = $TIMEOUT_REFRESH
            $IE_WINDOW = Select-Window IEXPLORE | Set-WindowActive
            $IE_WINDOW | Send-Keys "{F5}"
            while($ie.Busy -eq $true) { Start-Sleep -seconds 2; }
            ECHO "[PAGE REFRESHED]"
         }


       } While($true)
