# PS-CFW
PowerShell-CommonFrameWork

# Git Basic process / usage.
---=== Short version ===---<BR>
 1) Open PowerShell (NOT as admin).<BR>
 2) Close all web browsing sessions.<BR>
 3) $objWebClient = New-Object System.Net.WebClient;<BR>
 4) $objCreds = Get-Credential;<BR>
  &nbsp; &nbsp; a) Enter Windows credentials.<BR>
 5) $objWebClient.Proxy.Credentials = $objCreds;<BR>
 6) C:;<BR>
 7) cd C:\Projects\<Project.to.work>;<BR>
 8) Open "https://github.com" in a web browser.<BR>
 9) git fetch origin;<BR>
 &nbsp; &nbsp; a) provide GitHub username.<BR>
 &nbsp; &nbsp; b) provide GitHub password.<BR>
10) git pull;<BR>
 &nbsp; &nbsp; a) provide GitHub username.<BR>
 &nbsp; &nbsp; b) provide GitHub password.<BR>
11) Close all web browsing sessions.<BR>
12) --- Make updates/changes ---<BR>
13) Open "https://github.com" in a web browser.<BR>
14) git commit -a -m "Brief desc of update";<BR>
15) git push origin master;<BR>
 &nbsp; &nbsp; a) provide GitHub username.<BR>
 &nbsp; &nbsp; b) provide GitHub password.<BR>
16) Close "https://github.com" (close all web browsers).<BR>
17) Done.<BR>

# Git Install / Config
