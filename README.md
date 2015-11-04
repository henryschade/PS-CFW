# PS-CFW
PowerShell-CommonFrameWork

# Basic process / usage.
---=== Short version ===---
 1) Open PowerShell (NOT as admin).
 2) Close all web browsing sessions.
 3) $objWebClient = New-Object System.Net.WebClient;
 4) $objCreds = Get-Credential;
  a) Enter Windows credentials.
 5) $objWebClient.Proxy.Credentials = $objCreds;
 6) C:;
 7) cd C:\Projects\<Project.to.work>;
 8) Open "https://github.com" in a web browser.
 9) git fetch origin;
10) git pull;
11) --- Make updates/changes ---
12) git commit -a -m "Brief desc of update";
13) git push origin master;
  a) provide GitHub username.
  b) provide GitHub password.
14) Close "https://github.com" (close all web browsers).
15 Done.

# Install / Config
