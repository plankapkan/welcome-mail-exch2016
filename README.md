# welcome-mail-exch2016
This script sends welcome mail to new mailboxes and configure them.
It works for exchange DAG cluster.

To run script in scheduler use:
Start a program
C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe
argumets:
-command ". 'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1'; Connect-ExchangeServer -auto; C:\{PATH_TO_SCRIPT}\WelcomeMSG_v2.ps1