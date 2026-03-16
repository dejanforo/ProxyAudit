## Name
proxyaudit.ps1

## Description
This PowerShell script collects Exchange proxy and Windows proxy configuration from all Exchange Servers in the organization and displays a report in tabular format. It also exports the result in tab separated file.

Sample output: 

    Timestamp           ServerName Role    Status             ExchProxy                       ExchExceptions                         WinHTTPProxy           WinHTTPBypass
    ---------           ---------- ----    ------             ---------                       --------------                         ------------           -------------
    2026-03-13 15:25:53 SERVER01   Mailbox Online             http://proxy.domain.com:8080/   mgmt.net, intranet.com, domainA.com    proxy.domain.com:8080  *.mgmt.net;*.intranet.com;domainA.com;<local>
    2026-03-13 15:25:57 SERVER02   Edge    Unreachable (Edge) None                            None                                   N/A                    N/A
    2026-03-13 15:25:59 SERVER03   Edge    Unreachable (Edge) None                            None                                   N/A                    N/A
    2026-03-13 15:26:00 SERVER04   Mailbox Online             http://proxy.domain.com:8080/   mgmt.net, intranet.com, domainA.com    proxy.domain.com:8080  *.mgmt.net;*.intranet.com;domainA.com;<local>
    2026-03-13 15:26:00 SERVER05   Mailbox Online             http://proxy.domain.com:8080/   mgmt.net, intranet.com, domainA.com    proxy.domain.com:8080  *.mgmt.net;*.intranet.com;domainA.com;<local>

    Audit complete. Results saved to: D:\Scripts\proxyaudit\ProxyAudit_20260313_1526.tsv

## Author
Dejan Foro, Exchangemaster GmbH  
E-mail: dejan.foro@exchangemaster.ch  
Web: https://www.exchangemaster.ch  
Github: https://github.com/dejanforo  
LinkedIn: https://www.linkedin.com/in/dejanforo/

## License
Free to use. Provided "as is", without warranty or guaranty of any kind. Use at your own risk. 

## Badges
![Static Badge](https://img.shields.io/badge/Shell-Powershell-blue) 
![Static Badge](https://img.shields.io/badge/Editor-VSCode-blue)
![Static Badge](https://img.shields.io/badge/AI-Gemini-blue)
![Static Badge](https://img.shields.io/badge/Made%20In-Switzerland-red?color=%20%23DA291C)
![Static Badge](https://img.shields.io/badge/Made%20by-Exchangemaster%20GmbH-blue)

