# KITA
## Kita is Interface Tools and Activator. (Office)

KITA is a script written in powershell. Facilitates office installation and activation.

<h2 align="center"><img src="https://external-content.duckduckgo.com/iu/?u=https%3A%2F%2Fwallpapercave.com%2Fwp%2Fwp11814665.jpg&f=1&nofb=1&ipt=f7dc49d41242c19471f495dc08823bfadf58601edc7fafca389ddb6355d824d6&ipo=images" width="600"></h2>

### Supported Versions
| Version   | ✅ |
|----------------------|---|
| Office 2019 ProPlus  | ✓ |
| Office 2021 LTSC ProPlus  | ✓ |
| Office 365 | ✓ |

### Usage
```powershell
iwr -useb "cutt.ly/kita" | iex
```
### Shorter
```powershell
irm cutt.ly/kita | iex
```
#### Issues TLS/SSL (LTSB Versions)
```powershell
[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; iwr -useb "cutt.ly/kita" | iex
```
