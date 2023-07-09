# KITA
## Kita is Interface Tools and Activator. (Office)

KITA is a script written in powershell. Facilitates office installation and activation.

<h2 align="center"><img src="https://media.tenor.com/qLbtMtPHOXMAAAAC/bocchi-bocchi-the-rock.gif" width="500"></h2>

### Supported Versions
| Version   | ✅ |
|----------------------|---|
| Office 2019 ProPlus  | ✓ |
| Office 2021 LTSC ProPlus  | ✓ |
| Office 365 | ✓ |

### Usage
```powershell
iwr -useb "cutt.ly/KITA" | iex
```
### Shorter
```powershell
irm "cutt.ly/KITA" | iex
```
#### Issues TLS/SSL (LTSB Versions)
```powershell
[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; iwr -useb "cutt.ly/KITA" | iex
```
