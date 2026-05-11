# Cybersecurity-Forensic-Investigation-Workbook
An automated Python tool for generating Tier-3 SOC, DFIR, and Penetration Testing workbooks.

## Features & Capabilities

### 🛡️ Defensive Operations (Blue Team / DFIR)
* **NIST SP 800-61 Alignment:** Full integration of the Incident Response Lifecycle (Preparation through Recovery).
* **Cryptographic Chain of Custody:** Automated SHA-256 hashing fields to ensure forensic evidence integrity and non-repudiation.
* **Chronological Master Timeline:** A synchronized event log that tracks both "Event Time" and "Analyst Discovery Time" to prevent time-skew errors.
* **SIEM Integration Ready:** Structured to accept and correlate logs from Security Onion (Zeek/Suricata), pfSense (Snort), and Windows Event Logs.


### ⚔️ Offensive Operations (Red Team / PenTest)
* **Multi-Framework Mapping:** Standardized reporting for **PTES** (Network), **OWASP Top 10** (Web), and **OSSTMM** (Operational Security).
* **Adversary Emulation:** Mapping of custom exploits and payloads directly to **MITRE ATT&CK®** techniques.
* **Tactical Operations Journal:** A high-fidelity "live-fire" log for tracking terminal commands, payload parameters, and raw tool outputs.


### 🛠️ Technical Architecture
* **Automated Report Generation:** Python-driven generation of professional, formatted Excel Dossiers using `pandas` and `openpyxl`.
* **Dynamic Pivot Matrix:** A layout designed for quick "Gap Analysis," allowing researchers to compare tool performance side-by-side.
* **Modular Design:** Easily extensible data structures to add new cybersecurity frameworks or custom lab tools.
#
## 📁 Workbook Structure
The generated `.xlsx` report contains the following specialized modules:
- **Tactical Activity Log:** 🔵 Blue Team command-and-control logging.
- **PTES Matrix:** 🟢 Network penetration testing milestones.
- **OWASP Tracker:** 🟠 Web application vulnerability mapping.
- **NIST IR Lifecycle:** 🔴 Incident response and forensic timeline.
- **Evidence Vault:** ⚫ SHA-256 hash verification and Chain of Custody.

## Technical Architecture
![Cybersecurity SOC Architecture Diagram](https://raw.githubusercontent.com/Eb0nyR0se/Cybersecurity-Forensic-Investigation-Workbook/main/cybersecurity_architecture.png.png)

