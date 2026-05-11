import pandas as pd
import os
from datetime import datetime

def generate_purple_team_workbook():
    # Set the specific filename requested
    filename = "Cybersecurity_Purple_Team_Investigation_Workbook.xlsx"
    today_str = datetime.now().strftime('%Y-%m-%d')
    
    # 1. BLUE TEAM: Network Defense & Asset Tracking
    blue_team_data = {
        'Date': [today_str],
        'Analyst': ['First Last Name'],
        'Source_MAC': [''], 
        'OUI_Vendor': [''],      
        'Detected_IP': ['192.168.x.x'],
        'TTL_Value': [''],       
        'Connection_Point': [''], 
        'Action_Taken': ['Logged'], 
        'Alert_ID': ['']         
    }

    # 2. RED TEAM: Tactical Engagement
    red_team_data = {
        'Date': [today_str],
        'Analyst': ['First Last Name'],
        'Target_Host': ['192.168.x.x'],       
        'Method': ['Nmap -sV'], 
        'Protocol_Port': [''],   
        'Mitre_Technique': ['T1595'], 
        'Result': ['Filtered'],  
        'Notes': ['']
    }

    # 3. TACTICAL LOG: The Raw Command Stream
    tactical_log_columns = [
        'Date', 
        'Timestamp', 
        'Source_IP',       # 192.168.x.x
        'Destination_IP',  # 192.168.x.x
        'Command_Executed', 
        'Raw_Output_Excerpt', # For logging nmap retransmission errors
        'Analyst_Observations'
    ]

    # 4. THREAT INTEL: MAC/IP Reputation
    intel_data = {
        'Indicator': ['192.168.x.x'],       
        'Type': ['IP/MAC Address'],
        'First_Seen': [today_str],
        'Last_Seen': [today_str],
        'Threat_Score': ['1'],   
        'Context': ['']          
    }

    # Generate DataFrames
    df_blue = pd.DataFrame(blue_team_data)
    df_red = pd.DataFrame(red_team_data)
    df_tactical = pd.DataFrame(columns=tactical_log_columns)
    df_intel = pd.DataFrame(intel_data)

    # Export to Excel
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df_blue.to_excel(writer, sheet_name='Blue_Team', index=False)
        df_red.to_excel(writer, sheet_name='Red_Team', index=False)
        df_tactical.to_excel(writer, sheet_name='Tactical_Log', index=False)
        df_intel.to_excel(writer, sheet_name='Threat_Intel', index=False)

    print(f"\n[+] Workbook Initialized: {filename}")
    print(f"[+] Analyst: First Last Name")
    print(f"[+] Default Subnet: 192.168.x.x")

if __name__ == "__main__":
    generate_purple_team_workbook()
