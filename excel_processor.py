import pandas as pd
import json
import os
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill

def generate_service_model_names(server_role, sid, sap_region, region_code="eu", dns_excel_path="SAP Service Model Names 10.11.xlsx"):
    """
    Generate service model names based on server role by reading from DNS Excel file
    
    Args:
        server_role (str): The role of the server
        sid (str): SID of the system
        sap_region (str): SAP Region (Sirius, U2K2, etc.)
        region_code (str): Region code for the URL (default: "eu")
        dns_excel_path (str): Path to the DNS naming convention Excel file
    
    Returns:
        str: Generated service model name(s)
    """
    # Map SAP Region to region code letter
    region_mapping = {
        "Sirius": "e",
        "U2K2": "a",
        "Cordillera": "t",
        "Global": "o",
        "POC/Model Env": "x"
    }
    
    # Get the region letter from mapping or default to "x"
    region_letter = region_mapping.get(sap_region, "x")
    
    # Read DNS mapping from Excel file
    try:
        dns_df = pd.read_excel(dns_excel_path, sheet_name="DNS -Azure", header=1)
        # Create a mapping from Identifier to DNS number
        identifier_mapping = {}
        for index, row in dns_df.iterrows():
            if pd.notna(row['Identifier']):
                # Get the 3-digit number from the No. column if it exists
                if 'No.' in dns_df.columns and pd.notna(row['No.']):
                    number = str(row['No.']).zfill(3)  # Ensure 3 digits with leading zeros
                else:
                    # Extract the number from FQDN Example if No. column doesn't exist or is empty
                    fqdn_example = row['FQDN Example'] if pd.notna(row['FQDN Example']) else ""
                    number = ""
                    if fqdn_example:
                        # Extract the number part after the SID
                        parts = fqdn_example.split("<")
                        if len(parts) >= 2 and ">" in parts[1]:
                            number_part = parts[1].split(">")[1]
                            # Extract just the number (3 digits)
                            import re
                            number_match = re.search(r'\d{3}', number_part)
                            if number_match:
                                number = number_match.group(0)
                
                identifier_mapping[row['Identifier']] = number
    except Exception as e:
        print(f"Warning: Could not read DNS mapping from Excel. Using default mapping. Error: {str(e)}")
        # Fallback mapping in case the Excel cannot be read
        identifier_mapping = {
            "NFS": "051",
            "HANA DB": "001",
            "SQL DB": "001",
            "ASCS": "002",
            "Central": "002",
            "PAS": "000",
            "SCS": "003"
        }
    
    # Handle combined roles (split by +)
    if "+" in server_role:
        combined_roles = server_role.split("+")
        model_names = []
        
        for role in combined_roles:
            role = role.strip()
            # Match role with identifier in mapping
            matched_identifier = None
            for identifier in identifier_mapping:
                if identifier in role:
                    matched_identifier = identifier
                    break
            
            if matched_identifier and identifier_mapping[matched_identifier]:
                number = identifier_mapping[matched_identifier]
                # Create model name with the format <sid><region_letter><number>a.<region_code>.unilever.com
                model_name = f"{region_letter}{sid.lower()}{number}.{region_code}.unilever.com"
                model_names.append(model_name)
        
        # Join all model names with semicolons
        return "; ".join(model_names)
    else:
        # Handle single role
        matched_identifier = None
        for identifier in identifier_mapping:
            if identifier in server_role:
                matched_identifier = identifier
                break
        
        if matched_identifier and identifier_mapping[matched_identifier]:
            number = identifier_mapping[matched_identifier]
            # Create model name with the format <sid><region_letter><number>a.<region_code>.unilever.com
            return f"{region_letter}{sid.lower()}{number}.{region_code}.unilever.com"
    
    # Default return empty if no mapping found
    return ""

def get_instance_number(server_role, environment_type, excel_path="SAP Service Model Names 10.11.xlsx"):
    """
    Get instance number for server role based on environment type by reading from Excel file
    
    Args:
        server_role (str): The server role
        environment_type (str): Environment type (Production or Non-Production)
        excel_path (str): Path to the Excel file containing instance numbers
        
    Returns:
        str: Instance number for the given server role and environment
    """
    import pandas as pd
    
    # Determine environment column to use
    env_column = "Production & DR" if "Production" in environment_type else "Non-Production"
    
    # Read instance numbers from Excel file
    try:
        df = pd.read_excel(excel_path, sheet_name="Instance Number")
        
        # Create a mapping from Server roles to instance numbers
        instance_mapping = {}
        
        # Find the relevant columns
        server_col = df.columns.get_loc("Server Role")
        prod_col = df.columns.get_loc("Production & DR")
        nonprod_col = df.columns.get_loc("Non-Production")
        
        # Create mapping
        for _, row in df.iterrows():
            server = row[server_col]
            prod_value = row[prod_col]
            nonprod_value = row[nonprod_col]
            
            if pd.notna(server) and str(server).strip():
                instance_mapping[str(server).strip()] = {
                    "Production & DR": str(prod_value).strip() if pd.notna(prod_value) else "",
                    "Non-Production": str(nonprod_value).strip() if pd.notna(nonprod_value) else ""
                }
    except Exception as e:
        print(f"Warning: Could not read instance numbers from Excel. Using default mapping. Error: {str(e)}")
        # Default mapping in case of errors
        instance_mapping = {
            "HANA DB": {"Non-Production": "00", "Production & DR": "00"},
            "ASCS": {"Non-Production": "01", "Production & DR": "21"},
            "SCS": {"Non-Production": "02", "Production & DR": "22"},
            "PAS": {"Non-Production": "00", "Production & DR": "20"},
            "AAS1..n": {"Non-Production": "00", "Production & DR": "20"},
            "Web Dispatcher": {"Non-Production": "10", "Production & DR": "20"}
        }
    
    # Handle combined roles (split by +)
    if "+" in server_role:
        combined_roles = server_role.split("+")
        instance_numbers = []
        
        for role in combined_roles:
            role = role.strip()
            # Find the matching entry in the mapping
            matched_role = None
            for mapped_role in instance_mapping:
                if mapped_role in role:
                    matched_role = mapped_role
                    break
            
            if matched_role and env_column in instance_mapping[matched_role]:
                instance_numbers.append(instance_mapping[matched_role][env_column])
        
        # Join all instance numbers with commas
        return ",".join(instance_numbers)
    else:
        # Handle suffix cases like -01, -HA, -DR
        base_role = server_role.split("-")[0] if "-" in server_role else server_role
        
        # Find the matching entry in the mapping
        matched_role = None
        for mapped_role in instance_mapping:
            if mapped_role in base_role:
                matched_role = mapped_role
                break
        
        if matched_role and env_column in instance_mapping[matched_role]:
            return instance_mapping[matched_role][env_column]
    
    # Default return empty if no mapping found
    return ""

# Region letter mapping function
def get_sap_region_letter(sap_region):
    """Get the letter code for a SAP region"""
    region_mapping = {
        "Sirius": "e",
        "U2K2": "a",
        "Cordillera": "t",
        "Global": "o",
        "POC/Model Env": "x"
    }
    return region_mapping.get(sap_region, "x")

def get_environment_code(environment):
    """Get the code for a specific environment"""
    environment_mapping = {
        "Fix Development": "d",
        "Project Development": "d", 
        "Fix Quality": "q",
        "Project Quality": "q",
        "Training": "q",
        "Fix Performance": "q",
        "Project performance": "q",
        "Project UAT": "q",
        "Fix Regression": "u",
        "Sandbox": "s",
        "Production": "p"
    }
    return environment_mapping.get(environment, "d") # default to 'd' if not found

def process_sap_data_to_excel(json_file_path, template_path, output_path=None):
    """
    Process SAP form data from JSON and fill in the Excel template.
    
    Args:
        json_file_path (str): Path to the JSON file with form data
        template_path (str): Path to the Excel template
        output_path (str, optional): Path to save the output Excel file. Defaults to None.
    
    Returns:
        str: Path to the saved Excel file
    """
    # Check if files exist
    if not os.path.exists(json_file_path):
        raise FileNotFoundError(f"JSON file not found: {json_file_path}")
    
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Excel template not found: {template_path}")
    
    # Load JSON data
    with open(json_file_path, 'r') as file:
        form_data = json.load(file)
    
    # Extract general config and server data
    general_config = form_data.get("general_config", {})
    server_data = form_data.get("server_data", [])
    
    # Load Excel template
    workbook = openpyxl.load_workbook(template_path)
    sheet = workbook["SAP"]  # Access the SAP sheet
    
    # Start filling data from row 12
    start_row = 12
    
    # For each server, fill a row in the Excel sheet
    for i, server in enumerate(server_data):
        row = start_row + i
        
        # Fill in values using the correct column numbers from the provided list
        # Column 1: Azure Region
        sheet.cell(row=row, column=1).value = general_config.get("Azure Region", "")
        
        # Column 2: Environment
        sheet.cell(row=row, column=2).value = general_config.get("Environment", "")
        
        # Column 3: SID
        sid = general_config.get("SID", "")
        sheet.cell(row=row, column=3).value = sid
        
        # Column 4: Instance Number
        server_role = server.get("Server Role", "")
        environment = general_config.get("Environment", "")
        instance_number = get_instance_number(server_role, environment)
        sheet.cell(row=row, column=4).value = instance_number
        
        # Column 5: Server Role
        sheet.cell(row=row, column=5).value = server.get("Server Role", "")
        
        # Column 6: Service Criticality
        sheet.cell(row=row, column=6).value = server.get("Service Criticality", "")
        
        # Column 7: ITSG ID
        sheet.cell(row=row, column=7).value = general_config.get("ITSG ID", "")
        
        # Column 8: OS Version
        sheet.cell(row=row, column=8).value = server.get("OS Version", "")
        
        # Column 9: Licensing (default to "UL: BYOL")
        sheet.cell(row=row, column=9).value = "UL: BYOL"
        
        # Column 10: Azure Instance Type
        sheet.cell(row=row, column=10).value = server.get("Instance Type", "")
        
        # Column 11: Memory / CPU
        sheet.cell(row=row, column=11).value = server.get("Memory/CPU", "")
        
        # Column 12: Service Model name
        service_model_name = generate_service_model_names(
            server.get("Server Role", ""), 
            general_config.get("SID", ""),
            general_config.get("SAP Region", "Sirius"),
        )
        sheet.cell(row=row, column=12).value = service_model_name
        
        # Column 13: A Record / CNAME
        sheet.cell(row=row, column=13).value = general_config.get("Record Type", "")
        
        # Column 14: VNET (default to "STS Vnet")
        sheet.cell(row=row, column=14).value = "STS Vnet"
        
        # Column 15: Subnet/Zone
        sheet.cell(row=row, column=15).value = general_config.get("Subnet/Zone", "")
        
        # Column 16: Subnet (leave blank)
        sheet.cell(row=row, column=16).value = ""
        
        # Column 17: Subnet Group (leave blank)
        sheet.cell(row=row, column=17).value = ""
        
        # Column 18: Azure Subscription
        sheet.cell(row=row, column=18).value = general_config.get("Azure Subscription", "")
        
        # Column 19: Azure Resource Group Name (generate based on pattern)
        region_code = general_config.get("Azure Region Code", "").lower()
        environment_code = get_environment_code(general_config.get("Environment", ""))
        itsg_id = general_config.get("ITSG ID", "")
        sid = general_config.get("SID", "").upper()
        subscription = general_config.get("Azure Subscription", "")
        subscription_number = subscription.split("-")[1].split(" ")[0]
        
        # Format: <REGION_CODES>-sp01-s-<ITSG>-<SID>-IS01-rg
        resource_group_name = f"{region_code}-sp{subscription_number}-{environment_code}-{itsg_id}-{sid}-IS01-rg"
        sheet.cell(row=row, column=19).value = resource_group_name
        
        # Column 20: Cluster
        sheet.cell(row=row, column=20).value = general_config.get("Cluster", "")
        
        # Column 21: Proximity Placement Group (updated format)
        server_role = server.get("Server Role", "")
        if any(role in server_role for role in ["PAS", "ASCS", "SCS"]):
            sap_region = general_config.get("SAP Region", "Sirius")
            region_letter = get_sap_region_letter(sap_region)
            az_zone = general_config.get("AZ Selection", "")

            # Format: <azure_region_code>-sp01-<SAPregion>-<ITSG>-<SID>-az<x>-ppg
            ppg = f"{region_code}-sp{subscription_number}-{region_letter}-{itsg_id}-{sid}-az{az_zone}-ppg"
            sheet.cell(row=row, column=21).value = ppg
        else:
            sheet.cell(row=row, column=21).value = ""
        
        # Column 22: Availability Set (updated format)
        if server.get("Availability Set", "").lower() == "yes":
            sap_region = general_config.get("SAP Region", "Sirius")
            region_letter = get_sap_region_letter(sap_region)
            az_zone = general_config.get("AZ Selection", "")
            
            # Format: <azure_region_code>-sp01-<SAPregion>-<ITSG>-<SID>-az<x>-as
            availability_set = f"{region_code}-sp{subscription_number}-{region_letter}-{itsg_id}-{sid}-az{az_zone}-as"
            sheet.cell(row=row, column=22).value = availability_set
        else:
            sheet.cell(row=row, column=22).value = ""
        
        # Column 23: Azure Availability Zone
        sheet.cell(row=row, column=23).value = general_config.get("AZ Selection", "")
        
        # Column 24: Accelerated Networking (default to "Yes")
        sheet.cell(row=row, column=24).value = "Yes"
        
        # Column 25: On Demand/Reservation
        sheet.cell(row=row, column=25).value = server.get("Reservation Type", "")
        
        # Column 26: Reservation Term
        sheet.cell(row=row, column=26).value = server.get("Reservation Term", "")
        
        # Column 27: OptInOptOut
        sheet.cell(row=row, column=27).value = server.get("OptInOptOut", "")
        
        # Column 28: Park My Cloud Schedule
        sheet.cell(row=row, column=28).value = server.get("Park My Cloud Schedule", "")
        
        # Column 29: Park My cloud team name and Member
        sheet.cell(row=row, column=29).value = server.get("Park My cloud team name and Member", "")
        
        # Columns 30-43: Type and Qty pairs (leaving blank)
        for col in range(30, 44):
            sheet.cell(row=row, column=col).value = ""
        
        # Column 44: Outbound Internet Access Required
        sheet.cell(row=row, column=44).value = server.get("Outbound Internet Access Required", "")
        
        # Column 45: Additional Requirements / Comments (leave blank)
        sheet.cell(row=row, column=45).value = ""
        
        # Column 46: Build RFC (leave blank)
        sheet.cell(row=row, column=46).value = ""
        
        # Column 47: iSCSI details (leave blank)
        sheet.cell(row=row, column=47).value = ""
        
        # Column 48: Instance Name (leave blank)
        sheet.cell(row=row, column=48).value = ""
        
        # Column 49: Private IP Address (leave blank)
        sheet.cell(row=row, column=49).value = ""
        
        # Column 50: TimeZone (default to "CET")=====
        sheet.cell(row=row, column=50).value = server.get("Timezone", "CET")
    
    # Process the sheets after handling the SAP sheet
    # Get the SID from the form data
    sid = general_config.get("SID", "").upper()
    sid_lower = sid.lower()
    sap_region = general_config.get("SAP Region", "Sirius")
    region_letter = get_sap_region_letter(sap_region)

    # Collect service model names and identify NFS and DB servers
    service_model_names = []
    nfs_service_model = None
    db_service_model = None

    for server in server_data:
        role = server.get("Server Role", "")
        server_model_name = generate_service_model_names(
            role, 
            sid,
            sap_region,
            general_config.get("Azure Region Code", "eu")[:2]
        )
        
        if server_model_name:
            service_model_names.append(server_model_name)
            
            # Check for NFS role
            if "NFS" in role:
                nfs_service_model = server_model_name
            
            # Check for DB role (HANA DB or DB2 DB)
            if "DB" in role:
                db_service_model = server_model_name

    # Rename the sheets to use the SID provided by the user
    if "SID_FS" in workbook.sheetnames:
        fs_sheet = workbook["SID_FS"]
        fs_sheet.title = f"{sid}_FS"  # Rename the sheet
        fs_sheet = workbook[f"{sid}_FS"]  # Get the renamed sheet
        
        # Process SID_FS sheet
        if service_model_names:
            # A2: Combined server names with '/' separator
            fs_sheet.cell(row=2, column=1).value = "/".join(service_model_names)
        
        # B5: <SID>saplocal
        fs_sheet.cell(row=5, column=2).value = f"{sid}saplocal"
        
        # B10 and C10: <SID>NFS
        fs_sheet.cell(row=10, column=2).value = f"{sid}NFS"
        fs_sheet.cell(row=10, column=3).value = f"{sid}NFS"
        
        # C6: <SID>sap
        fs_sheet.cell(row=6, column=3).value = f"{sid}sap"
        
        # C8: <SID>ascs
        fs_sheet.cell(row=8, column=3).value = f"{sid}ascs"
        
        # D6: /usr/sap/<SID>
        fs_sheet.cell(row=6, column=4).value = f"/usr/sap/{sid}"
        
        # D8: /usr/sap/<SID>/ASCS01
        fs_sheet.cell(row=8, column=4).value = f"/usr/sap/{sid}/ASCS01"
        
        # D10: /srv/nfs/{region_letter}{sid_lower}
        fs_sheet.cell(row=10, column=4).value = f"/srv/nfs/{region_letter}{sid_lower}"
        fs_sheet.cell(row=10, column=5).value = f"/srv/nfs/{region_letter}{sid_lower}002"
        
        # Process the NFS directory structure (rows 15-19)
        # Row 15: <NFS Server>/srv/nfs/{region_letter}{sid_lower}/interface 
        fs_sheet.cell(row=15, column=1).value = f"<NFS Server>/srv/nfs/{region_letter}{sid_lower}/interface"
        fs_sheet.cell(row=15, column=4).value = f"/interface/sap{sid.lower()}"
        
        # Row 16: <NFS Server>/srv/nfs/{region_letter}{sid_lower}/sapmnt
        fs_sheet.cell(row=16, column=1).value = f"<NFS Server>/srv/nfs/{region_letter}{sid_lower}/sapmnt"
        fs_sheet.cell(row=16, column=4).value = f"/sapmnt/{sid}"
        
        # Row 17: <NFS Server>/srv/nfs/{region_letter}{sid_lower}/{region_letter}{sid_lower}002
        fs_sheet.cell(row=17, column=1).value = f"<NFS Server>/srv/nfs/{region_letter}{sid_lower}/{region_letter}{sid_lower}002"
        fs_sheet.cell(row=17, column=4).value = f"/app/{region_letter}{sid_lower}002"
        
        # Row 18: <NFS Server>/srv/nfs/{region_letter}{sid_lower}/trans
        fs_sheet.cell(row=18, column=1).value = f"<NFS Server>/srv/nfs/{region_letter}{sid_lower}/trans"
        fs_sheet.cell(row=18, column=4).value = f"/usr/sap/trans"

    # Process SID_DB sheet if it exists
    if "SID_DB" in workbook.sheetnames:
        db_sheet = workbook["SID_DB"]
        db_sheet.title = f"{sid}_DB"  # Rename the sheet
        db_sheet = workbook[f"{sid}_DB"]  # Get the renamed sheet
        
        # Title in cell A1: "<SID> Hana Standalone on Azure"
        db_sheet.cell(row=1, column=1).value = f"{sid} Hana Standalone on Azure"
        
        # Update cells in rows 2-3
        db_sheet.cell(row=2, column=1).value = f"{sid} - System DB"
        db_sheet.cell(row=3, column=1).value = f"{sid} - Tenant DB"
        
        # Update volume group and logical volume names (rows 7-13)
        # Row 9: P1Xsaplocal -> <SID>saplocal
        db_sheet.cell(row=8, column=3).value = f"{sid}saplocal"
        db_sheet.cell(row=9, column=4).value = f"{sid}sap"
        db_sheet.cell(row=9, column=5).value = f"/usr/sap/{sid}"
        
        # Row 10: Update DAAsap for the tenant DB
        db_sheet.cell(row=10, column=4).value = f"DAAsap"
        db_sheet.cell(row=10, column=5).value = f"/usr/sap/DAA"
        
        # Row 11: P1Xhanashared -> <SID>hanashared
        db_sheet.cell(row=11, column=3).value = f"{sid}hanashared"
        db_sheet.cell(row=11, column=4).value = f"{sid}shared"
        db_sheet.cell(row=11, column=5).value = f"/hana/shared/"
        
        # Row 12: P1Xhanalog -> <SID>hanalog
        db_sheet.cell(row=12, column=3).value = f"{sid}hanalog"
        db_sheet.cell(row=12, column=4).value = f"{sid}log"
        db_sheet.cell(row=12, column=5).value = f"/hana/log/"
        
        # Row 13: P1Xhanadata -> <SID>hanadata
        db_sheet.cell(row=13, column=3).value = f"{sid}hanadata"
        db_sheet.cell(row=13, column=4).value = f"{sid}data"
        db_sheet.cell(row=13, column=5).value = f"/hana/data/"
        
        # Update NFS Mounts (row 17)
        db_sheet.cell(row=17, column=5).value = f"<NFS Server>/srv/nfs/{region_letter}{sid_lower}/sapmnt"
        db_sheet.cell(row=17, column=6).value = f"/sapmnt/{sid}"
        
        # Set DB server name in row 2 column A (if we have a DB service model)
        if db_service_model:
            # Use DB service model - update cell C2
            db_sheet.cell(row=4, column=1).value = f"Node 1"
    
    
    # Save the workbook
    if output_path is None:
        output_path = f"Filled_{os.path.basename(template_path)}"
    
    workbook.save(output_path)
    return output_path

if __name__ == "__main__":
    # This block only executes when the script is run directly
    # It won't run when the script is imported
    # import sys
    
    # # Check if command line arguments are provided
    # if len(sys.argv) < 3:
    #     print("Usage: python excel_processor.py <json_file> <template_file> [output_file]")
    #     sys.exit(1)
    
    # json_file = sys.argv[1]
    # template_file = sys.argv[2]
    # output_file = sys.argv[3] if len(sys.argv) > 3 else None
    json_file = "sap_form_data.json"
    template_file = "Template.xlsx"
    output_file = "Filled_SAP_Template.xlsx"
    
    try:
        output_path = process_sap_data_to_excel(json_file, template_file, output_file)
        print(f"Processing completed successfully. Output saved to: {output_path}")
    except Exception as e:
        print(f"Error: {str(e)}")
        # sys.exit(1)