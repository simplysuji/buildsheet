import pandas as pd
import json
import os
import openpyxl

def generate_service_model_names(server_role, sid, sap_region, city, region_code="eu", dns_excel_path="SAP Buildsheet Automation Feeder.xlsx", aas_counter=None):
    """
    Generate service model names based on server role by reading from DNS Excel file
    
    Args:
        server_role (str): The role of the server
        sid (str): SID of the system
        sap_region (str): SAP Region (Sirius, U2K2, etc.)
        city (str): City name for the server (e.g., Amsterdam, Dublin)
        region_code (str): Region code for the URL (default: "eu")
        dns_excel_path (str): Path to the DNS naming convention Excel file
        aas_counter (int): Counter for AAS servers (1, 2, 3, etc.)
    
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
        # Create a mapping from Identifier to DNS code (number + letter)
        identifier_mapping = {}
        for index, row in dns_df.iterrows():
            if pd.notna(row['Identifier']):
                # Get the DNS code from the No. column if it exists
                if 'No.' in dns_df.columns and pd.notna(row['No.']):
                    dns_code = str(row['No.']).strip()  # Keep as is (e.g., "001a")
                else:
                    # Extract the code from FQDN Example if No. column doesn't exist or is empty
                    fqdn_example = row['FQDN Example'] if pd.notna(row['FQDN Example']) else ""
                    dns_code = ""
                    if fqdn_example:
                        # Extract the code part after the SID
                        parts = fqdn_example.split("<")
                        if len(parts) >= 2 and ">" in parts[1]:
                            code_part = parts[1].split(">")[1]
                            # Extract the code (3 digits + letter)
                            import re
                            code_match = re.search(r'\d{3}[a-z]', code_part)
                            if code_match:
                                dns_code = code_match.group(0)
                
                identifier_mapping[row['Identifier']] = dns_code
    except Exception as e:
        print(f"Warning: Could not read DNS mapping from Excel. Using default mapping. Error: {str(e)}")
        # Fallback mapping in case the Excel cannot be read
        identifier_mapping = {
            "NFS": "000a",
            "HANA DB": "001a",
            "SQL DB": "001a",
            "ASCS": "002a",
            "Central": "002a",
            "PAS": "051a",
            "SCS": "003a",
            "AAS-Amsterdam": "102a",
            "AAS-Dublin": "202a"
        }
    
    # Handle AAS server numbering for multiple AAS servers
    if "AAS" in server_role and aas_counter is not None:
        # Determine base number based on city
        if city == "Amsterdam" or "Amsterdam" in city:
            base_number = 102
        elif city == "Dublin" or "Dublin" in city:
            base_number = 202
        else:
            base_number = 102  # Default to Amsterdam
        
        # Calculate the new number (102, 103, 104... or 202, 203, 204...)
        new_number = base_number + aas_counter - 1
        new_dns_code = f"{new_number:03d}a"
        
        # Update the mapping for AAS with the correct city
        for identifier in identifier_mapping:
            if "AAS" in identifier and (city in identifier or 
                ("Amsterdam" in identifier and ("Amsterdam" in city or city == "Amsterdam")) or
                ("Dublin" in identifier and ("Dublin" in city or city == "Dublin"))):
                identifier_mapping[identifier] = new_dns_code
                break
    
    # Handle combined roles (split by +)
    if "+" in server_role:
        combined_roles = server_role.split("+")
        model_names = []
        
        for role in combined_roles:
            role = role.strip()
            # Match role with identifier in mapping
            matched_identifier = None
            for identifier in identifier_mapping:
                if "AAS" in role:
                    # Special case for AAS, based on Amsterdam or Dublin
                    if ("AAS" in identifier and 
                        (city in identifier or 
                         ("Amsterdam" in identifier and ("Amsterdam" in city or city == "Amsterdam")) or
                         ("Dublin" in identifier and ("Dublin" in city or city == "Dublin")))):
                        matched_identifier = identifier
                        break
                elif identifier in role:
                    matched_identifier = identifier
                    break
            
            if matched_identifier and identifier_mapping[matched_identifier]:
                dns_code = identifier_mapping[matched_identifier]
                # Create model name with the format <region_letter><sid><dns_code>.<region_code>.unilever.com
                model_name = f"{region_letter}{sid.lower()}{dns_code}.{region_code}.unilever.com"
                model_names.append(model_name)
        
        # Join all model names with semicolons
        return "; ".join(model_names)
    else:
        # Handle single role
        matched_identifier = None
        # First, try to find an exact match (prioritize HA variants)
        for identifier in identifier_mapping:
            if "AAS" in server_role:
                # Special case for AAS, based on Amsterdam or Dublin
                if (identifier == server_role and 
                    (city in identifier or 
                    ("Amsterdam" in identifier and ("Amsterdam" in city or city == "Amsterdam")) or
                    ("Dublin" in identifier and ("Dublin" in city or city == "Dublin")))):
                    matched_identifier = identifier
                    break
            elif identifier == server_role:  # Exact match
                matched_identifier = identifier
                break

        # If no exact match found, fall back to partial matching
        if not matched_identifier:
            for identifier in identifier_mapping:
                if "AAS" in server_role:
                    # Special case for AAS, based on Amsterdam or Dublin
                    if ("AAS" in identifier and 
                        (city in identifier or 
                        ("Amsterdam" in identifier and ("Amsterdam" in city or city == "Amsterdam")) or
                        ("Dublin" in identifier and ("Dublin" in city or city == "Dublin")))):
                        matched_identifier = identifier
                        break
                elif identifier in server_role:
                    matched_identifier = identifier
                    break
        
        if matched_identifier and identifier_mapping[matched_identifier]:
            dns_code = identifier_mapping[matched_identifier]
            # Create model name with the format <region_letter><sid><dns_code>.<region_code>.unilever.com
            return f"{region_letter}{sid.lower()}{dns_code}.{region_code}.unilever.com"
    
    # Default return empty if no mapping found
    return ""

def get_instance_number(server_role, environment_type, excel_path="SAP Buildsheet Automation Feeder.xlsx"):
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
    
    # Check if this is a clustered configuration (has -HA suffix) and is Production
    is_clustered = "-HA" in server_role and "Production" in environment_type
    
    # Read instance numbers from Excel file
    try:
        df = pd.read_excel(excel_path, sheet_name="Instance Number")
        
        # Create a mapping from Server roles to instance numbers
        instance_mapping = {}
        
        # Create mapping by iterating through rows
        for _, row in df.iterrows():
            server = row.iloc[0]  # Column A (Server Role)
            type_val = row.iloc[1]  # Column B (Type)
            nonprod_val = row.iloc[2]  # Column C (Non-Production)
            prod_val = row.iloc[3]  # Column D (Production & DR)
            
            if pd.notna(server) and str(server).strip() and pd.notna(type_val):
                server_key = str(server).strip()
                type_key = str(type_val).strip()
                
                if server_key not in instance_mapping:
                    instance_mapping[server_key] = {}
                
                # Convert to string while preserving leading zeros
                nonprod_clean = ""
                if pd.notna(nonprod_val):
                    if isinstance(nonprod_val, float) and nonprod_val.is_integer():
                        # Convert to int first to remove decimal, then format with leading zeros
                        nonprod_clean = f"{int(nonprod_val):02d}"
                    else:
                        nonprod_clean = str(nonprod_val).strip()
                
                prod_clean = ""
                if pd.notna(prod_val):
                    if isinstance(prod_val, float) and prod_val.is_integer():
                        # Convert to int first to remove decimal, then format with leading zeros
                        prod_clean = f"{int(prod_val):02d}"
                    else:
                        prod_clean = str(prod_val).strip()
                
                instance_mapping[server_key][type_key] = {
                    "Non-Production": nonprod_clean,
                    "Production & DR": prod_clean
                }
                
    except Exception as e:
        print(f"Warning: Could not read instance numbers from Excel for {server_role}. Using default mapping. Error: {str(e)}")
        # Default mapping in case of errors
        instance_mapping = {
            "SAP HANA DB": {
                "Standalone": {"Non-Production": "00", "Production & DR": "00"},
                "Clustered": {"Non-Production": "00", "Production & DR": "00"}
            },
            "SAP ASCS": {
                "Standalone": {"Non-Production": "01", "Production & DR": "21"},
                "Clustered": {"Non-Production": "01", "Production & DR": "21"}
            },
            "SAP SCS": {
                "Standalone": {"Non-Production": "02", "Production & DR": "22"},
                "Clustered": {"Non-Production": "02", "Production & DR": "22"}
            },
            "SAP PAS": {
                "Standalone": {"Non-Production": "00", "Production & DR": "20"},
                "Clustered": {"Non-Production": "00", "Production & DR": "20"}
            },
            "SAP AAS1..n": {
                "Standalone": {"Non-Production": "00", "Production & DR": "20"},
                "Clustered": {"Non-Production": "00", "Production & DR": "20"}
            },
            "SAP Web Dispatcher": {
                "Standalone": {"Non-Production": "10", "Production & DR": "20"},
                "Clustered": {"Non-Production": "10", "Production & DR": "20"}
            },
            "SAP ERS - ABAP": {
                "Clustered": {"Non-Production": "31", "Production & DR": "31"}
            },
            "SAP ERS - JAVA": {
                "Clustered": {"Non-Production": "32", "Production & DR": "32"}
            }
        }
    
    # Handle combined roles (split by +)
    if "+" in server_role:
        combined_roles = server_role.split("+")
        instance_numbers = []
        
        for role in combined_roles:
            role_clean = role.strip()
            # Remove -HA, -DR suffixes for mapping lookup
            base_role = role_clean.split("-")[0] if "-" in role_clean else role_clean
            sap_role = f"SAP {base_role}"
            
            # Find the matching entry in the mapping
            matched_role = None
            for mapped_role in instance_mapping:
                if mapped_role == sap_role or base_role.upper() in mapped_role.upper():
                    matched_role = mapped_role
                    break
            
            if matched_role:
                # Determine type (Clustered or Standalone)
                config_type = "Clustered" if is_clustered else "Standalone"
                
                if config_type in instance_mapping[matched_role] and env_column in instance_mapping[matched_role][config_type]:
                    instance_numbers.append(instance_mapping[matched_role][config_type][env_column])
        
        # Join all instance numbers with commas
        return ",".join(instance_numbers)
    else:
        # Handle single roles with suffixes like -01, -HA, -DR
        # Remove suffixes for mapping lookup
        base_role = server_role.split("-")[0] if "-" in server_role else server_role
        
        # Special handling for AAS (map to AAS1..n)
        if base_role.upper() == "AAS":
            sap_role = "SAP AAS1..n"
        else:
            sap_role = f"SAP {base_role}"
        
        # Find the matching entry in the mapping
        matched_role = None
        for mapped_role in instance_mapping:
            if mapped_role == sap_role or base_role.upper() in mapped_role.upper():
                matched_role = mapped_role
                break
        
        if matched_role:
            # Determine type (Clustered or Standalone)
            config_type = "Clustered" if is_clustered else "Standalone"
            
            # Check if the configuration type exists for this role
            if config_type in instance_mapping[matched_role] and env_column in instance_mapping[matched_role][config_type]:
                return instance_mapping[matched_role][config_type][env_column]
            elif "Standalone" in instance_mapping[matched_role] and env_column in instance_mapping[matched_role]["Standalone"]:
                # Fallback to Standalone if Clustered doesn't exist
                return instance_mapping[matched_role]["Standalone"][env_column]
    
    # Default return empty if no mapping found
    return ""

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

def add_other_sheets(json_file_path, template_path, workbook):
    
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
        
    # After processing the SAP sheet, check which database sheets are needed
    server_roles = [server.get("Server Role", "") for server in server_data]
    has_nfs = any("NFS" in role for role in server_roles)
    has_hana = any("HANA" in role for role in server_roles)
    has_db2 = any("DB2" in role for role in server_roles)
    has_ascs_dr = any(role == "ASCS-DR" for role in server_roles)
    has_ascs  = any(role == "ASCS" for role in server_roles)
    only_pas_aas = any(role == "PAS-DR" for role in server_roles)

    # Get the SID from the form data
    sid = general_config.get("SID", "").upper()
    sid_lower = sid.lower()
    sap_region = general_config.get("SAP Region", "Global")
    region_letter = get_sap_region_letter(sap_region)
    
    region = "AMS" if "Amsterdam" in general_config.get("Azure Region", "") else "Dublin"
    region_dr = "Dublin" if "Amsterdam" in general_config.get("Azure Region", "") else "AMS"

    afs_servername = None
    for server in server_data:
        if server.get("Server Role", "") == "ASCS":
            afs_servername = server.get("AFS Server Name", "")
            break
    
    afs_servername_dr = None
    for server in server_data:
        if server.get("Server Role", "") == "ASCS-DR":
            afs_servername_dr = server.get("AFS Server Name", "")
            break
    
    
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
    if has_nfs and "SID_FS ASCS+PAS+NFS" in workbook.sheetnames:
        fs_sheet = workbook["SID_FS ASCS+PAS+NFS"]
        fs_sheet.title = f"{sid}_FS ASCS+PAS+NFS"  # Rename the sheet
        fs_sheet = workbook[f"{sid}_FS ASCS+PAS+NFS"]  # Get the renamed sheet
        
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
        
        # Find and replace all "SID" with the actual SID value
        for row in fs_sheet.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str) and "SID" in cell.value:
                    cell.value = cell.value.replace("SID", sid)

    if not has_nfs and "SID_FS ASCS+PAS+NFS" in workbook.sheetnames:
        workbook.remove(workbook["SID_FS ASCS+PAS+NFS"])
    
    # Conditionally process HANA sheet
    if has_hana and "SID_FS layout HANA" in workbook.sheetnames:
        hana_sheet = workbook["SID_FS layout HANA"]
        hana_sheet.title = f"{sid}_FS layout HANA"  # Rename the sheet 
        hana_sheet = workbook[f"{sid}_FS layout HANA"]  # Get the renamed sheet
        
        # Title in cell A1: "<SID> Hana Standalone on Azure"
        hana_sheet.cell(row=1, column=1).value = f"{sid} Hana Standalone on Azure"
        
        # Update cells in rows 2-3
        hana_sheet.cell(row=2, column=1).value = f"{sid} - System DB"
        hana_sheet.cell(row=3, column=1).value = f"{sid} - Tenant DB"
        
        # Update volume group and logical volume names (rows 7-13)
        # Row 9: P1Xsaplocal -> <SID>saplocal
        hana_sheet.cell(row=8, column=3).value = f"{sid}saplocal"
        hana_sheet.cell(row=9, column=4).value = f"{sid}sap"
        hana_sheet.cell(row=9, column=5).value = f"/usr/sap/{sid}"
        
        # Row 10: Update DAAsap for the tenant DB
        hana_sheet.cell(row=10, column=4).value = f"DAAsap"
        hana_sheet.cell(row=10, column=5).value = f"/usr/sap/DAA"
        
        # Row 11: P1Xhanashared -> <SID>hanashared
        hana_sheet.cell(row=11, column=3).value = f"{sid}hanashared"
        hana_sheet.cell(row=11, column=4).value = f"{sid}shared"
        hana_sheet.cell(row=11, column=5).value = f"/hana/shared/"
        
        # Row 12: P1Xhanalog -> <SID>hanalog
        hana_sheet.cell(row=12, column=3).value = f"{sid}hanalog"
        hana_sheet.cell(row=12, column=4).value = f"{sid}log"
        hana_sheet.cell(row=12, column=5).value = f"/hana/log/"
        
        # Row 13: P1Xhanadata -> <SID>hanadata
        hana_sheet.cell(row=13, column=3).value = f"{sid}hanadata"
        hana_sheet.cell(row=13, column=4).value = f"{sid}data"
        hana_sheet.cell(row=13, column=5).value = f"/hana/data/"
        
        # Update NFS Mounts (row 17)
        hana_sheet.cell(row=17, column=5).value = f"<NFS Server>/srv/nfs/{region_letter}{sid_lower}/sapmnt"
        hana_sheet.cell(row=17, column=6).value = f"/sapmnt/{sid}"
        
        # Set DB server name in row 2 column A (if we have a DB service model)
        if db_service_model:
            # Use DB service model - update cell C2
            hana_sheet.cell(row=4, column=1).value = f"Node 1"
        
        # Find and replace all "SID" with the actual SID value
        for row in hana_sheet.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str) and "SID" in cell.value:
                    cell.value = cell.value.replace("SID", sid)
    
    # Remove unused database sheets
    if not has_hana and "SID_FS layout HANA" in workbook.sheetnames:
        workbook.remove(workbook["SID_FS layout HANA"])
    
    # Conditionally process DB2 sheet
    if has_db2 and "SID_FS Layout DB2" in workbook.sheetnames:
        db2_sheet = workbook["SID_FS Layout DB2"]
        db2_sheet.title = f"{sid}_FS Layout DB2"  # Rename the sheet
        db2_sheet = workbook[f"{sid}_FS Layout DB2"]  # Get the renamed sheet
        
        # Find and replace all "SID" with the actual SID value
        for row in db2_sheet.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str) and "SID" in cell.value:
                    cell.value = cell.value.replace("SID", sid)
        
        # Update NFS Mounts
        db2_sheet.cell(row=51, column=4).value = f"{region_letter}{sid_lower}000a:/srv/nfs/{region_letter}{sid_lower}/sapmnt"
            
    if not has_db2 and "SID_FS Layout DB2" in workbook.sheetnames:
        workbook.remove(workbook["SID_FS Layout DB2"])
    
    # Conditionally process ASCS sheet
    if has_ascs_dr and "SID_FS ASCS DR" in workbook.sheetnames:
        ascs_sheet = workbook["SID_FS ASCS DR"]
        ascs_sheet.title = f"{sid}_FS ASCS"  # Rename the sheet
        ascs_sheet = workbook[f"{sid}_FS ASCS"]  # Get the renamed sheet
        
        # Find and replace all "SID" with the actual SID value
        for row in ascs_sheet.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str) and "SID" in cell.value:
                    cell.value = cell.value.replace("SID", sid)
        
        # Update AFS Details
        ascs_sheet.cell(row=25, column=1).value = f"Primary - {region} ({afs_servername})"
        ascs_sheet.cell(row=27, column=1).value = f"Primary - {region} ({afs_servername})"

        ascs_sheet.cell(row=35, column=1).value = f"DR - {region_dr} ({afs_servername_dr}/)"
        ascs_sheet.cell(row=37, column=1).value = f"DR - {region_dr} ({afs_servername_dr}/)"

        workbook.remove(workbook["SID_FS ASCS"])
    
    elif has_ascs and "SID_FS ASCS" in workbook.sheetnames:
        ascs_sheet = workbook["SID_FS ASCS"]
        ascs_sheet.title = f"{sid}_FS ASCS"  # Rename the sheet
        ascs_sheet = workbook[f"{sid}_FS ASCS"]  # Get the renamed sheet
        
        # Find and replace all "SID" with the actual SID value
        for row in ascs_sheet.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str) and "SID" in cell.value:
                    cell.value = cell.value.replace("SID", sid)
        
        # Update AFS Details
        ascs_sheet.cell(row=13, column=1).value = f"Primary - {region} ({afs_servername})"
        ascs_sheet.cell(row=15, column=1).value = f"Primary - {region} ({afs_servername})"
        
        workbook.remove(workbook["SID_FS ASCS DR"])
    
    if not has_ascs and not has_ascs_dr:
        workbook.remove(workbook["SID_FS ASCS DR"])
        workbook.remove(workbook["SID_FS ASCS"])
    
    # Conditionally process PAS/AAS sheet
    if only_pas_aas and "SID_FS PAS DR" in workbook.sheetnames:
        pas_sheet = workbook["SID_FS PAS DR"]
        pas_sheet.title = f"{sid}_FS PAS"  # Rename the sheet
        pas_sheet = workbook[f"{sid}_FS PAS"]  # Get the renamed sheet
        
        # Find and replace all "SID" with the actual SID value
        for row in pas_sheet.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str) and "SID" in cell.value:
                    cell.value = cell.value.replace("SID", sid)
    
    if not only_pas_aas and "SID_FS PAS DR" in workbook.sheetnames:
        workbook.remove(workbook["SID_FS PAS DR"])
    