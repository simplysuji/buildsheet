import streamlit as st
import pandas as pd
import json
import os
from excel_processor import process_non_prod_data_to_excel, process_prod_data_to_excel

def save_form_data(form_data, filename="sap_form_data.json"):
    """
    Save the form data to a JSON file
    
    Args:
        form_data (dict): The form data to save
        filename (str): The name of the JSON file
        
    Returns:
        str: The path to the saved file
    """
    # Create the file path
    file_path = os.path.join(os.getcwd(), filename)
    
    # Save the data to a JSON file
    with open(file_path, 'w') as f:
        json.dump(form_data, f, indent=4)
    
    return file_path

def create_download_link(excel_file_path):
    """
    Creates a download button for the Excel file
    
    Args:
        excel_file_path (str): Path to the Excel file
        
    Returns:
        None: Displays the download button in the app
    """
    with open(excel_file_path, "rb") as file:
        st.download_button(
            label="📥 Download Excel File",
            data=file,
            file_name=os.path.basename(excel_file_path),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# Function to get region code based on Azure region
def get_region_code(azure_region):
    """
    Get the region code based on the selected Azure region
    
    Args:
        azure_region (str): The selected Azure region
        
    Returns:
        str: The corresponding region code
    """
    if "Amsterdam" in azure_region or "NLWE" in azure_region:
        return "bnlwe"
    elif "Dublin" in azure_region or "IENO" in azure_region:
        return "bieno"
    else:
        return "bnlwe"  # Default fallback

# Add this function after the imports and before the constants
def load_vm_sku_data():
    """
    Load VM SKU data from the Excel file
    
    Returns:
        dict: Dictionary mapping VM Type to VM Size
    """
    try:
        # Read the SKU Name sheet from the Excel file
        df = pd.read_excel("SAP Buildsheet Automation Feeder.xlsx", sheet_name="SKU Name")
        
        # Create a dictionary mapping VM Type to VM Size
        vm_mapping = {}
        for _, row in df.iterrows():
            vm_type = row['VM Type']
            vm_size = row['VM Size']
            if pd.notna(vm_type) and pd.notna(vm_size):
                vm_mapping[vm_type] = vm_size
        
        return vm_mapping
    except Exception as e:
        st.error(f"Error loading VM SKU data: {str(e)}")
        return {}

def get_server_role_versions(server_role, form_options):
    """
    Get available versions for a specific server role
    
    Args:
        server_role (str): The selected server role
        form_options (dict): Form options containing version mappings
        
    Returns:
        list: List of available versions for the server role
    """
    versions_map = form_options.get('SERVER_ROLES_VERSIONS_MAP', {})
    return versions_map.get(server_role, [])

def load_form_field_options():
    """
    Load form field options from the Excel file
    
    Returns:
        dict: Dictionary containing all form field options
    """
    try:
        # Read the Form Fields sheet from the Excel file
        df = pd.read_excel("SAP Buildsheet Automation Feeder.xlsx", sheet_name="Form Fields")
        
        # Create a dictionary to store all options
        form_options = {}
        
        # Process each column
        for column in df.columns:
            # Get all non-null values from the column and convert to list
            options = df[column].dropna().tolist()
            # Remove any empty strings
            options = [str(option).strip() for option in options if str(option).strip()]
            form_options[column] = options
        
        # Create server role versions mapping by reading both columns together
        server_roles_versions = {}
        if 'SERVER_ROLES' in df.columns and 'SERVER_ROLES_VERSIONS' in df.columns:
            # Iterate through the DataFrame rows to map roles to versions
            for _, row in df.iterrows():
                role = row.get('SERVER_ROLES')
                versions = row.get('SERVER_ROLES_VERSIONS')
                
                # Only process if role exists and is not null
                if pd.notna(role) and str(role).strip():
                    role = str(role).strip()
                    
                    # Check if versions exist and are not null/empty
                    if pd.notna(versions) and str(versions).strip():
                        # Split comma-separated values and clean them
                        version_list = [v.strip() for v in str(versions).split(',') if v.strip()]
                        if version_list:
                            server_roles_versions[role] = version_list
        
        form_options['SERVER_ROLES_VERSIONS_MAP'] = server_roles_versions
        
        return form_options
    except Exception as e:
        st.error(f"Error loading form field options: {str(e)}")
        # Return default options if Excel loading fails
        return get_default_form_options()

def get_default_form_options():
    """
    Fallback function with default form options
    
    Returns:
        dict: Dictionary containing default form field options
    """
    return {
        'ENVIRONMENTS': ["Fix Development", "Fix Quality", "Fix Regression", "Fix Performance", 
                        "Project performance", "Project Development", "Project Quality", "Training", 
                        "Sandbox", "Project UAT", "Production"],
        'AZURE_SUBSCRIPTIONS': [
            "SAP Technical Services-01 (Global)", "SAP Technical Services-02 (Sirius)",
            "SAP Technical Services-03 (U2K2)", "SAP Technical Services-04 (Cordillera)",
            "SAP Technical Services-05 (Fusion)", "SAP Technical Services-98 (Model Environment)"
        ],
        'SAP_REGIONS': ["Global", "Sirius", "U2K2", "Cordillera", "Fusion", "POC/Model Env"],
        'SERVICE_CRITICALITY': ["SC 1", "SC 2", "SC 3", "SC 4"],
        'AZURE_REGIONS': [
            "Azure: Northern Europe (Dublin) (IENO)",
            "Azure: Western Europe (Amsterdam) (NLWE)",
        ],
        'SERVER_ROLES': [
            "AAS", "ASCS", "ASCS+NFS", "ASCS+PAS", "ASCS+PAS+NFS", "ASCS-HA", "SCS+PAS+NFS", 
            "CS+NFS", "CS+NFS+PAS", "DB2 DB", "DB2 DB-HA", "HANA DB", "HANA DB-HA", "iSCSI SBD", "SCS", "SCS+NFS", "SCS+PAS",
            "PAS", "Web Dispatcher", "Web Dispatcher-HA", "Maxdb", "CS", "Optimizers", "IQ roles",
            "PAS-DR", "AAS-DR", "ASCS-DR", "HANA DB-DR", "SCS-DR", "SCS-HA", "iSCSI SBD-DR", "Web Dispatcher-DR"
        ],
        'OS_VERSIONS': [
            "RHEL 7.9 for SAP", "RHEL 8.10 SAP", "SLES 12 SP3", "SLES 12 SP4", 
            "SLES 12 SP5", "SLES 15 SP1", "SLES 15 SP2", "Windows 2016", "Windows 2019", "Windows 2022", "Windows 2025"
        ],
        'PARK_SCHEDULES': [
            "Weekdays-12 hours Snooze(5pm IST to 5am IST) and Weekends Off",
            "Weekdays-12 hours Snooze(6pm IST to 6am IST) and Weekends Off",
            "Weekdays-12 hours Snooze(7pm IST to 7am IST) and Weekends Off",
            "Weekdays-12 hours Snooze(8pm IST to 8am IST) and Weekends Off",
            "Weekdays-12 hours Snooze(9pm IST to 9am IST) and Weekends Off",
            "Weekdays-12 hours Snooze(10pm IST to 10am IST) and Weekends Off",
            "Weekdays-12 hours Snooze(11pm IST to 11am IST) and Weekends Off"
        ],
        'TIMEZONE': ["IST", "UTC", "CET", "CEST", "GMT", "CST", "PST", "EST", "BST"],
        'RECORD_TYPES': ["A Record", "CNAME"]
    }

# Add this function to get VM size based on VM type
def get_vm_size(vm_type, vm_mapping):
    """
    Get VM size based on VM type
    
    Args:
        vm_type (str): The selected VM type
        vm_mapping (dict): Dictionary mapping VM Type to VM Size
        
    Returns:
        str: The corresponding VM size
    """
    return vm_mapping.get(vm_type, "Not Available")

# Function to update number of servers
def update_num_servers(tab_key):
    st.session_state[f'num_servers_{tab_key}'] = st.session_state[f'num_servers_input_{tab_key}']

# Function to check if server role contains PAS
def contains_pas(server_role):
    return "PAS" in server_role.upper() or "AAS" in server_role.upper() or "Optimizer" in server_role

# Function to check if server role contains ASCS
def contains_ascs(server_role):
    return "ASCS" == server_role.upper() or "ASCS-DR" == server_role.upper()

def requires_cluster(server_role):
    """
    Check if the server role requires cluster configuration
    
    Args:
        server_role (str): The server role
        
    Returns:
        bool: True if cluster option is applicable for this role
    """
    cluster_roles = ["ASCS", "HANA DB", "DB2 DB", "Maxdb"]
    return any(role in server_role for role in cluster_roles)


def load_az_zone_data():
    """
    Load AZ zone data from the Excel file
    
    Returns:
        dict: Dictionary mapping (subscription, region) to (primary_zone, ha_zone)
    """
    try:
        # Read the AZ Zone sheet from the Excel file
        df = pd.read_excel("SAP Buildsheet Automation Feeder.xlsx", sheet_name="AZ Zone")
        
        # Create a dictionary mapping (subscription, region) to zones
        az_mapping = {}
        for _, row in df.iterrows():
            subscription = row['Subscription']
            region = row['Region']
            primary_zone = str(row['Primary Zone'])
            ha_zone = str(row['HA Zone'])
            
            if pd.notna(subscription) and pd.notna(region):
                # Create key using subscription and region
                key = (subscription, region)
                az_mapping[key] = {
                    'primary_zone': primary_zone,
                    'ha_zone': ha_zone
                }
        
        return az_mapping
    except Exception as e:
        st.error(f"Error loading AZ zone data: {str(e)}")
        return {}

def get_az_zones(azure_subscription, azure_region, az_mapping):
    """
    Get primary and HA AZ zones based on subscription and region
    
    Args:
        azure_subscription (str): The selected Azure subscription
        azure_region (str): The selected Azure region
        az_mapping (dict): AZ zone mapping data
        
    Returns:
        tuple: (primary_zone, ha_zone)
    """
    # Extract region name from azure_region (Amsterdam or Dublin)
    if "Amsterdam" in azure_region or "NLWE" in azure_region:
        region = "Amsterdam"
    elif "Dublin" in azure_region or "IENO" in azure_region:
        region = "Dublin"
    else:
        region = "Amsterdam"  # Default fallback
    
    # Look up the zones
    key = (azure_subscription, region)
    if key in az_mapping:
        return az_mapping[key]['primary_zone'], az_mapping[key]['ha_zone']
    else:
        # Default fallback if not found
        return "1", "2"

def remove_dr_server(tab_key, server_index):
    """Remove a DR server from the enabled list"""
    if server_index in st.session_state[f'dr_servers_enabled_{tab_key}']:
        st.session_state[f'dr_servers_enabled_{tab_key}'].remove(server_index)
        
        # Clear session state for this specific DR server
        keys_to_clear = [key for key in st.session_state.keys() 
                        if key.startswith('dr_') and key.endswith(f'_{tab_key}_{server_index}')]
        for key in keys_to_clear:
            del st.session_state[key]

def render_file_management_tab():
    """
    Render the file management tab for updating Excel files
    """
    st.header("📁 File Management")
    st.markdown("Update the configuration files used by the application")
    
    # Initialize session state for download tracking
    if 'feeder_downloaded' not in st.session_state:
        st.session_state.feeder_downloaded = False
    if 'template_downloaded' not in st.session_state:
        st.session_state.template_downloaded = False
    
    # Create two columns for the cards
    col1, col2 = st.columns(2)
    
    # SAP Buildsheet Automation Feeder.xlsx Card
    with col1:
        with st.container():
            st.markdown("""
                <div style="
                    border: 2px solid #e0e0e0;
                    border-radius: 10px;
                    padding: 20px;
                    margin: 10px 0;
                    background-color: #f8f9fa;
                    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                ">
                    <h3 style="color: #2c3e50; text-align: center; margin-bottom: 20px;">
                        🗂️ SAP Buildsheet Automation Feeder
                    </h3>
                </div>
            """, unsafe_allow_html=True)
            
            st.markdown("**Purpose:** Contains form field options, VM SKU mappings, and other configuration data")
            
            # Check if file exists
            feeder_file_path = "SAP Buildsheet Automation Feeder.xlsx"
            if os.path.exists(feeder_file_path):
                file_size = os.path.getsize(feeder_file_path)
                file_size_mb = round(file_size / (1024 * 1024), 2)
                modification_time = os.path.getmtime(feeder_file_path)
                mod_date = pd.to_datetime(modification_time, unit='s').tz_localize('UTC').tz_convert('Asia/Kolkata').strftime('%d-%m-%Y %H:%M:%S IST')
                
                st.info(f"📊 **Current File Info:**\n- Size: {file_size_mb} MB\n- Last Modified: {mod_date}")
                
                # Download button
                with open(feeder_file_path, "rb") as file:
                    if st.download_button(
                        label="📥 Download Current Version",
                        data=file,
                        file_name="SAP Buildsheet Automation Feeder.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_feeder",
                        help="Download the current version to update it",
                        use_container_width=True
                    ):
                        st.session_state.feeder_downloaded = True
                        st.success("✅ File downloaded! You can now upload an updated version.")
                        # st.rerun()
            else:
                st.error("❌ Current file not found!")
                st.session_state.feeder_downloaded = True  # Allow upload if no current file exists
            
            # Upload section
            st.markdown("---")
            
            if st.session_state.feeder_downloaded:
                st.markdown("**📤 Upload Updated Version:**")
                uploaded_feeder = st.file_uploader(
                    "Choose updated SAP Buildsheet Automation Feeder file",
                    type=['xlsx'],
                    key="upload_feeder",
                    help="Upload the updated Excel file with your changes"
                )
                
                if uploaded_feeder is not None:
                    # Validate the uploaded file
                    try:
                        # Try to read the file to validate it's a proper Excel file
                        test_df = pd.read_excel(uploaded_feeder, sheet_name=None)
                        
                        # Check for required sheets
                        required_sheets = ["DNS -Azure", "Instance Number", "Form Fields", "SKU Name", "AZ Zone"]
                        missing_sheets = [sheet for sheet in required_sheets if sheet not in test_df.keys()]
                        
                        if missing_sheets:
                            st.error(f"❌ Missing required sheets: {', '.join(missing_sheets)}")
                        else:
                            # Show file info
                            file_size_mb = round(uploaded_feeder.size / (1024 * 1024), 2)
                            st.success(f"✅ Valid Excel file uploaded! Size: {file_size_mb} MB")
                            
                            # Save the file
                            if st.button("💾 Save Updated File", key="save_feeder", use_container_width=True):
                                try:
                                    # Create backup of existing file if it exists
                                    if os.path.exists(feeder_file_path):
                                        backup_path = f"backup_{pd.Timestamp.now().strftime('%d%m%Y_%H%M%S')}_SAP Buildsheet Automation Feeder.xlsx"
                                        os.rename(feeder_file_path, backup_path)
                                        st.info(f"🔄 Existing file backed up as: {backup_path}")
                                    
                                    # Save the new file
                                    with open(feeder_file_path, "wb") as f:
                                        f.write(uploaded_feeder.getbuffer())
                                    
                                    st.success("🎉 SAP Buildsheet Automation Feeder.xlsx updated successfully!")
                                    st.balloons()
                                    
                                    # Reset download state
                                    st.session_state.feeder_downloaded = False
                                    st.rerun()
                                    
                                except Exception as e:
                                    st.error(f"❌ Error saving file: {str(e)}")
                                    
                    except Exception as e:
                        st.error(f"❌ Invalid Excel file: {str(e)}")
            else:
                st.warning("⚠️ Please download the current version first before uploading a new one.")
    
    # Template.xlsx Card
    with col2:
        with st.container():
            st.markdown("""
                <div style="
                    border: 2px solid #e0e0e0;
                    border-radius: 10px;
                    padding: 20px;
                    margin: 10px 0;
                    background-color: #f8f9fa;
                    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                ">
                    <h3 style="color: #2c3e50; text-align: center; margin-bottom: 20px;">
                        📋 Template File
                    </h3>
                </div>
            """, unsafe_allow_html=True)
            
            st.markdown("**Purpose:** Excel template used for generating the final buildsheet output")
            
            # Check if file exists
            template_file_path = "Template.xlsx"
            if os.path.exists(template_file_path):
                file_size = os.path.getsize(template_file_path)
                file_size_mb = round(file_size / (1024 * 1024), 2)
                modification_time = os.path.getmtime(template_file_path)
                mod_date = pd.to_datetime(modification_time, unit='s').tz_localize('UTC').tz_convert('Asia/Kolkata').strftime('%d-%m-%Y %H:%M:%S IST')

                st.info(f"📊 **Current File Info:**\n- Size: {file_size_mb} MB\n- Last Modified: {mod_date}")
                
                # Download button
                with open(template_file_path, "rb") as file:
                    if st.download_button(
                        label="📥 Download Current Version",
                        data=file,
                        file_name="Template.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_template",
                        help="Download the current version to update it",
                        use_container_width=True
                    ):
                        st.session_state.template_downloaded = True
                        st.success("✅ File downloaded! You can now upload an updated version.")
                        st.rerun()
            else:
                st.error("❌ Current file not found!")
                st.session_state.template_downloaded = True  # Allow upload if no current file exists
            
            # Upload section
            st.markdown("---")
            
            if st.session_state.template_downloaded:
                st.markdown("**📤 Upload Updated Version:**")
                uploaded_template = st.file_uploader(
                    "Choose updated Template file",
                    type=['xlsx'],
                    key="upload_template",
                    help="Upload the updated Excel template file"
                )
                
                if uploaded_template is not None:
                    # Validate the uploaded file
                    try:
                        # Try to read the file to validate it's a proper Excel file
                        test_df = pd.read_excel(uploaded_template, sheet_name=None)
                        
                        # Check for required sheets
                        required_sheets = ["SAP"]
                        missing_sheets = [sheet for sheet in required_sheets if sheet not in test_df.keys()]

                        if missing_sheets:
                            st.error(f"❌ Missing required sheet(s): {', '.join(missing_sheets)}")
                            return  # Stop processing if required sheets are missing

                        # Show file info
                        file_size_mb = round(uploaded_template.size / (1024 * 1024), 2)
                        st.success(f"✅ Valid Excel file uploaded! Size: {file_size_mb} MB")
                        
                        # Save the file
                        if st.button("💾 Save Updated File", key="save_template", use_container_width=True):
                            try:
                                # Create backup of existing file if it exists
                                if os.path.exists(template_file_path):
                                    backup_path = f"backup_{pd.Timestamp.now().strftime('%d%m%Y_%H%M%S')}_Template.xlsx"
                                    os.rename(template_file_path, backup_path)
                                    st.info(f"🔄 Existing file backed up as: {backup_path}")
                                
                                # Save the new file
                                with open(template_file_path, "wb") as f:
                                    f.write(uploaded_template.getbuffer())
                                
                                st.success("🎉 Template.xlsx updated successfully!")
                                st.balloons()
                                
                                # Reset download state
                                st.session_state.template_downloaded = False
                                st.rerun()
                                
                            except Exception as e:
                                st.error(f"❌ Error saving file: {str(e)}")
                                
                    except Exception as e:
                        st.error(f"❌ Invalid Excel file: {str(e)}")
            else:
                st.warning("⚠️ Please download the current version first before uploading a new one.")
    
    # Instructions section
    st.markdown("---")
    with st.expander("ℹ️ Instructions", expanded=False):
        st.markdown("""
        **How to update files:**
        
        1. **Download Current Version**: Click the download button to get the current file
        2. **Update the File**: Make your changes to the downloaded file using Excel
        3. **Upload Updated Version**: Once downloaded, the upload button will become active
        4. **Save**: Click the save button to replace the current file
        
        **Important Notes:**
        - 🔄 You must download the current version before you can upload a new one
        - ✅ DO not delete any existing sheets from the Excel files
        - 🚀 Changes will take effect immediately after saving
        """)

def get_suggested_dr_role(primary_server_role, form_options):
    """
    Dynamically find the corresponding DR role for a primary server role
    
    Args:
        primary_server_role (str): The primary server role
        form_options (dict): Form options containing SERVER_ROLES list
        
    Returns:
        str: The suggested DR role if found, empty string otherwise
    """
    # Get the list of all server roles
    all_server_roles = form_options.get('SERVER_ROLES', [])
    
    # Filter to get only DR roles
    dr_server_roles = [role for role in all_server_roles if '-DR' in role]
    
    # If the primary role already contains DR, return empty
    if "DR" in primary_server_role:
        return ""
    
    # Try to find exact match: "{primary_server_role}-DR"
    suggested_dr_role = f"{primary_server_role}-DR"
    if suggested_dr_role in dr_server_roles:
        return suggested_dr_role
    
    # If exact match not found, try some common pattern matching
    # This handles cases where the primary role might have slight variations
    for dr_role in dr_server_roles:
        # Remove "-DR" from the DR role to get the base name
        base_dr_role = dr_role.replace("-DR", "")
        
        # Check if the primary role matches the base DR role
        if primary_server_role.lower() == base_dr_role.lower():
            return dr_role
        
        # Check if primary role is contained in the base DR role or vice versa
        # This handles cases like "ASCS+PAS" -> "ASCS-DR" (if available)
        if primary_server_role in base_dr_role or base_dr_role in primary_server_role:
            return dr_role
    
    # If no match found, return empty string
    return ""


def render_dr_server_config(tab_key, vm_sku_mapping, server_index, region_code, az_zone_mapping):
    """
    Render DR server configuration form
    
    Args:
        tab_key (str): Tab key for session state management
        vm_sku_mapping (dict): VM SKU mapping data
        server_index (int): Index of the server
    
    Returns:
        dict: DR server configuration data
    """
    
    INSTANCE_TYPES = list(vm_sku_mapping.keys()) if vm_sku_mapping else [
        "D8asv4", "E8asv4", "E16_v3", "E16as_v4", "E16ds_v4", "E16s_v3", "E2_v3", 
        "E20as_v4", "E20ds_v4", "E20s_v3", "E2as_v4", "E2ds_v4", "E2s_v3", "E32_v3",
        "E32as_v4", "E32ds_v4", "E32s_v3", "E4_v3", "E48as_v4", "E48ds_v4", "E48s_v3",
        "E4as_v4", "E4ds_v4", "E4s_v3"
    ]
    
    # Load form options and get values from Excel
    form_options = load_form_field_options()
    dr_server_roles = [role for role in form_options['SERVER_ROLES'] if '-DR' in role]
    OS_VERSIONS = form_options['OS_VERSIONS']
    PARK_SCHEDULES = form_options['PARK_SCHEDULES']
    RECORD_TYPES = form_options['RECORD_TYPES']
    
    # Get the primary server role to create DR equivalent
    primary_server_role = st.session_state.get(f"server_role_{tab_key}_{server_index}", "") 

    suggested_dr_role = get_suggested_dr_role(primary_server_role, form_options)
    
    # Set default index for DR role
    default_dr_index = 0
    if suggested_dr_role and suggested_dr_role in dr_server_roles:
        default_dr_index = dr_server_roles.index(suggested_dr_role)
    
    
    with st.expander(f"🔄 DR Server {server_index+1} Configuration", expanded=True):
        # Add remove button at the top
        col_header, col_remove = st.columns([13, 1])
        with col_remove:
            if st.button("🗑️ Remove", 
                        key=f"remove_dr_{tab_key}_{server_index}",
                        help=f"Remove DR Server {server_index+1}",
                        type="secondary"):
                
                remove_dr_server(tab_key, server_index)
                st.rerun()
                
        # First row: Server Role, A Record/CNAME, Service Criticality
        col1, col2 = st.columns(2)
        with col1:
            dr_server_role = st.selectbox(
                "DR Server Role", 
                dr_server_roles, 
                index=default_dr_index,
                key=f"dr_server_role_{tab_key}_{server_index}",
                help=f"DR counterpart for primary server: {primary_server_role}"
            )
        with col2:
            dr_record_type = st.selectbox(
                "A Record / CNAME", 
                RECORD_TYPES, 
                key=f"dr_record_type_{tab_key}_{server_index}"
            )
        
        # Second row: OS Version, Availability Set (if applicable)
        col1, col2, col3 = st.columns(3)
        with col1:
            dr_os_version = st.selectbox(
                "OS Version", 
                OS_VERSIONS, 
                key=f"dr_os_version_{tab_key}_{server_index}"
            )
        with col2:
            # Show Availability Set option if DR server role contains PAS
            dr_availability_set = "No"  # Default value
            if contains_pas(dr_server_role):
                dr_availability_set = st.selectbox(
                    "Availability Set", 
                    ["Yes", "No"], 
                    key=f"dr_availability_set_{tab_key}_{server_index}",
                    help="Required for PAS DR servers for high availability"
                )
            else:
                st.text_input(
                    "Availability Set", 
                    value="N/A", 
                    disabled=True,
                    key=f"dr_availability_set_display_{tab_key}_{server_index}"
                )
        
        with col3:
            dr_az_zones = ["1", "2", "3"]
            dr_aas_servers_so_far = 0
            for idx in st.session_state[f'dr_servers_enabled_{tab_key}']:
                if idx < server_index:  # Count AAS servers before current one
                    prev_dr_role = st.session_state.get(f"dr_server_role_{tab_key}_{idx}", "")
                    if "AAS" in prev_dr_role:
                        dr_aas_servers_so_far += 1

            # Get DR region HA/Primary zones (opposite of primary region)
            dr_region = "Dublin" if "bnlwe" in region_code.lower() else "Amsterdam"
            current_azure_subscription = st.session_state.get(f"azure_subscription_{tab_key}", "")
            dr_suggested_primary, dr_suggested_ha = get_az_zones(current_azure_subscription, dr_region, az_zone_mapping)

            # Determine zone for current DR server
            if "AAS" in dr_server_role:
                dr_aas_servers_so_far += 1
                # Odd DR AAS servers get HA zone, even get Primary zone (same pattern as primary)
                selected_dr_zone = dr_suggested_ha if dr_aas_servers_so_far % 2 == 1 else dr_suggested_primary
                dr_zone_help = f"Load balanced zone for DR AAS server #{dr_aas_servers_so_far}"
            else:
                # Non-AAS servers use suggested primary zone (same as primary logic)
                selected_dr_zone = dr_suggested_primary
                dr_zone_help = f"Suggested Primary Zone: {dr_suggested_primary} (based on subscription and region)"

            selected_dr_zone_index = dr_az_zones.index(selected_dr_zone) if selected_dr_zone in dr_az_zones else 0

            dr_az_selection = st.selectbox(
                "AZ Selection - Zone", 
                dr_az_zones,
                index=selected_dr_zone_index,
                key=f"dr_az_selection_{tab_key}_{server_index}",
                help=dr_zone_help
            )

        
        # Show AFS Servername option if server role contains PAS
        dr_afs_needed = "NA"  # Default value
        if contains_ascs(dr_server_role):
            dr_region_code = "bieno" if "bnlwe" in region_code.lower() else "bnlwe"

            dr_afs_needed = st.text_input(f'AFS Server Name ({dr_region_code+"stgunileversp*****"})', 
                                        key=f"dr_afs_needed_{tab_key}_{server_index}",
                                        help="Required for ASCS servers")
            
            if dr_afs_needed and not dr_afs_needed.isdigit():
                st.error("Please enter **only digits** to be filled in (*****)")

            dr_afs = dr_region_code+"stgunileversp"+dr_afs_needed
        
        # Third row: Instance Type, Memory/CPU
        col1, col2 = st.columns(2)
        with col1:
            dr_instance_type = st.selectbox(
                "Azure Instance Type", 
                INSTANCE_TYPES, 
                key=f"dr_instance_type_{tab_key}_{server_index}"
            )
        with col2:
            dr_memory_cpu = get_vm_size(dr_instance_type, vm_sku_mapping)
            st.text_input(
                "Memory / CPU", 
                value=dr_memory_cpu,
                disabled=True,
                key=f"dr_memory_cpu_{tab_key}_{server_index}",
                help="This value is automatically populated based on the selected Azure Instance Type"
            )
        
        # Fourth row: Reservation options
        col1, col2 = st.columns(2)
        with col1:
            dr_reservation_type = st.selectbox(
                "On Demand/Reservation", 
                ["On Demand", "Reservation"], 
                key=f"dr_reservation_type_{tab_key}_{server_index}"
            )
        if dr_reservation_type == "Reservation":
            with col2:
                dr_reservation_term = st.selectbox(
                    "Reservation Term", 
                    ["One Year", "Three Years"],
                    disabled=dr_reservation_type != "Reservation", 
                    key=f"dr_reservation_term_{tab_key}_{server_index}"
                )
        
        # Fifth row: Cloud management
        col1, col2 = st.columns(2)
        with col1:
            dr_opt_in_out = st.selectbox(
                "OptInOptOut", 
                ["In", "Out"], 
                index=1,  # Default to "Out" for production DR
                help="In - Can be parked/managed using ParkMyCloud\nOut - Cannot be parked/Managed using Park My Cloud",
                key=f"dr_opt_in_out_{tab_key}_{server_index}"
            )
        with col2:
            dr_internet_access = st.selectbox(
                "Outbound Internet Access Required", 
                ["Yes", "No"], 
                key=f"dr_internet_access_{tab_key}_{server_index}"
            )
        
        
        # Sixth row: Team information, Internet access, Timezone
        col1, col2, = st.columns(2)
        with col2:
            dr_team_name = st.text_input(
                "Park My cloud team name and Member", 
                disabled=dr_opt_in_out != "In",
                key=f"dr_team_name_{tab_key}_{server_index}"
            )
        with col1:
            dr_park_schedule = st.selectbox(
                "Park My Cloud Schedule", 
                PARK_SCHEDULES, 
                disabled=dr_opt_in_out != "In",
                key=f"dr_park_schedule_{tab_key}_{server_index}"
            )
        
    
    # Return DR server configuration
    dr_server_config = {
        "Server Number": f"DR-{server_index+1}",
        "Server Role": dr_server_role,
        "Record Type": dr_record_type,
        "OS Version": dr_os_version,
        "Availability Set": dr_availability_set if contains_pas(dr_server_role) else "N/A",
        "AFS Server Name": dr_afs if contains_ascs(dr_server_role) else "N/A",
        "AZ Selection": dr_az_selection,
        "Azure Instance Type": dr_instance_type,
        "Memory/CPU": get_vm_size(dr_instance_type, vm_sku_mapping),
        "Reservation Type": dr_reservation_type,
        "Reservation Term": dr_reservation_term if dr_reservation_type == "Reservation" else "N/A",
        "OptInOptOut": dr_opt_in_out,
        "Park My Cloud Schedule": dr_park_schedule if dr_opt_in_out == "In" else "N/A",
        "Park My cloud team name and Member": dr_team_name if dr_opt_in_out == "In" else "N/A",
        "Outbound Internet Access Required": dr_internet_access,
        "Server Type": "DR"  # Add identifier for DR servers
    }
    
    return dr_server_config

# Function to render the form content (shared between tabs with different configurations)
def render_form_content(tab_key, is_production=False):
    """
    Render the form content for either Production or Non-Production
    
    Args:
        tab_key (str): Either 'prod' or 'nonprod'
        is_production (bool): Whether this is the production form
    """
    
    # Initialize session state for this tab
    if f'num_servers_{tab_key}' not in st.session_state:
        st.session_state[f'num_servers_{tab_key}'] = 1
    if f'form_submitted_{tab_key}' not in st.session_state:
        st.session_state[f'form_submitted_{tab_key}'] = False
    if f'excel_file_path_{tab_key}' not in st.session_state:
        st.session_state[f'excel_file_path_{tab_key}'] = None
    if f'json_file_path_{tab_key}' not in st.session_state:
        st.session_state[f'json_file_path_{tab_key}'] = None
    if f'server_data_{tab_key}' not in st.session_state:
        st.session_state[f'server_data_{tab_key}'] = []
    if f'dr_servers_enabled_{tab_key}' not in st.session_state:
        st.session_state[f'dr_servers_enabled_{tab_key}'] = list(range(st.session_state[f'num_servers_{tab_key}']))

    # Load form field options from Excel
    form_options = load_form_field_options()
    az_zone_mapping = load_az_zone_data()
    
    # Define options based on Production vs Non-Production
    if is_production:
        # Production-specific configurations
        ENVIRONMENTS = ["Production"]
        SUBNETS = ["Production STS"]

    else:
        # Non-Production configurations
        ENVIRONMENTS = [env for env in form_options['ENVIRONMENTS'] if env != "Production"]
        SUBNETS = ["Non-Production STS"]
    
    # Use loaded options from Excel
    AZURE_SUBSCRIPTIONS = form_options['AZURE_SUBSCRIPTIONS']
    SAP_REGIONS = form_options['SAP_REGIONS']
    SERVICE_CRITICALITY = form_options['SERVICE_CRITICALITY']
    AZURE_REGIONS = form_options['AZURE_REGIONS']
    SERVER_ROLES = [role for role in form_options['SERVER_ROLES'] if '-DR' not in role and '-HA' not in role]
    OS_VERSIONS = form_options['OS_VERSIONS']
    RECORD_TYPES = form_options['RECORD_TYPES']
    PARK_SCHEDULES = form_options['PARK_SCHEDULES']
    TIMEZONE = form_options['TIMEZONE']

    
    vm_sku_mapping = load_vm_sku_data()
    INSTANCE_TYPES = list(vm_sku_mapping.keys()) if vm_sku_mapping else [
        "D8asv4", "E8asv4", "E16_v3", "E16as_v4", "E16ds_v4", "E16s_v3", "E2_v3", 
        "E20as_v4", "E20ds_v4", "E20s_v3", "E2as_v4", "E2ds_v4", "E2s_v3", "E32_v3",
        "E32as_v4", "E32ds_v4", "E32s_v3", "E4_v3", "E48as_v4", "E48ds_v4", "E48s_v3",
        "E4as_v4", "E4ds_v4", "E4s_v3"
    ]

    # GENERAL CONFIGURATION SECTION
    st.subheader("General Configuration")

    # First Row
    col1, col2, col3 = st.columns(3)
    with col1:
        sap_region = st.selectbox("SAP Region", SAP_REGIONS, key=f"sap_region_{tab_key}")
    with col2:
        azure_region = st.selectbox("Azure Region", AZURE_REGIONS, key=f"azure_region_{tab_key}")
    with col3:
        azure_subscription = st.selectbox("Azure Subscription", AZURE_SUBSCRIPTIONS, key=f"azure_subscription_{tab_key}")

    # Auto-determine region code based on azure_region selection
    region_code = get_region_code(azure_region)

    # Display the auto-selected region code for user confirmation
    # st.info(f"Region Code set to: **{region_code}** (based on {azure_region})")

    # Second Row
    col1, col2, col3 = st.columns(3)
    with col3:
        subnet = st.selectbox("Subnet/Zone", SUBNETS, key=f"subnet_{tab_key}")
    with col2:
        environment = st.selectbox("Environment", ENVIRONMENTS, key=f"environment_{tab_key}")
    with col1:
        sid = st.text_input("SID", key=f"sid_{tab_key}")

    # Third Row
    col1, col2 = st.columns(2)
    with col1:
        itsg_id = st.text_input("ITSG ID", key=f"itsg_id_{tab_key}")
    with col2:
        num_servers = st.number_input("Number of Servers", 
                                    min_value=1, 
                                    value=st.session_state[f'num_servers_{tab_key}'],
                                    key=f"num_servers_input_{tab_key}",
                                    on_change=lambda: update_num_servers(tab_key))

    # Fourth Row
    col1, col2 = st.columns(2)
    with col1:
        service_criticality = st.selectbox("Service Criticality", SERVICE_CRITICALITY, key=f"service_criticality_{tab_key}")
    with col2:
        timezone = st.selectbox(
            "Timezone", 
            TIMEZONE,
            key=f"timezone_{tab_key}"
        )
    
    # Fifth Row - Only for Non-Production
    if not is_production:
        col1, col2 = st.columns(2)
        with col1:
            # Get suggested AZ zones based on subscription and region
            current_azure_subscription = st.session_state.get(f"azure_subscription_{tab_key}", "")
            current_azure_region = st.session_state.get(f"azure_region_{tab_key}", "")
            suggested_primary, suggested_ha = get_az_zones(current_azure_subscription, current_azure_region, az_zone_mapping)
            current_environment = st.session_state.get(f"environment_{tab_key}", "")
            
            if current_environment in ["Fix Development", "Fix Quality", "Project Development", "Project Quality", "Sandbox",]:
            
                # Display AZ zone as read-only (grayed out)
                az_selection = st.text_input(
                    "AZ Selection - Zone", 
                    value=suggested_ha,
                    disabled=True,
                    key=f"az_selection_{tab_key}",
                    help="Auto-assigned based on subscription and region selection"
                )
            
            else:
                # Display AZ zone as read-only (grayed out)
                az_selection = st.text_input(
                    "AZ Selection - Zone", 
                    value=suggested_primary,
                    disabled=True,
                    key=f"az_selection_{tab_key}",
                    help="Auto-assigned based on subscription and region selection"
                )            
            
        with col2:
            record_type = st.selectbox("A Record / CNAME", RECORD_TYPES, key=f"record_type_{tab_key}")
    

    # SERVER CONFIGURATION SECTION
    st.subheader("Primary Region Server Configuration")

    # Create a section for each server
    server_data = []
    dr_server_data = []
    aas_server_count = 0  # Track AAS servers for load balancing
    
    for i in range(st.session_state[f'num_servers_{tab_key}']):
        with st.expander(f"🖥️ Primary Region Server {i+1}", expanded=True):
                       
            # Basic server info
            col1, col2, col3 = st.columns(3)
            with col1:
                server_role = st.selectbox("Server Role", SERVER_ROLES, key=f"server_role_{tab_key}_{i}")

            # Get available versions for the selected server role
            available_versions = get_server_role_versions(server_role, form_options)

            with col2:
                if available_versions:
                    # Show version dropdown if versions are available
                    server_role_version = st.selectbox(
                        "Server Role Version", 
                        available_versions, 
                        key=f"server_role_version_{tab_key}_{i}",
                        help=f"Available versions for {server_role}"
                    )
                else:
                    # Show disabled field if no versions available
                    st.text_input(
                        "Server Role Version", 
                        value="N/A",
                        disabled=True,
                        key=f"server_role_version_display_{tab_key}_{i}",
                        help="No versions available for this server role"
                    )
                    server_role_version = "N/A"

            with col3:            
                # Show Availability Set option if server role contains PAS
                availability_set = "No"  # Default value
                if contains_pas(server_role):
                    availability_set = st.selectbox("Availability Set", ["Yes", "No"], 
                                                key=f"availability_set_{tab_key}_{i}",
                                                help="Required for PAS servers for high availability")
                else:
                    st.text_input(
                        "Availability Set", 
                        value="N/A", 
                        disabled=True,
                        key=f"availability_set_display_{tab_key}_{i}"
                    )
                
            # Show AFS Servername option if server role contains PAS
            afs_needed = "NA"  # Default value
            if contains_ascs(server_role):
                afs_needed = st.text_input(f'AFS Server Name ({region_code+"stgunileversp*****"})',  
                                              key=f"afs_needed_{tab_key}_{i}",
                                              help="Required for ASCS servers")
                if afs_needed and not afs_needed.isdigit():
                    st.error("Please enter **only digits** to be filled in (*****)")
                primary_afs = region_code+"stgunileversp"+afs_needed
                
            # Production-specific fields: A Record/CNAME, AZ Selection, and Cluster
            if is_production:
                col1, col2, col3 = st.columns(3)
                with col1:
                    record_type = st.selectbox("A Record / CNAME", RECORD_TYPES, key=f"record_type_{tab_key}_{i}")
                with col2:
                    
                    # Get suggested AZ zones based on subscription and region
                    current_azure_subscription = st.session_state.get(f"azure_subscription_{tab_key}", "")
                    current_azure_region = st.session_state.get(f"azure_region_{tab_key}", "")
                    suggested_primary, suggested_ha = get_az_zones(current_azure_subscription, current_azure_region, az_zone_mapping)
                    # Set default index for suggested primary zone
                    az_zones = ["1", "2", "3"]
                    # default_primary_index = 0
                    
                    if "AAS" in server_role:
                        aas_server_count += 1
                        # Odd AAS servers get HA zone, even get Primary zone
                        selected_zone = suggested_ha if aas_server_count % 2 == 1 else suggested_primary
                        selected_zone_index = az_zones.index(selected_zone) if selected_zone in az_zones else 0
                    else:
                        # Non-AAS servers use suggested primary zone
                        selected_zone_index = az_zones.index(suggested_primary)
                    
                    az_selection = st.selectbox(
                        "AZ Selection - Zone", 
                        az_zones,
                        index=selected_zone_index,
                        disabled=True,
                        key=f"az_selection_{tab_key}_{i}",
                        help=f"{'Load balanced zone for AAS server #' + str(aas_server_count) if 'AAS' in server_role else 'Suggested Primary Zone: ' + suggested_primary} (based on subscription and region)"
                    )
                with col3:
                    # Show cluster option only for specific server roles
                    if requires_cluster(server_role):
                        cluster = st.selectbox("Cluster", 
                                             ["Yes", "No"],
                                             key=f"cluster_{tab_key}_{i}",
                                             help="Required for Production environments with ASCS, HANA DB, DB2 DB, or Maxdb roles")
                    else:
                        # Show disabled field for other roles
                        cluster = st.text_input(
                            "Cluster", 
                            value="N/A",
                            disabled=True,
                            key=f"cluster_display_{tab_key}_{i}",
                            help="Cluster option not applicable for this server role"
                        )
                        # Set cluster value for data collection
                        cluster = "N/A"
            
            col1, col2, col3 = st.columns(3)
            with col1:
                os_version = st.selectbox("OS Version", OS_VERSIONS, key=f"os_version_{tab_key}_{i}")
            with col2:
                instance_type = st.selectbox("Azure Instance Type", INSTANCE_TYPES, key=f"instance_type_{tab_key}_{i}")
            
            with col3:
                # Auto-populate Memory/CPU based on selected instance type
                memory_cpu = get_vm_size(instance_type, vm_sku_mapping)
                st.text_input(
                    "Memory / CPU", 
                    value=memory_cpu,
                    disabled=True,
                    key=f"memory_cpu_{tab_key}_{i}",
                    help="This value is automatically populated based on the selected Azure Instance Type"
                )
                        
            # Reservation options - per server
            col1, col2 = st.columns(2)
            with col1:
                reservation_type = st.selectbox(
                    "On Demand/Reservation", 
                    ["On Demand", "Reservation"], 
                    key=f"reservation_type_{tab_key}_{i}"
                )
            if reservation_type == "Reservation":
                with col2:
                    reservation_term = st.selectbox(
                        "Reservation Term", 
                        ["One Year", "Three Years"],
                        disabled=reservation_type != "Reservation", 
                        key=f"reservation_term_{tab_key}_{i}"
                    )

            # Cloud management - per server (different defaults for prod vs non-prod)
            col1, col2 = st.columns(2)
            with col1:
                # Production systems typically should not be parked, non-production can be
                default_opt = "Out" if is_production else "In"
                opt_in_out = st.selectbox(
                    "OptInOptOut", 
                    ["In", "Out"], 
                    index=1 if default_opt == "Out" else 0,
                    help="In - Can be parked/managed using ParkMyCloud\nOut - Cannot be parked/Managed using Park My Cloud",
                    key=f"opt_in_out_{tab_key}_{i}"
                )
            # Team information and Internet access - per server
            with col2:
                internet_access = st.selectbox(
                    "Outbound Internet Access Required", 
                    ["Yes", "No"], 
                    key=f"internet_access_{tab_key}_{i}"
                )

            # Park My Cloud Schedule and Team Name - per server
            col1, col2 = st.columns(2)
            with col2:
                team_name = st.text_input(
                    "Park My cloud team name and Member", 
                    disabled=opt_in_out != "In",
                    key=f"team_name_{tab_key}_{i}"
                )
            with col1:
                park_schedule = st.selectbox(
                    "Park My Cloud Schedule", 
                    PARK_SCHEDULES, 
                    disabled=opt_in_out != "In",
                    key=f"park_schedule_{tab_key}_{i}"
                )
            
            # Store server data in list for later
            server_config = {
                "Server Number": i+1,
                "Server Role": server_role,
                "Server Role Version": server_role_version,
                "Availability Set": availability_set if contains_pas(server_role) else "N/A",
                "AFS Server Name": primary_afs if contains_ascs(server_role) else "N/A",
                "OS Version": os_version,
                "Instance Type": instance_type,
                "Memory/CPU": get_vm_size(instance_type, vm_sku_mapping),
                "Reservation Type": reservation_type,
                "Reservation Term": reservation_term if reservation_type == "Reservation" else "N/A",
                "OptInOptOut": opt_in_out,
                "Park My Cloud Schedule": park_schedule if opt_in_out == "In" else "N/A",
                "Park My cloud team name and Member": team_name if opt_in_out == "In" else "N/A",
                "Outbound Internet Access Required": internet_access,
                "Server Type": "Primary"
            }
            
            # Add production-specific fields (per server)
            if is_production:
                server_config["Record Type"] = record_type
                server_config["AZ Selection"] = az_selection
                # Handle cluster value based on server role
                if requires_cluster(server_role):
                    server_config["Cluster"] = cluster
                    
                    # Add HA_Role field when Cluster is Yes and HA variant exists
                    if cluster == "Yes" and "DB2" in server_role:
                        server_config["HA_Role"] = "DB2 DB-HA"
                        server_config["HA_Zone"] = suggested_ha

                    elif cluster == "Yes":
                        ha_role = f"{server_role}-HA"
                        if ha_role in form_options['SERVER_ROLES']:
                            server_config["HA_Role"] = ha_role
                            server_config["HA_Zone"] = suggested_ha
                        else:
                            server_config["HA_Role"] = "HA variant not available"
                    else:
                        server_config["HA_Role"] = "N/A"
                        server_config["HA_Zone"] = "N/A"
                else:
                    server_config["Cluster"] = "N/A"
                    server_config["HA_Role"] = "N/A"
                    server_config["HA_Zone"] = "N/A"

            
            server_data.append(server_config)

    # DR SERVER CONFIGURATION SECTION (Only for Production)
    if is_production:
        st.subheader("Disaster Recovery (DR) Region Server Configuration")
        
        # Update the enabled DR servers list when number of servers changes
        current_num_servers = st.session_state[f'num_servers_{tab_key}']
        
        # If number of servers increased, add new DR servers to enabled list
        max_existing_index = max(st.session_state[f'dr_servers_enabled_{tab_key}']) if st.session_state[f'dr_servers_enabled_{tab_key}'] else -1
        for i in range(max_existing_index + 1, current_num_servers):
            if i not in st.session_state[f'dr_servers_enabled_{tab_key}']:
                st.session_state[f'dr_servers_enabled_{tab_key}'].append(i)
        
        # If number of servers decreased, remove DR servers that exceed the primary server count
        st.session_state[f'dr_servers_enabled_{tab_key}'] = [
            i for i in st.session_state[f'dr_servers_enabled_{tab_key}'] 
            if i < current_num_servers
        ]
        
        # Sort the enabled list to maintain order
        st.session_state[f'dr_servers_enabled_{tab_key}'].sort()
        
        # Show summary
        # enabled_count = len(st.session_state[f'dr_servers_enabled_{tab_key}'])
        # st.info(f"DR Servers: {enabled_count} of {current_num_servers} primary servers")
        
        # Create DR configuration only for enabled DR servers
        for i in st.session_state[f'dr_servers_enabled_{tab_key}']:
            dr_config = render_dr_server_config(tab_key, vm_sku_mapping, i, region_code, az_zone_mapping)
            dr_server_data.append(dr_config)

    # SUBMISSION FORM
    with st.form(f"submit_form_{tab_key}"):
        st.write("Review the configuration above and submit when ready.")
        if is_production:
            enabled_dr_count = len(st.session_state[f'dr_servers_enabled_{tab_key}'])
            st.write(f"**Summary:** {st.session_state[f'num_servers_{tab_key}']} Primary region servers + {enabled_dr_count} DR region servers = {st.session_state[f'num_servers_{tab_key}'] + enabled_dr_count} total servers")
        submit_button = st.form_submit_button(f"Submit {'Production' if is_production else 'Non-Production'} Request")

    # Handle form submission
    if submit_button:
        # General info - collect all the form data
        general_config = {
            "Form Type": "Production" if is_production else "Non-Production",
            "SAP Region": sap_region,
            "Azure Region": azure_region,
            "Azure Region Code": region_code,
            "Environment": environment,
            "SID": sid,
            "ITSG ID": itsg_id,
            "Number of Primary Servers": num_servers,
            "Number of DR Servers": len(st.session_state[f'dr_servers_enabled_{tab_key}']) if is_production else 0,
            "Total Servers": num_servers + (len(st.session_state[f'dr_servers_enabled_{tab_key}']) if is_production else 0),
            "Azure Subscription": azure_subscription,
            "Subnet/Zone": subnet,
            "Timezone": timezone,
            "Service Criticality": service_criticality,
        }
        
        # Add non-production specific fields to general config
        if not is_production:
            general_config["AZ Selection"] = az_selection
            general_config["Record Type"] = record_type
        
        # Combine primary and DR server data for production
        all_server_data = server_data.copy()
        if is_production:
            all_server_data.extend(dr_server_data)
        
        # Combine general config and server data
        form_data = {
            "general_config": general_config,
            "server_data": all_server_data,
            "primary_servers": server_data,
            "dr_servers": dr_server_data if is_production else []
        }
        
        # Save the data to a JSON file with appropriate naming
        form_type = "prod" if is_production else "nonprod"
        json_filename = f"sap_form_data_{form_type}_{sid}.json"
        json_file_path = save_form_data(form_data, json_filename)
        st.session_state[f'json_file_path_{tab_key}'] = json_file_path
        
        # Set form submitted state
        st.session_state[f'form_submitted_{tab_key}'] = True
        st.session_state[f'form_data_{tab_key}'] = form_data
        st.session_state[f'just_submitted_{tab_key}'] = True
        
        # Force rerun to show results
        st.rerun()

    # Display results after form submission
    if st.session_state.get(f'just_submitted_{tab_key}', False):
        # Clear the just_submitted flag to prevent re-processing
        st.session_state[f'just_submitted_{tab_key}'] = False
        st.success(f"{'Production' if is_production else 'Non-Production'} form submitted successfully!")
        
        # # Create a summary of the request
        # st.subheader("Request Summary")
        
        # # Display summary information
        form_data = st.session_state[f'form_data_{tab_key}']

        
        # Process the JSON file and generate Excel
        template_path = "Template.xlsx"
        
        # Check if template exists
        if not os.path.exists(template_path):
            st.error(f"Excel template not found: {template_path}")
            st.info("Please ensure 'Template.xlsx' is in the same directory as this script.")
        else:
            if "Amsterdam" in form_data['general_config'].get('Azure Region'):
                region = "AMS"
            else:
                region = "DUB"

            output_excel_path = f"{form_data['general_config'].get('SID', 'Unknown')}_{form_data['general_config'].get('SAP Region', '')}_{region}_Unilever_Build_sheet.xlsx"

            try:
                # Process the data and generate Excel file
                if not is_production:
                    excel_file_path = process_non_prod_data_to_excel(st.session_state[f'json_file_path_{tab_key}'], template_path, output_excel_path)
                else:
                    excel_file_path = process_prod_data_to_excel(st.session_state[f'json_file_path_{tab_key}'], template_path, output_excel_path)
                
                st.session_state[f'excel_file_path_{tab_key}'] = excel_file_path
                
                # Display success message
                st.success(f"Excel file generated successfully: {excel_file_path}")
                
                # Provide download link
                with open(excel_file_path, "rb") as file:
                    st.download_button(
                        label="📥 Download Excel File",
                        data=file,
                        file_name=os.path.basename(excel_file_path),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"download_{tab_key}"
                    )
                    
            except Exception as e:
                st.error(f"Error generating Excel file: {str(e)}")
                st.info("You can still use the JSON file for processing later.")
        
        # Reset button to clear the form
        if st.button(f"Reset {'Production' if is_production else 'Non-Production'} Form", key=f"reset_{tab_key}"):
            # Clear all session state for this tab
            keys_to_clear = [key for key in st.session_state.keys() if key.endswith(f'_{tab_key}')]
            for key in keys_to_clear:
                del st.session_state[key]
            # Also clear the just_submitted flag specifically
            if f'just_submitted_{tab_key}' in st.session_state:
                del st.session_state[f'just_submitted_{tab_key}']
            if f'dr_servers_enabled_{tab_key}' in st.session_state:
                del st.session_state[f'dr_servers_enabled_{tab_key}']
            st.rerun()


# Set page configuration
st.set_page_config(page_title="SAP Buildsheet Request Form", layout="wide")

# Add title and description
st.title("SAP Buildsheet Request Form")
st.markdown("Complete the form below to request SAP buildsheet generation")

# Create tabs for Production, Non-Production, and File Management
tab1, tab2, tab3 = st.tabs(["🏭 Production", "🧪 Non-Production", "📁 File Management"])

with tab1:
    st.header("Production Environment Request")
    render_form_content("prod", is_production=True)

with tab2:
    st.header("Non-Production Environment Request")
    render_form_content("nonprod", is_production=False)
    
with tab3:
    render_file_management_tab()
