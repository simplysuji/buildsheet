import streamlit as st
import pandas as pd
import json
import os
import pandas as pd
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
    return "PAS" in server_role.upper() or "AAS" in server_role.upper()

# Function to check if server role contains ASCS
def contains_ascs(server_role):
    return "ASCS" == server_role.upper()

# Add this function after the get_region_code function and before load_vm_sku_data

def get_dr_az_zone(azure_region, azure_subscription):
    """
    Get the DR AZ zone based on the primary server's region and subscription
    
    Args:
        azure_region (str): The selected Azure region
        azure_subscription (str): The selected Azure subscription
        
    Returns:
        str: The corresponding DR AZ zone
    """
    # Determine if it's Amsterdam or Dublin
    is_amsterdam = "Amsterdam" in azure_region or "NLWE" in azure_region
    is_dublin = "Dublin" in azure_region or "IENO" in azure_region
    
    # Extract subscription type from the full subscription name
    if "01" in azure_subscription or "Global" in azure_subscription:
        subscription_type = "Global"
    elif "02" in azure_subscription or "Sirius" in azure_subscription:
        subscription_type = "Sirius"
    elif "03" in azure_subscription or "U2K2" in azure_subscription:
        subscription_type = "U2K2"
    elif "04" in azure_subscription or "Cordillera" in azure_subscription:
        subscription_type = "Cordillera"
    elif "05" in azure_subscription or "Fusion" in azure_subscription:
        subscription_type = "Fusion"
    elif "98" in azure_subscription or "Model Environment" in azure_subscription:
        subscription_type = "ME"
    else:
        subscription_type = "Global"  # Default fallback
    
    # AZ mapping based on the table provided
    az_mapping = {
        ("Global", "Amsterdam"): "3",
        ("Global", "Dublin"): "2",
        ("Sirius", "Amsterdam"): "3",
        ("Sirius", "Dublin"): "2",
        ("U2K2", "Amsterdam"): "1",
        ("U2K2", "Dublin"): "2",
        ("Cordillera", "Amsterdam"): "3",
        ("Cordillera", "Dublin"): "1",
        ("Fusion", "Amsterdam"): "3",
        ("Fusion", "Dublin"): "2",
        ("ME", "Amsterdam"): "1",
        ("ME", "Dublin"): "2"
    }
    
    # Determine location key
    if is_amsterdam:
        location = "Amsterdam"
    elif is_dublin:
        location = "Dublin"
    else:
        location = "Amsterdam"  # Default fallback
    
    # Get the AZ zone
    return az_mapping.get((subscription_type, location), "1")  # Default to "1" if not found


def render_dr_server_config(tab_key, vm_sku_mapping, server_index):
    """
    Render DR server configuration form
    
    Args:
        tab_key (str): Tab key for session state management
        vm_sku_mapping (dict): VM SKU mapping data
        server_index (int): Index of the server
    
    Returns:
        dict: DR server configuration data
    """
    SERVICE_CRITICALITY = ["SC 1", "SC 2", "SC 3", "SC 4"]
    RECORD_TYPES = ["A Record", "CNAME"]
    OS_VERSIONS = [
        "RHEL 7.9 for SAP", "RHEL 8.10 SAP", "SLES 12 SP3", "SLES 12 SP4", 
        "SLES 12 SP5", "SLES 15 SP1", "SLES 15 SP2", "Windows 2016", "Windows 2019", "Windows 2022", "Windows 2025"
    ]
    INSTANCE_TYPES = list(vm_sku_mapping.keys()) if vm_sku_mapping else [
        "D8asv4", "E8asv4", "E16_v3", "E16as_v4", "E16ds_v4", "E16s_v3", "E2_v3", 
        "E20as_v4", "E20ds_v4", "E20s_v3", "E2as_v4", "E2ds_v4", "E2s_v3", "E32_v3",
        "E32as_v4", "E32ds_v4", "E32s_v3", "E4_v3", "E48as_v4", "E48ds_v4", "E48s_v3",
        "E4as_v4", "E4ds_v4", "E4s_v3"
    ]
    PARK_SCHEDULES = ["Weekdays-12 hours Snooze(5pm IST to 5am IST) and Weekends Off",
        "Weekdays-12 hours Snooze(6pm IST to 6am IST) and Weekends Off",
        "Weekdays-12 hours Snooze(7pm IST to 7am IST) and Weekends Off",
        "Weekdays-12 hours Snooze(8pm IST to 8am IST) and Weekends Off",
        "Weekdays-12 hours Snooze(9pm IST to 9am IST) and Weekends Off",
        "Weekdays-12 hours Snooze(10pm IST to 10am IST) and Weekends Off",
        "Weekdays-12 hours Snooze(11pm IST to 11am IST) and Weekends Off"
    ]
    TIMEZONE = ["IST", "UTC", "CET", "GMT", "CST", "PST", "EST", "BST"]
    
    # Get the primary server role to create DR equivalent
    primary_server_role = st.session_state.get(f"server_role_{tab_key}_{server_index}", "")
    
    # Suggest DR server role based on primary server role
    dr_server_roles = [
        "AAS-DR", "ASCS-DR", "HANA DB-DR", "PAS-DR", "SCS-DR", 
        "Web Dispatcher-DR", "iSCSI SBD-DR", "DB2-DR"
    ]

    # Try to auto-suggest DR role
    suggested_dr_role = ""
    if "AAS" in primary_server_role and "DR" not in primary_server_role:
        suggested_dr_role = "AAS-DR"
    elif "ASCS" in primary_server_role and "DR" not in primary_server_role:
        suggested_dr_role = "ASCS-DR"
    elif "HANA DB" in primary_server_role and "DR" not in primary_server_role:
        suggested_dr_role = "HANA DB-DR"
    elif "PAS" in primary_server_role and "DR" not in primary_server_role:
        suggested_dr_role = "PAS-DR"
    elif "SCS" in primary_server_role and "DR" not in primary_server_role:
        suggested_dr_role = "SCS-DR"
    elif "Web Dispatcher" in primary_server_role and "DR" not in primary_server_role:
        suggested_dr_role = "Web Dispatcher-DR"
    elif "iSCSI SBD" in primary_server_role and "DR" not in primary_server_role:
        suggested_dr_role = "iSCSI SBD-DR"
    
    # Set default index for DR role
    default_dr_index = 0
    if suggested_dr_role and suggested_dr_role in dr_server_roles:
        default_dr_index = dr_server_roles.index(suggested_dr_role)
    
    with st.expander(f"🔄 DR Server {server_index+1} Configuration", expanded=True):      
        # First row: Server Role, A Record/CNAME, Service Criticality
        col1, col2, col3 = st.columns(3)
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
        with col3:
            dr_service_criticality = st.selectbox(
                "Service Criticality", 
                SERVICE_CRITICALITY, 
                key=f"dr_service_criticality_{tab_key}_{server_index}"
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
            # Auto-determine DR AZ Selection based on primary server's region and subscription
            primary_azure_region = st.session_state.get(f"azure_region_{tab_key}", "")
            primary_azure_subscription = st.session_state.get(f"azure_subscription_{tab_key}", "")
            suggested_dr_az = get_dr_az_zone(primary_azure_region, primary_azure_subscription)
            
            # Available AZ zones
            dr_az_zones = ["1", "2", "3"]
            
            # Set default index for suggested AZ
            default_az_index = 0
            if suggested_dr_az in dr_az_zones:
                default_az_index = dr_az_zones.index(suggested_dr_az)
            
            dr_az_selection = st.selectbox(
                "AZ Selection - Zone", 
                dr_az_zones,
                index=default_az_index,
                key=f"dr_az_selection_{tab_key}_{server_index}",
                help=f"Suggested: {suggested_dr_az} (based on primary server region and subscription)"
            )
        
        # Show AFS Servername option if server role contains PAS
        dr_afs_needed = "NA"  # Default value
        if contains_ascs(dr_server_role):
            dr_afs_needed = st.text_input("AFS Server Name", 
                                        key=f"dr_afs_needed_{tab_key}_{server_index}",
                                        help="Required for ASCS servers")
        
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
            dr_park_schedule = st.selectbox(
                "Park My Cloud Schedule", 
                PARK_SCHEDULES, 
                disabled=dr_opt_in_out != "In",
                key=f"dr_park_schedule_{tab_key}_{server_index}"
            )
        
        # Sixth row: Team information, Internet access, Timezone
        col1, col2, col3 = st.columns(3)
        with col1:
            dr_team_name = st.text_input(
                "Park My cloud team name and Member", 
                disabled=dr_opt_in_out != "In",
                key=f"dr_team_name_{tab_key}_{server_index}"
            )
        with col2:
            dr_internet_access = st.selectbox(
                "Outbound Internet Access Required", 
                ["Yes", "No"], 
                key=f"dr_internet_access_{tab_key}_{server_index}"
            )
        with col3:
            dr_timezone = st.selectbox(
                "Timezone", 
                TIMEZONE,
                key=f"dr_timezone_{tab_key}_{server_index}"
            )
    
    # Return DR server configuration
    dr_server_config = {
        "Server Number": f"DR-{server_index+1}",
        "Server Role": dr_server_role,
        "Record Type": dr_record_type,
        "Service Criticality": dr_service_criticality,
        "OS Version": dr_os_version,
        "Availability Set": dr_availability_set if contains_pas(dr_server_role) else "N/A",
        "AFS Server Name": dr_afs_needed if contains_ascs(dr_server_role) else "N/A",
        "AZ Selection": dr_az_selection,
        "Azure Instance Type": dr_instance_type,
        "Memory/CPU": get_vm_size(dr_instance_type, vm_sku_mapping),
        "Reservation Type": dr_reservation_type,
        "Reservation Term": dr_reservation_term if dr_reservation_type == "Reservation" else "N/A",
        "OptInOptOut": dr_opt_in_out,
        "Park My Cloud Schedule": dr_park_schedule if dr_opt_in_out == "In" else "N/A",
        "Park My cloud team name and Member": dr_team_name if dr_opt_in_out == "In" else "N/A",
        "Outbound Internet Access Required": dr_internet_access,
        "Timezone": dr_timezone,
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

    # Define options based on Production vs Non-Production
    if is_production:
        # Production-specific configurations
        ENVIRONMENTS = ["Production"]
        SUBNETS = ["Production STS"]

    else:
        # Non-Production configurations (existing)
        ENVIRONMENTS = ["Fix Development", "Fix Quality", "Fix Regression", "Fix Performance", 
                       "Project performance", "Project Development", "Project Quality", "Training", 
                       "Sandbox", "Project UAT"]
        SUBNETS = ["Non-Production STS"]
    
    # Common options for both
    AZURE_SUBSCRIPTIONS = [
            "SAP Technical Services-01 (Global)", "SAP Technical Services-02 (Sirius)",
            "SAP Technical Services-03 (U2K2)", "SAP Technical Services-04 (Cordillera)",
            "SAP Technical Services-05 (Fusion)", "SAP Technical Services-98 (Model Environment)"
    ]
    SAP_REGIONS = ["Sirius", "U2K2", "Cordillera", "Global", "POC/Model Env", "Fusion"]
    SERVICE_CRITICALITY = ["SC 1", "SC 2", "SC 3", "SC 4"]
    AZURE_REGIONS = [
        "Azure: Northern Europe (Dublin) (IENO)",
        "Azure: Western Europe (Amsterdam) (NLWE)",
    ]
    AZ_ZONES = ["1", "2", "3"]
    SERVER_ROLES = [
        "AAS", "ASCS", "ASCS+NFS", "ASCS+PAS", "ASCS+PAS+NFS", "ASCS-HA", "SCS+PAS+NFS", 
        "CS+NFS", "CS+NFS+PAS", "DB2 DB", "HANA DB", "HANA DB-HA", "iSCSI SBD", "SCS", "SCS+NFS", "SCS+PAS",
        "PAS", "Web Dispatcher",  "Web Dispatcher-HA", "Maxdb", "CS", "Optimizers", "IQ roles",
        "PAS-DR", "AAS-DR", "ASCS-DR", "HANA DB-DR", "SCS-DR", "SCS-HA", "iSCSI SBD-DR", "Web Dispatcher-DR"
    ]
    OS_VERSIONS = [
        "RHEL 7.9 for SAP", "RHEL 8.10 SAP", "SLES 12 SP3", "SLES 12 SP4", 
        "SLES 12 SP5", "SLES 15 SP1", "SLES 15 SP2", "Windows 2016", "Windows 2019", "Windows 2022", "Windows 2025"
    ]
    
    vm_sku_mapping = load_vm_sku_data()
    INSTANCE_TYPES = list(vm_sku_mapping.keys()) if vm_sku_mapping else [
        "D8asv4", "E8asv4", "E16_v3", "E16as_v4", "E16ds_v4", "E16s_v3", "E2_v3", 
        "E20as_v4", "E20ds_v4", "E20s_v3", "E2as_v4", "E2ds_v4", "E2s_v3", "E32_v3",
        "E32as_v4", "E32ds_v4", "E32s_v3", "E4_v3", "E48as_v4", "E48ds_v4", "E48s_v3",
        "E4as_v4", "E4ds_v4", "E4s_v3"
    ]

    RECORD_TYPES = ["A Record", "CNAME"]
    PARK_SCHEDULES = ["Weekdays-12 hours Snooze(5pm IST to 5am IST) and Weekends Off",
        "Weekdays-12 hours Snooze(6pm IST to 6am IST) and Weekends Off",
        "Weekdays-12 hours Snooze(7pm IST to 7am IST) and Weekends Off",
        "Weekdays-12 hours Snooze(8pm IST to 8am IST) and Weekends Off",
        "Weekdays-12 hours Snooze(9pm IST to 9am IST) and Weekends Off",
        "Weekdays-12 hours Snooze(10pm IST to 10am IST) and Weekends Off",
        "Weekdays-12 hours Snooze(11pm IST to 11am IST) and Weekends Off"
    ]
    TIMEZONE = ["IST", "UTC", "CET", "CEST", "GMT", "CST", "PST", "EST", "BST"]

    # GENERAL CONFIGURATION SECTION
    st.subheader("General Configuration")

    # First Row
    col1, col2 = st.columns(2)
    with col1:
        sap_region = st.selectbox("SAP Region", SAP_REGIONS, key=f"sap_region_{tab_key}")
    with col2:
        azure_region = st.selectbox("Azure Region", AZURE_REGIONS, key=f"azure_region_{tab_key}")

    # Auto-determine region code based on azure_region selection
    region_code = get_region_code(azure_region)

    # Display the auto-selected region code for user confirmation
    st.info(f"Region Code set to: **{region_code}** (based on {azure_region})")

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

    # Fourth Row - Only for Non-Production
    if not is_production:
        col1, col2, col3 = st.columns(3)
        with col2:
            az_selection = st.selectbox("AZ Selection - Zone", AZ_ZONES, key=f"az_selection_{tab_key}")
        with col1:
            azure_subscription = st.selectbox("Azure Subscription", AZURE_SUBSCRIPTIONS, key=f"azure_subscription_{tab_key}")
        with col3:
            record_type = st.selectbox("A Record / CNAME", RECORD_TYPES, key=f"record_type_{tab_key}")
    else:
        # Production only has Azure Subscription (no fourth row layout)
        col1, col2 = st.columns(2)
        with col1:
            azure_subscription = st.selectbox("Azure Subscription", AZURE_SUBSCRIPTIONS, key=f"azure_subscription_{tab_key}")
        with col2:
            st.empty()

    # SERVER CONFIGURATION SECTION
    st.subheader("Primary Server Configuration")

    # Create a section for each server
    server_data = []
    dr_server_data = []
    
    for i in range(st.session_state[f'num_servers_{tab_key}']):
        with st.expander(f"🖥️ Primary Server {i+1}", expanded=True):
                       
            # Basic server info
            col1, col2 = st.columns(2)
            with col1:
                server_role = st.selectbox("Server Role", SERVER_ROLES, key=f"server_role_{tab_key}_{i}")
            with col2:
                service_criticality = st.selectbox("Service Criticality", SERVICE_CRITICALITY, key=f"service_criticality_{tab_key}_{i}")
            
            # Show Availability Set option if server role contains PAS
            availability_set = "No"  # Default value
            if contains_pas(server_role):
                availability_set = st.selectbox("Availability Set", ["Yes", "No"], 
                                              key=f"availability_set_{tab_key}_{i}",
                                              help="Required for PAS servers for high availability")
                
            # Show AFS Servername option if server role contains PAS
            afs_needed = "NA"  # Default value
            if contains_ascs(server_role):
                afs_needed = st.text_input("AFS Server Name", 
                                              key=f"afs_needed_{tab_key}_{i}",
                                              help="Required for ASCS servers")
            
            # Production-specific fields: A Record/CNAME, AZ Selection, and Cluster
            if is_production:
                col1, col2, col3 = st.columns(3)
                with col1:
                    record_type = st.selectbox("A Record / CNAME", RECORD_TYPES, key=f"record_type_{tab_key}_{i}")
                with col2:
                    az_selection = st.selectbox("AZ Selection - Zone", AZ_ZONES, key=f"az_selection_{tab_key}_{i}")
                with col3:
                    cluster = st.selectbox("Cluster", 
                                         ["Yes", "No"],
                                         key=f"cluster_{tab_key}_{i}",
                                         help="Required for Production environments")
            
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
            with col2:
                park_schedule = st.selectbox(
                    "Park My Cloud Schedule", 
                    PARK_SCHEDULES, 
                    disabled=opt_in_out != "In",
                    key=f"park_schedule_{tab_key}_{i}"
                )

            # Team information and Internet access - per server
            col1, col2, col3 = st.columns(3)
            with col1:
                team_name = st.text_input(
                    "Park My cloud team name and Member", 
                    disabled=opt_in_out != "In",
                    key=f"team_name_{tab_key}_{i}"
                )
            with col2:
                internet_access = st.selectbox(
                    "Outbound Internet Access Required", 
                    ["Yes", "No"], 
                    key=f"internet_access_{tab_key}_{i}"
                )
            with col3:
                timezone = st.selectbox(
                    "Timezone", 
                    TIMEZONE,
                    key=f"timezone_{tab_key}_{i}"
                )
            
            # Store server data in list for later
            server_config = {
                "Server Number": i+1,
                "Server Role": server_role,
                "Service Criticality": service_criticality,
                "Availability Set": availability_set if contains_pas(server_role) else "N/A",
                "AFS Server Name": afs_needed if contains_ascs(server_role) else "N/A",
                "OS Version": os_version,
                "Instance Type": instance_type,
                "Memory/CPU": get_vm_size(instance_type, vm_sku_mapping),
                "Reservation Type": reservation_type,
                "Reservation Term": reservation_term if reservation_type == "Reservation" else "N/A",
                "OptInOptOut": opt_in_out,
                "Park My Cloud Schedule": park_schedule if opt_in_out == "In" else "N/A",
                "Park My cloud team name and Member": team_name if opt_in_out == "In" else "N/A",
                "Outbound Internet Access Required": internet_access,
                "Timezone": timezone,
                "Server Type": "Primary"
            }
            
            # Add production-specific fields (per server)
            if is_production:
                server_config["Record Type"] = record_type
                server_config["AZ Selection"] = az_selection
                server_config["Cluster"] = cluster
                
                # Add HA_Role field when Cluster is Yes and HA variant exists
                if cluster == "Yes":
                    ha_role = f"{server_role}-HA"
                    if ha_role in SERVER_ROLES:
                        server_config["HA_Role"] = ha_role
                    else:
                        server_config["HA_Role"] = "HA variant not available"
            
            server_data.append(server_config)

    # DR SERVER CONFIGURATION SECTION (Only for Production)
    if is_production:
        st.subheader("Disaster Recovery (DR) Server Configuration")
        
        # Create DR configuration for each primary server
        for i in range(st.session_state[f'num_servers_{tab_key}']):
            dr_config = render_dr_server_config(tab_key, vm_sku_mapping, i)
            dr_server_data.append(dr_config)

    # SUBMISSION FORM
    with st.form(f"submit_form_{tab_key}"):
        st.write("Review the configuration above and submit when ready.")
        if is_production:
            st.write(f"**Summary:** {st.session_state[f'num_servers_{tab_key}']} Primary servers + {st.session_state[f'num_servers_{tab_key}']} DR servers = {st.session_state[f'num_servers_{tab_key}'] * 2} total servers")
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
            "Number of DR Servers": num_servers if is_production else 0,
            "Total Servers": num_servers * 2 if is_production else num_servers,
            "Azure Subscription": azure_subscription,
            "Subnet/Zone": subnet
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
        
        # Force rerun to show results
        st.rerun()

    # Display results after form submission
    if st.session_state[f'form_submitted_{tab_key}']:
        st.success(f"{'Production' if is_production else 'Non-Production'} form submitted successfully!")
        
        # # Create a summary of the request
        # st.subheader("Request Summary")
        
        # # Display summary information
        form_data = st.session_state[f'form_data_{tab_key}']
        
        # col1, col2, col3 = st.columns(3)
        # with col1:
        #     st.metric("Primary Servers", len(form_data['primary_servers']))
        # with col2:
        #     if is_production:
        #         st.metric("DR Servers", len(form_data['dr_servers']))
        #     else:
        #         st.metric("DR Servers", "N/A")
        # with col3:
        #     st.metric("Total Servers", len(form_data['server_data']))
        
        # # Show server breakdown
        # if is_production:
        #     with st.expander("📋 Server Configuration Summary", expanded=True):
        #         st.write("**Primary Servers:**")
        #         for server in form_data['primary_servers']:
        #             st.write(f"- Server {server['Server Number']}: {server['Server Role']} ({server['Instance Type']})")
                
        #         st.write("**DR Servers:**")
        #         for dr_server in form_data['dr_servers']:
        #             st.write(f"- {dr_server['Server Number']}: {dr_server['Server Role']} ({dr_server['Azure Instance Type']})")
        
        # Process the JSON file and generate Excel
        template_path = "Template.xlsx"
        
        # Check if template exists
        if not os.path.exists(template_path):
            st.error(f"Excel template not found: {template_path}")
            st.info("Please ensure 'Template.xlsx' is in the same directory as this script.")
        else:
            form_type = "prod" if is_production else "nonprod"
            output_excel_path = f"SAP_Request_{form_type}_{form_data['general_config'].get('SID', 'Unknown')}_{form_data['general_config'].get('Environment', '')}.xlsx"
            
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
            st.rerun()


# Set page configuration
st.set_page_config(page_title="SAP Buildsheet Request Form", layout="wide")

# Add title and description
st.title("SAP Buildsheet Request Form")
st.markdown("Complete the form below to request SAP buildsheet generation")

# Create tabs for Production and Non-Production
tab1, tab2 = st.tabs(["🏭 Production", "🧪 Non-Production"])

with tab1:
    st.header("Production Environment Request")
    render_form_content("prod", is_production=True)

with tab2:
    st.header("Non-Production Environment Request")
    render_form_content("nonprod", is_production=False)