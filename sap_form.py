import streamlit as st
import pandas as pd
import json
import os
import pandas as pd
from excel_processor import process_sap_data_to_excel

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

# Set page configuration
st.set_page_config(page_title="SAP Buildsheet Request Form", layout="wide")

# Add title and description
st.title("SAP Buildsheet Request Form")
st.markdown("Complete the form below to request SAP buildsheet generation")

# Initialize session state for number of servers
if 'num_servers' not in st.session_state:
    st.session_state.num_servers = 1

if 'form_submitted' not in st.session_state:
    st.session_state.form_submitted = False
    
if 'excel_file_path' not in st.session_state:
    st.session_state.excel_file_path = None

if 'json_file_path' not in st.session_state:
    st.session_state.json_file_path = None

# Store server data for summary
if 'server_data' not in st.session_state:
    st.session_state.server_data = []

# Define options from the PDF documents
SAP_REGIONS = ["Sirius", "U2K2", "Cordillera", "Global", "POC/Model Env", "Fusion"]
AZURE_REGIONS = [
    "Azure: Northern Europe (Dublin) (IENO)",
    "Azure: Western Europe (Amsterdam) (NLWE)",
    "Azure: Northern Europe NEW (Dublin) (IENO)",
    "Azure: Central India (Pune) (INCE)"
]
REGION_CODES = ["bnlwe", "bieno"]
AZ_ZONES = ["1", "2", "3"]
ENVIRONMENTS = ["Fix Development", "Fix Quality", "Fix Regression", "Fix Performance", 
                "Project performance", "Project Development", "Project Quality", "Training", 
                "Sandbox", "Project UAT", "Production"]
SERVER_ROLES = [
    "AAS", "AAS-DR", "ASCS", "ASCS", "ASCS+NFS", "ASCS+PAS", "ASCS+PAS+NFS", "ASCS-DR", "ASCS-HA",
    "CS+NFS", "CS+NFS+PAS", "DB2 DB", "HANA DB", "HANA DB-DR", "HANA DB-HA", "iSCSI SBD", "iSCSI SBD-DR",
    "PAS", "PAS-DR", "SCS", "SCS+NFS", "SCS+PAS", "SCS+PAS+NFS", "SCS-DR", "SCS-HA",
    "Web Dispatcher", "Web Dispatcher-DR", "Web Dispatcher-HA"
]
SERVICE_CRITICALITY = ["SC 1", "SC 2", "SC 3", "SC 4"]
OS_VERSIONS = [
    "RHEL 7.9 for SAP", "RHEL 8.10 SAP", "SLES 12 SP3", "SLES 12 SP4", 
    "SLES 12 SP5", "SLES 15 SP1", "SLES 15 SP2"
]
INSTANCE_TYPES = [
    "D8asv4", "E8asv4", "E16_v3", "E16as_v4", "E16ds_v4", "E16s_v3", "E2_v3", 
    "E20as_v4", "E20ds_v4", "E20s_v3", "E2as_v4", "E2ds_v4", "E2s_v3", "E32_v3",
    "E32as_v4", "E32ds_v4", "E32s_v3", "E4_v3", "E48as_v4", "E48ds_v4", "E48s_v3",
    "E4as_v4", "E4ds_v4", "E4s_v3"
]
MEMORY_CPU = [
    "8 vCPU, 32 GiB", "8 vCPU, 64 GiB", "16vCPUs/128GBs", "2vCPUs/16GBs", 
    "20vCPUs/160GBs", "32vCPUs/256GBs", "4vCPUs/32GBs", "48vCPUs/384GBs"
]
RECORD_TYPES = ["A Record", "CNAME"]
SUBNETS = ["Production STS", "Non-Production STS"]
AZURE_SUBSCRIPTIONS = [
    "SAP Technical Services-01 (Global)", "SAP Technical Services-02 (Sirius)",
    "SAP Technical Services-03 (U2K2)", "SAP Technical Services-04 (Cordillera)",
    "SAP Technical Services-05 (Fusion)", "SAP Technical Services-98 (Model Environment)"
]
PARK_SCHEDULES = ["Weekdays-12 hours Snooze(11pm IST to 11am IST) and Weekends Off"]
TIMEZONE = ["IST", "UTC", "CET", "GMT", "CST", "PST", "EST"]


# Function to update number of servers
def update_num_servers():
    st.session_state.num_servers = st.session_state.num_servers_input

# Function to check if server role contains PAS
def contains_pas(server_role):
    return "PAS" in server_role.upper()

# GENERAL CONFIGURATION SECTION - Outside any form
st.subheader("General Configuration")

# First Row
col1, col2, col3 = st.columns(3)
with col1:
    sap_region = st.selectbox("SAP Region", SAP_REGIONS, key="sap_region")
with col2:
    azure_region = st.selectbox("Azure Region", AZURE_REGIONS, key="azure_region")
with col3:
    region_code = st.selectbox("Azure Region Code", REGION_CODES, key="region_code")

# Second Row
col1, col2, col3 = st.columns(3)
with col1:
    az_selection = st.selectbox("AZ Selection - Zone", AZ_ZONES, key="az_selection")
with col2:
    environment = st.selectbox("Environment", ENVIRONMENTS, key="environment")
with col3:
    sid = st.text_input("SID", key="sid")

# Third Row
col1, col2 = st.columns(2)
with col1:
    itsg_id = st.text_input("ITSG ID", key="itsg_id")
with col2:
    # This number input is outside the form to allow immediate UI updates
    num_servers = st.number_input("Number of Servers", 
                                min_value=1, 
                                value=st.session_state.num_servers,
                                key="num_servers_input",
                                on_change=update_num_servers)

# Fourth Row
col1, col2 = st.columns(2)
with col1:
    record_type = st.selectbox("A Record / CNAME", RECORD_TYPES, key="record_type")
with col2:
    subnet = st.selectbox("Subnet/Zone", SUBNETS, key="subnet")

# Fifth Row
col1, col2 = st.columns(2)
with col1:
    azure_subscription = st.selectbox("Azure Subscription", AZURE_SUBSCRIPTIONS, key="azure_subscription")
with col2:
    cluster = st.text_input("Cluster", disabled=subnet != "Production", key="cluster")

# SERVER CONFIGURATION SECTION
st.subheader("Server Configuration")

# Create a section for each server
server_data = []
for i in range(st.session_state.num_servers):
    with st.expander(f"Server {i+1}", expanded=True):
        # Basic server info
        col1, col2 = st.columns(2)
        with col1:
            server_role = st.selectbox("Server Role", SERVER_ROLES, key=f"server_role_{i}")
        with col2:
            service_criticality = st.selectbox("Service Criticality", SERVICE_CRITICALITY, key=f"service_criticality_{i}")
        
        # Show Availability Set option if server role contains PAS
        availability_set = "No"  # Default value
        if contains_pas(server_role):
            availability_set = st.selectbox("Availability Set", ["Yes", "No"], 
                                          key=f"availability_set_{i}",
                                          help="Required for PAS servers for high availability")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            os_version = st.selectbox("OS Version", OS_VERSIONS, key=f"os_version_{i}")
        with col2:
            instance_type = st.selectbox("Azure Instance Type", INSTANCE_TYPES, key=f"instance_type_{i}")
        with col3:
            memory_cpu = st.selectbox("Memory / CPU", MEMORY_CPU, key=f"memory_cpu_{i}")
            
        # Reservation options - per server
        col1, col2 = st.columns(2)
        with col1:
            reservation_type = st.selectbox(
                "On Demand/Reservation", 
                ["On Demand", "Reservation"], 
                key=f"reservation_type_{i}"
            )
        with col2:
            reservation_term = st.text_input(
                "Reservation Term", 
                disabled=reservation_type != "Reservation", 
                key=f"reservation_term_{i}"
            )

        # Cloud management - per server
        col1, col2 = st.columns(2)
        with col1:
            opt_in_out = st.selectbox(
                "OptInOptOut", 
                ["In", "Out"], 
                help="In - Can be parked/managed using ParkMyCloud\nOut - Cannot be parked/Managed using Park My Cloud",
                key=f"opt_in_out_{i}"
            )
        with col2:
            park_schedule = st.selectbox(
                "Park My Cloud Schedule", 
                PARK_SCHEDULES, 
                disabled=opt_in_out != "In",
                key=f"park_schedule_{i}"
            )

        # Team information and Internet access - per server
        col1, col2, col3 = st.columns(3)
        with col1:
            team_name = st.text_input(
                "Park My cloud team name and Member", 
                disabled=opt_in_out != "In",
                key=f"team_name_{i}"
            )
        with col2:
            internet_access = st.selectbox(
                "Outbound Internet Access Required", 
                ["Yes", "No"], 
                key=f"internet_access_{i}"
            )
        with col3:
            timezone = st.selectbox(
                "Timezone", 
                TIMEZONE,
                key=f"timezone_{i}"
            )
        
        
        # Store server data in list for later
        server_data.append({
            "Server Number": i+1,
            "Server Role": server_role,
            "Service Criticality": service_criticality,
            "Availability Set": availability_set if contains_pas(server_role) else "N/A",
            "OS Version": os_version,
            "Instance Type": instance_type,
            "Memory/CPU": memory_cpu,
            "Reservation Type": reservation_type,
            "Reservation Term": reservation_term if reservation_type == "Reservation" else "N/A",
            "OptInOptOut": opt_in_out,
            "Park My Cloud Schedule": park_schedule if opt_in_out == "In" else "N/A",
            "Park My cloud team name and Member": team_name if opt_in_out == "In" else "N/A",
            "Outbound Internet Access Required": internet_access,
            "Timezone": timezone
        })

# SUBMISSION FORM - Separate simple form just for submission
with st.form("submit_form"):
    st.write("Review the configuration above and submit when ready.")
    submit_button = st.form_submit_button("Submit Full Request")

# Handle form submission
if submit_button:
    # General info - collect all the form data
    general_config = {
        "SAP Region": sap_region,
        "Azure Region": azure_region,
        "Azure Region Code": region_code,
        "AZ Selection": az_selection,
        "Environment": environment,
        "SID": sid,
        "ITSG ID": itsg_id,
        "Number of Servers": num_servers,
        "Record Type": record_type,
        "Subnet/Zone": subnet,
        "Azure Subscription": azure_subscription,
        "Cluster": cluster if subnet == "Production" else "N/A"
    }
    
    # Combine general config and server data
    form_data = {
        "general_config": general_config,
        "server_data": server_data
    }
    
    # Save the data to a JSON file
    json_file_path = save_form_data(form_data)
    st.session_state.json_file_path = json_file_path
    
    # Set form submitted state to trigger processing outside the form
    st.session_state.form_submitted = True
    
    # Store form data in session state
    st.session_state.form_data = form_data
    
    # Force rerun to show results outside form
    st.rerun()

# Display results after form submission
if st.session_state.form_submitted:
    st.success("Form submitted successfully!")
    
    # Create a summary of the request
    st.subheader("Request Summary")
    
    # Display success message with file path
    # st.success(f"Form data saved to: {st.session_state.json_file_path}")
    
    # Process the JSON file and generate Excel
    template_path = "Template.xlsx"
    
    # Check if template exists
    if not os.path.exists(template_path):
        st.error(f"Excel template not found: {template_path}")
        st.info("Please ensure 'Template.xlsx' is in the same directory as this script.")
    else:
        # Generate a unique output filename
        output_excel_path = f"SAP_Request_{st.session_state.form_data['general_config'].get('SID', 'Unknown')}_{st.session_state.form_data['general_config'].get('Environment', '')}.xlsx"
        
        try:
            # Process the data and generate Excel file
            excel_file_path = process_sap_data_to_excel(st.session_state.json_file_path, template_path, output_excel_path)
            st.session_state.excel_file_path = excel_file_path
            
            # Display success message
            st.success(f"Excel file generated successfully: {excel_file_path}")
            
            # Provide download link (outside the form)
            with open(excel_file_path, "rb") as file:
                st.download_button(
                    label="📥 Download Excel File",
                    data=file,
                    file_name=os.path.basename(excel_file_path),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
        except Exception as e:
            st.error(f"Error generating Excel file: {str(e)}")
            st.info("You can still use the JSON file for processing later.")
    
        # # Display JSON summary if requested
        # if st.checkbox("Show JSON data"):
        #     st.json(st.session_state.form_data)
        
    # Reset button to clear the form
    if st.button("Reset Form"):
        st.session_state.form_submitted = False
        st.session_state.server_data = []
        st.rerun()