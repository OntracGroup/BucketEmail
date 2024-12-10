#!/usr/bin/env python
# coding: utf-8

# In[ ]:

import streamlit as st
import pandas as pd
import smtplib
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.pdfgen import canvas
import io
from io import BytesIO
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import gspread
from fpdf import FPDF
from google.oauth2.service_account import Credentials
import sendgrid
from sendgrid.helpers.mail import Mail, Email, To, Content, Attachment
import base64
from PIL import Image
import math

def add_section_title(title, df):
    """Add a section title and return the dataframe with title as the first row."""
    title_row = pd.DataFrame([[title] * len(df.columns)], columns=df.columns)
    df_with_title = pd.concat([title_row, df], ignore_index=True)
    return df_with_title

def generate_pdf(side_by_side_df, loadout_productivity_df, swings_simulation_df, improved_cycle_df):
    """Generate a polished PDF with user results and separate tables for each section."""
    pdf_output = BytesIO()

    # Create the PDF document
    doc = SimpleDocTemplate(pdf_output, pagesize=letter, leftMargin=30, rightMargin=30, topMargin=30, bottomMargin=30)

    # Prepare the flowables (content) to add to the PDF
    elements = []

    # Add background and content
    add_background(None, None, side_by_side_df, loadout_productivity_df, swings_simulation_df, improved_cycle_df, elements)

    # Build the PDF
    doc.build(elements)

    pdf_output.seek(0)
    return pdf_output

def add_background(canvas, doc, side_by_side_df, loadout_productivity_df, swings_simulation_df, improved_cycle_df):
    """Add a dark background to the PDF page before adding content."""
    # Set the background color to dark
    canvas.setFillColor(colors.HexColor("#121212"))  # Dark background
    canvas.rect(0, 0, letter[0], letter[1], fill=True)

    # Now add the content
    elements = []  # List of all elements to be added to the PDF
    styles = getSampleStyleSheet()
    title_style = styles['Title']
    heading_style = styles['Heading1']
    subheading_style = styles['Heading2']
    normal_style = styles['Normal']
    
    # Custom styles for dark mode
    title_style.fontSize = 22
    title_style.textColor = colors.HexColor("#f4c542")  # Orange title color
    
    heading_style.fontSize = 18
    heading_style.textColor = colors.HexColor("#f4c542")  # Orange heading color
    
    subheading_style.fontSize = 14
    subheading_style.textColor = colors.HexColor("#f4c542")  # Orange subheading color
    subheading_style.underline = True  # Underline subheadings
    
    normal_style.fontSize = 10  # Small text for normal content
    normal_style.textColor = colors.HexColor("#e0e0e0")  # Light gray text color for dark mode
    
    # 1️⃣ Add Title
    elements.append(Paragraph("ONTRAC Excavator Bucket Optimization Results", title_style))
    elements.append(Spacer(1, 12))  # Space below the title

    # 2️⃣ Add Section 1: Side-by-Side Bucket Comparison
    elements.append(Paragraph("Side-by-Side Bucket Comparison", heading_style))
    elements.append(Spacer(1, 8))  # Space below the heading
    
    # Create the side-by-side table
    side_by_side_table_data = [side_by_side_df.columns.to_list()] + side_by_side_df.values.tolist()
    side_by_side_table = Table(side_by_side_table_data)
    side_by_side_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#2a2a2a")),  # Header background
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor("#ffffff")),  # Header text color
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor("#121212")),  # Table body background
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.HexColor("#e0e0e0")),  # Body text color
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#333333")),  # Grid lines
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('ALIGN', (0, 1), (-1, -1), 'CENTER'),
        ('PADDING', (0, 0), (-1, -1), 10),
    ]))
    elements.append(side_by_side_table)
    elements.append(Spacer(1, 20))  # Space below the table

    # 3️⃣ Add Section 2: Loadout Productivity & Truck Pass Simulation
    elements.append(Paragraph("Loadout Productivity & Truck Pass Simulation", heading_style))
    elements.append(Spacer(1, 8))  # Space below the heading
    
    # Create the loadout productivity table
    loadout_productivity_table_data = [loadout_productivity_df.columns.to_list()] + loadout_productivity_df.values.tolist()
    loadout_productivity_table = Table(loadout_productivity_table_data)
    loadout_productivity_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#2a2a2a")),  # Header background
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor("#ffffff")),  # Header text color
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor("#121212")),  # Table body background
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.HexColor("#e0e0e0")),  # Body text color
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#333333")),  # Grid lines
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('ALIGN', (0, 1), (-1, -1), 'CENTER'),
        ('PADDING', (0, 0), (-1, -1), 10),
    ]))
    elements.append(loadout_productivity_table)
    elements.append(Spacer(1, 20))  # Space below the table

    # 4️⃣ Add Section 3: Swings Simulation
    elements.append(Paragraph("Swings Simulation", heading_style))
    elements.append(Spacer(1, 8))  # Space below the heading
    
    # Create the swings simulation table
    swings_simulation_table_data = [swings_simulation_df.columns.to_list()] + swings_simulation_df.values.tolist()
    swings_simulation_table = Table(swings_simulation_table_data)
    swings_simulation_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#2a2a2a")),  # Header background
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor("#ffffff")),  # Header text color
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor("#121212")),  # Table body background
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.HexColor("#e0e0e0")),  # Body text color
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#333333")),  # Grid lines
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('ALIGN', (0, 1), (-1, -1), 'CENTER'),
        ('PADDING', (0, 0), (-1, -1), 10),
    ]))
    elements.append(swings_simulation_table)
    elements.append(Spacer(1, 20))  # Space below the table

    # 5️⃣ Add Section 4: Improved Cycle Times
    elements.append(Paragraph("Improved Cycle Times", heading_style))
    elements.append(Spacer(1, 8))  # Space below the heading
    
    # Create the improved cycle table
    improved_cycle_table_data = [improved_cycle_df.columns.to_list()] + improved_cycle_df.values.tolist()
    improved_cycle_table = Table(improved_cycle_table_data)
    improved_cycle_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#2a2a2a")),  # Header background
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor("#ffffff")),  # Header text color
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor("#121212")),  # Table body background
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.HexColor("#e0e0e0")),  # Body text color
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#333333")),  # Grid lines
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('ALIGN', (0, 1), (-1, -1), 'CENTER'),
        ('PADDING', (0, 0), (-1, -1), 10),
    ]))
    elements.append(improved_cycle_table)
    elements.append(Spacer(1, 20))  # Space below the table

    # Finalize and build the PDF with the elements
    doc.build(elements)


def adjust_payload_for_new_bucket(dump_truck_payload, new_payload):
    max_payload = dump_truck_payload * 1.10  # Allow up to 10% adjustment
    increment = dump_truck_payload * 0.001   # Fine adjustment increments

    # Try to achieve swing values within ±0.14 tolerance
    current_payload = dump_truck_payload
    while current_payload <= max_payload:
        swings_to_fill_truck_new = current_payload / new_payload
        if abs(swings_to_fill_truck_new - math.ceil(swings_to_fill_truck_new)) <= 0.05:
            return current_payload, swings_to_fill_truck_new
        current_payload += increment

    # If no suitable payload is found, return the original payload with calculated swings
    swings_to_fill_truck_new = dump_truck_payload / new_payload
    return dump_truck_payload, swings_to_fill_truck_new

def adjust_payload_for_old_bucket(dump_truck_payload, old_payload):
    max_payload = dump_truck_payload * 1.10  # Allow up to 10% adjustment
    increment = dump_truck_payload * 0.001   # Fine adjustment increments

    # Try to achieve swing values within ±0.14 tolerance
    current_payload = dump_truck_payload
    while current_payload <= max_payload:
        swings_to_fill_truck_old = current_payload / old_payload
        if abs(swings_to_fill_truck_old - math.ceil(swings_to_fill_truck_old)) <= 0.05:
            return current_payload, swings_to_fill_truck_old
        current_payload += increment

    # If no suitable payload is found, return the original payload with calculated swings
    swings_to_fill_truck_old = dump_truck_payload / old_payload
    return dump_truck_payload, swings_to_fill_truck_old

def generate_html_table(data, title):
    """
    Generate a simple HTML table from a dictionary where keys are column headers
    and values are lists of data. The table will have a dynamic title, styled for dark mode.
    """
    # Extract headers dynamically from the keys of the data dictionary
    headers = list(data.keys())
    
    # Find the maximum length of the lists (rows) in the data dictionary
    num_rows = max(len(data[header]) for header in headers)
    
    # Start the HTML table structure with fixed table width
    html = """
    <style>
        /* Global styles for Dark Mode */
        body {
            background-color: #121212; /* Dark background for the body */
            color: #e0e0e0; /* Light text for dark mode */
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 20px;
        }
        table {
            width: 100%; /* Set a fixed width for the table */
            margin: 0 auto; /* Center the table horizontally */
            border-collapse: collapse;
            font-size: 16px;
            text-align: left;
            background-color: #1e1e1e; /* Table background for dark mode */
            color: #e0e0e0; /* Light text for table content */
            border-radius: 8px; /* Rounded corners for modern look */
        }
        th, td {
            padding: 12px 15px;
            border: 1px solid #333; /* Border color for dark mode */
            text-align: center; /* Centered text for better readability */
        }
        th {
            background-color: #1e1e1e; /* Pale yellow-orange color for headers */
            color: #ffffff; /* White text for headers */
            font-weight: bold;
        }
        tr:nth-child(even) {
            background-color: #2a2a2a; /* Slightly lighter row for contrast */
        }
        tr:nth-child(odd) {
            background-color: #1e1e1e; /* Darker odd rows */
        }
        tr:hover {
            background-color: #444; /* Highlight row on hover */
        }
        h3 {
            font-size: 22px;
            color: #f4c542; /* Orange color for the title */
            font-weight: bold;
            border-bottom: 2px solid #f4c542;
            padding-bottom: 5px;
            margin-bottom: 5px; /* Reduced margin to remove gap */
        }
        /* Optionally style the container for better layout */
        .table-container {
            background-color: #181818;
            padding: 15px;
            border-radius: 10px;
        }
    </style>
    """
    
    # Use the title for both the h3 and table
    html += f"<h3>{title}</h3>"
    html += '<div class="table-container">'
    html += "<table><thead><tr>"
    
    # Add table headers
    for header in headers:
        html += f"<th>{header}</th>"
    
    html += "</tr></thead><tbody>"
    
    # Add rows to the table, ensuring to handle any missing data gracefully
    for i in range(num_rows):
        html += "<tr>"
        for header in headers:
            value = data[header][i] if i < len(data[header]) else ""
            html += f"<td>{value}</td>"
        html += "</tr>"
    
    html += "</tbody></table>"
    html += "</div>"
    
    return html

# Google Sheets credentials and setup
def connect_to_google_sheet(sheet_name):
    """Connect to Google Sheets and return a sheet object."""
    scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    credentials_dict = st.secrets["gcp_service_account"]  # Load service account credentials from Streamlit secrets
    credentials = Credentials.from_service_account_info(credentials_dict, scopes=scope)
    client = gspread.authorize(credentials)
    sheet = client.open("bucket_email").worksheet("Sheet1")  # Open the first sheet of the workbook
    return sheet

# Define custom CSS
st.markdown("""
    <style>
    .custom-font {
        font-family: 'Arial'; /* Change to your desired font */
        font-size: 16px;      /* Adjust the font size */
        font-weight: bold;    /* Make the text bold */
        color: #000000;       /* Customize the text color */
    }
    </style>
    """, unsafe_allow_html=True)

# Define your CSV file paths here (use raw strings or double backslashes)
swl_csv = 'excavator_swl.csv'  # Ensure this file exists
bucket_csv = 'bucket_data.csv'  # Make sure this file exists
bhc_bucket_csv = 'bhc_bucket_data.csv'  # Make sure this file exists
dump_truck_csv = 'dump_trucks.csv'  # Path to dump truck CSV

# Load datasets
def load_bucket_data(bucket_csv):
    return pd.read_csv(bucket_csv)

def load_bhc_bucket_data(bhc_bucket_csv):
    return pd.read_csv(bhc_bucket_csv)

def load_dump_truck_data(dump_truck_csv):
    return pd.read_csv(dump_truck_csv)

# Function to generate CSV data from a DataFrame
def load_excavator_swl_data(swl_csv):
    swl_data = pd.read_csv(swl_csv)
    swl_data['boom_length'] = pd.to_numeric(swl_data['boom_length'], errors='coerce')
    swl_data['arm_length'] = pd.to_numeric(swl_data['arm_length'], errors='coerce')
    swl_data['CWT'] = pd.to_numeric(swl_data['CWT'], errors='coerce')
    swl_data['shoe_width'] = pd.to_numeric(swl_data['shoe_width'], errors='coerce')
    swl_data['reach'] = pd.to_numeric(swl_data['reach'], errors='coerce')
    swl_data['class'] = pd.to_numeric(swl_data['class'], errors='coerce')
    return swl_data

# Load the data
dump_truck_data = load_dump_truck_data(dump_truck_csv)
swl_data = load_excavator_swl_data(swl_csv)

def generate_csv(comparison_df):
    buffer = io.StringIO()
    comparison_df.to_csv(buffer, index=False)
    buffer.seek(0)
    return buffer

# Email sending function
def send_email_with_csv(email, csv_data):
    try:
        msg = EmailMessage()
        msg['Subject'] = 'ONTRAC XMOR® Bucket Solution Results'
        msg['From'] = 'bucketontrac@gmail.com'
        msg['To'] = email

        # HTML content with a table for the comparison results
        html_content = f"""
        <html>
        <head>
            <style>
                table {{
                    width: 100%;
                    border-collapse: collapse;
                    margin-top: 20px;
                }}
                th, td {{
                    border: 1px solid #ddd;
                    text-align: left;
                    padding: 10px;
                }}
                th {{
                    background-color: #f2f2f2;
                    padding-top: 12px;
                    padding-bottom: 12px;
                    font-weight: bold;
                }}
                td {{
                    padding-top: 8px;
                    padding-bottom: 8px;
                }}
                    .subheading {{
                    background-color: #e0e0e0;
                    font-weight: bold;
                    text-align: center;
                    padding-top: 10px;
                    padding-bottom: 10px;
                }}
            </style>
        </head>
        <body>
        <h2>ONTRAC XMOR® Bucket Solution Results</h2>
        <p>G'day from the ONTRAC team! Here's your side-by-side comparison:</p>
        <table>
            <tr>
                <th>Description</th>
                <th>Old Bucket</th>
                <th>XMOR® Bucket</th>
                <th>Difference</th>
                <th>% Difference</th>
            </tr>
        """
        
        for index, row in comparison_df.iterrows():
            # Apply subheading style to specific rows (0, 6, 14, and 21)
            if index in [0, 6, 14, 21]:
                html_content += f"""
                <tr>
                    <td class="subheading" colspan="5">{row['Description']}</td>
                </tr>
                """
            else:
                html_content += f"""
                <tr>
                    <td>{row['Description']}</td>
                    <td>{row['Old Bucket']}</td>
                    <td>{row['XMOR® Bucket']}</td>
                    <td>{row['Difference']}</td>
                    <td>{row['% Difference']}</td>
                </tr>
                """
        
        html_content += """
        </table>
        <p>Thank you for using ONTRAC's bucket optimiser, we hope to hear from you soon!</p>
        </body>
        </html>
        """

        msg.add_alternative(html_content, subtype='html')
        
    # Send the email (using SMTP)
    
        with smtplib.SMTP_SSL(smtp_server, smtp_port) as smtp:
            smtp.login(email_username, email_password)  # Login using Streamlit secrets
            smtp.sendmail(msg['From'], msg['To'], msg.as_string())  # Send email
            print("Email sent successfully")
    except Exception as e:
        print(f"Failed to send email: {e}")

def send_email_with_pdf(email, pdf_file):
    """Send the PDF file via email."""
    # Retrieve SendGrid credentials from Streamlit secrets
    sendgrid_api_key = st.secrets["sendgrid"]["api_key"]
    from_email = st.secrets["sendgrid"]["from_email"]

    # Create the SendGrid client
    sg = sendgrid.SendGridAPIClient(api_key=sendgrid_api_key)

    # Create the email components
    from_email = Email(from_email)
    to_email = To(email)
    subject = "Your ONTRAC Excavator Results"
    content = Content("text/plain", "Please find the attached PDF with your results.")
    
    # Create the email object
    mail = Mail(from_email, to_email, subject, content)
    
    # Read and encode the PDF file as base64
    pdf_content = pdf_file.read()
    encoded_file = base64.b64encode(pdf_content).decode('utf-8')
    
    # Create the attachment object
    attachment = Attachment(
        file_content=encoded_file,
        file_type='application/pdf',
        file_name='comparison.pdf',
        disposition='attachment'
    )
    
    # Attach the file to the email
    mail.add_attachment(attachment)
    
    # Send the email
    try:
        response = sg.send(mail)
        if response.status_code == 202:
            print("Email sent successfully")
        else:
            print(f"Failed to send email: Status code {response.status_code}")
    except Exception as e:
        print(f"Failed to send email: {e}")
        
# Main Streamlit App UI
def app():
    st.write("Copyright © ONTRAC Group Pty Ltd 2024.")
    
ONTRAC_IMAGE = Image.open('dark_ontrac_logo.png')
# Show images
st.image([ONTRAC_IMAGE], width=200)

# Streamlit UI
st.title("ONTRAC XMOR® Bucket Solution\n\n")
st.caption("Select your load out equipment to discover the bucket for you.")

st.title("Excavator Selection")

# Excavator inputs
excavator_make = st.selectbox("Select Excavator Make", swl_data['make'].unique())
excavator_model = st.selectbox("Select Excavator Model", swl_data[swl_data['make'] == excavator_make]['model'].unique())

boom_length = st.selectbox("Select Boom Length (m)", swl_data[swl_data['model'] == excavator_model]['boom_length'].unique())
arm_length = st.selectbox("Select Arm Length (m)", swl_data[swl_data['model'] == excavator_model]['arm_length'].unique())
cwt = st.selectbox("Select Counterweight (CWT in kg)", swl_data[swl_data['model'] == excavator_model]['CWT'].unique())
shoe_width = st.selectbox("Select Shoe Width (mm)", swl_data[swl_data['model'] == excavator_model]['shoe_width'].unique())
reach = st.selectbox("Select Reach (m)", swl_data[swl_data['model'] == excavator_model]['reach'].unique())

st.title("Dump Truck Selection")

# Dump truck inputs
truck_brand = st.selectbox("Select Dump Truck Brand", dump_truck_data['brand'].unique())
truck_type = st.selectbox("Select Dump Truck Type", dump_truck_data[dump_truck_data['brand'] == truck_brand]['type'].unique())
truck_model = st.selectbox("Select Dump Truck Model", dump_truck_data[(dump_truck_data['brand'] == truck_brand) & 
                                                                     (dump_truck_data['type'] == truck_type)]['model'].unique())
truck_payload = st.selectbox("Select Dump Truck Payload (tons)", dump_truck_data[dump_truck_data['model'] == truck_model]['payload'].unique())

# Additional Inputs
st.title("Additional Information")
material_density = st.number_input("Material Density (kg/m³)     e.g. 1500", min_value=0.0)
quick_hitch_weight = st.number_input("Quick Hitch Weight (kg)     Leave as 0 for direct pin", min_value=0.0)
current_bucket_size = st.number_input("Current Bucket Size (m³)", min_value=0.0)
current_bucket_weight = st.number_input("Current Bucket Weight (kg)", min_value=0.0)
machine_swings_per_minute = st.number_input("Machine Swings per Minute", min_value=0.0)

# Checkbox for BHC buckets
select_bhc = st.checkbox("Select from BHC buckets only (Heavy Duty)")

# Find matching SWL and optimal bucket
def find_matching_swl(user_data):
    matching_excavator = swl_data[
        (swl_data['make'] == user_data['make']) &
        (swl_data['model'] == user_data['model']) & 
        (swl_data['CWT'] == user_data['cwt']) & 
        (swl_data['shoe_width'] == user_data['shoe_width']) & 
        (swl_data['reach'] == user_data['reach']) & 
        (swl_data['boom_length'] == user_data['boom_length']) & 
        (swl_data['arm_length'] == user_data['arm_length'])
    ]
    if matching_excavator.empty:
        return None
    swl = matching_excavator.iloc[0]['swl']
    return swl

# Function to calculate bucket load
def calculate_bucket_load(bucket_size, material_density):
    return bucket_size * material_density

def select_optimal_bucket(user_data, bucket_data, swl):
    current_bucket_size = user_data['current_bucket_size']
    optimal_bucket = None
    highest_bucket_size = 0

    selected_model = user_data['model']
    excavator_class = swl_data[swl_data['model'] == selected_model]['class'].iloc[0]

    for index, bucket in bucket_data.iterrows():
        if bucket['class'] > excavator_class + 10:
            continue

        bucket_load = calculate_bucket_load(bucket['bucket_size'], user_data['material_density'])
        total_bucket_weight = user_data['quick_hitch_weight'] + bucket_load + bucket['bucket_weight']

        if total_bucket_weight <= swl and bucket['bucket_size'] > highest_bucket_size:
            highest_bucket_size = bucket['bucket_size']
            optimal_bucket = {
                'bucket_name': bucket['bucket_name'],
                'bucket_size': highest_bucket_size,
                'bucket_weight': bucket['bucket_weight'],
                'total_bucket_weight': total_bucket_weight
            }

    return optimal_bucket

# Get user input data
user_data = {
    'make': excavator_make,
    'model': excavator_model,
    'boom_length': boom_length,
    'arm_length': arm_length,
    'cwt': cwt,
    'shoe_width': shoe_width,
    'reach': reach,
    'material_density': material_density,
    'quick_hitch_weight': quick_hitch_weight,
    'current_bucket_size': current_bucket_size,
    'current_bucket_weight': current_bucket_weight,
    'dump_truck_payload': truck_payload,
    'machine_swings_per_minute': machine_swings_per_minute
}
    
def generate_comparison_df(user_data, optimal_bucket, swl):
    # Load selected bucket data
    selected_bucket_csv = bhc_bucket_csv if select_bhc else bucket_csv
    bucket_data = load_bucket_data(selected_bucket_csv)

    optimal_bucket = select_optimal_bucket(user_data, bucket_data, swl)

    if optimal_bucket:
        old_capacity = user_data['current_bucket_size']
        new_capacity = optimal_bucket['bucket_size']
        old_payload = calculate_bucket_load(old_capacity, user_data['material_density'])
        new_payload = calculate_bucket_load(new_capacity, user_data['material_density'])

        dump_truck_payload = user_data['dump_truck_payload'] * 1000
        machine_swings_per_minute = user_data['machine_swings_per_minute']

        # Total suspended load
        old_total_load = old_payload + user_data['current_bucket_weight'] + user_data['quick_hitch_weight']
        new_total_load = optimal_bucket['total_bucket_weight']  # Corrected variable

        # Adjust payload for the new bucket using the function
        dump_truck_payload_new, swings_to_fill_truck_new = adjust_payload_for_new_bucket(dump_truck_payload, new_payload)
        dump_truck_payload_old, swings_to_fill_truck_old = adjust_payload_for_old_bucket(dump_truck_payload, old_payload)

        # Time to fill truck in minutes
        time_to_fill_truck_old = swings_to_fill_truck_old / machine_swings_per_minute
        time_to_fill_truck_new = swings_to_fill_truck_new / machine_swings_per_minute

        # Average number of trucks per hour at 75% efficiency
        avg_trucks_per_hour_old = (60 / time_to_fill_truck_old) * 0.75 if time_to_fill_truck_old > 0 else 0
        avg_trucks_per_hour_new = (60 / time_to_fill_truck_new) * 0.75 if time_to_fill_truck_new > 0 else 0

        # Swings per hour
        swings_per_hour_old = swings_to_fill_truck_old * avg_trucks_per_hour_old
        swings_per_hour_new = swings_to_fill_truck_new * avg_trucks_per_hour_new

        # Total swings per hour
        total_swings_per_hour = 60 * machine_swings_per_minute

        # Truck Tonnes per hour
        truck_tonnage_per_hour_old = swings_per_hour_old * old_capacity * user_data['material_density'] / 1000
        truck_tonnage_per_hour_new = swings_per_hour_new * new_capacity * user_data['material_density'] / 1000

        # Production (t/hr)
        total_tonnage_per_hour_old = total_swings_per_hour * old_capacity * user_data['material_density'] / 1000
        total_tonnage_per_hour_new = total_swings_per_hour * new_capacity * user_data['material_density'] / 1000

        # Production (t/hr)
        tonnage_per_hour_old = avg_trucks_per_hour_old * dump_truck_payload_old / 1000
        tonnage_per_hour_new = avg_trucks_per_hour_new * dump_truck_payload_new / 1000

        # Assuming 1800 swings in a day
        total_m3_per_day_old = 1000 * old_capacity
        total_m3_per_day_new = 1000 * new_capacity

        # Total tonnage per day
        total_tonnage_per_day_old = total_m3_per_day_old * user_data['material_density'] / 1000
        total_tonnage_per_day_new = total_m3_per_day_new * user_data['material_density'] / 1000

        # Total number of trucks per day
        total_trucks_per_day_old = total_tonnage_per_day_old / dump_truck_payload * 1000
        total_trucks_per_day_new = total_tonnage_per_day_new / dump_truck_payload * 1000

        Productivity = f"{(1.1 * total_tonnage_per_hour_new - total_tonnage_per_hour_old) / total_tonnage_per_hour_old * 100:.0f}%"

        st.success(f"Great news! ONTRAC could improve your productivity by up to {Productivity}!")
        st.success(f"Your ONTRAC XMOR® Bucket Solution is the: {optimal_bucket['bucket_name']} ({optimal_bucket['bucket_size']} m³)")
        # Load and display images from a local directory
        XMOR_IMAGE = Image.open('XMOR_BHC_IMAGE.png') if select_bhc else Image.open('XMOR_BHB_IMAGE.png')

        # Show images
        st.image([XMOR_IMAGE], caption=[f"{optimal_bucket['bucket_name']} ({optimal_bucket['bucket_size']} m³)"], width=400)
    
       # Side-by-Side Bucket Comparison Data
        side_by_side_data = {
            'Description': [
                 'Capacity (m³)', 'Material Density (kg/m³)', 'Bucket Payload (kg)', 
                'Total Suspended Load (kg)'
            ],
            'Old Bucket': [
                 f"{old_capacity:.1f}", f"{user_data['material_density']:.0f}", f"{old_payload:.0f}", 
                f"{old_total_load:.0f}"
            ],
            'XMOR® Bucket': [
                 f"{new_capacity:.1f}", f"{user_data['material_density']:.0f}", f"{new_payload:.0f}", 
                f"{new_total_load:.0f}"
            ],
            'Difference': [
                 f"{new_capacity - old_capacity:.1f}", '-', f"{new_payload - old_payload:.0f}", 
                f"{new_total_load - old_total_load:.0f}"
            ],
            '% Difference': [
                 f"{(new_capacity - old_capacity) / old_capacity * 100:.0f}%", '-', f"{(new_payload - old_payload) / old_payload * 100:.0f}%", 
                f"{(new_total_load - old_total_load) / old_total_load * 100:.0f}%"
            ]
        }
        
        # Loadout Productivity & Truck Pass Simulation Data
        loadout_productivity_data = {
            'Description': [
                 f"{truck_brand} {truck_model} Payload (kg)", 'Avg No. Swings to Fill Truck', 
                'Time to Fill Truck (min)', 'Avg Trucks/Hour @ 75% eff', 'Swings/Hour', 'Tonnes/Hour'
            ],
            'Old Bucket': [
                 f"{dump_truck_payload_old:.0f}{'*' if dump_truck_payload_old != dump_truck_payload else ''}", f"{swings_to_fill_truck_old:.1f}", 
                f"{time_to_fill_truck_old:.1f}", f"{avg_trucks_per_hour_old:.1f}", f"{swings_per_hour_old:.0f}", f"{truck_tonnage_per_hour_old:.0f}"
            ],
            'XMOR® Bucket': [
                 f"{dump_truck_payload_new:.0f}{'*' if dump_truck_payload_new != dump_truck_payload else ''}", f"{swings_to_fill_truck_new:.1f}", 
                f"{time_to_fill_truck_new:.1f}", f"{avg_trucks_per_hour_new:.1f}", f"{swings_per_hour_new:.0f}", f"{truck_tonnage_per_hour_new:.0f}"
            ],
            'Difference': [
                 f"{dump_truck_payload_new - dump_truck_payload_old:.0f}", f"{swings_to_fill_truck_new - swings_to_fill_truck_old:.1f}", 
                f"{time_to_fill_truck_new - time_to_fill_truck_old:.1f}", f"{avg_trucks_per_hour_new - avg_trucks_per_hour_old:.1f}",
                "-", f"{truck_tonnage_per_hour_new - truck_tonnage_per_hour_old:.0f}"
            ],
            '% Difference': [
                 f"{(dump_truck_payload_new - dump_truck_payload_old) / dump_truck_payload_old * 100:.0f}%", 
                f"{(swings_to_fill_truck_new - swings_to_fill_truck_old) / swings_to_fill_truck_old * 100:.0f}%",
                f"{(time_to_fill_truck_new - time_to_fill_truck_old) / time_to_fill_truck_old * 100:.0f}%",
                f"{(avg_trucks_per_hour_new - avg_trucks_per_hour_old) / avg_trucks_per_hour_old * 100:.0f}%",
                "-",
                f"{(truck_tonnage_per_hour_new - truck_tonnage_per_hour_old) / truck_tonnage_per_hour_old * 100:.0f}%"
            ]
        }
        
        # 1000 Swings Side-by-Side Simulation Data
        swings_simulation_data = {
            'Description': [
                 'Number of Swings', 'Total Volume (m³)', 
                'Total Tonnes', 'Total Trucks'
            ],
            'Old Bucket': [
                '1000', f"{total_m3_per_day_old:.0f}", f"{total_tonnage_per_day_old:.0f}", 
                f"{total_trucks_per_day_old:.0f}"
            ],
            'XMOR® Bucket': [
                '1000', f"{total_m3_per_day_new:.0f}", f"{total_tonnage_per_day_new:.0f}", 
                f"{total_trucks_per_day_new:.0f}"
            ],
            'Difference': [
                '-', f"{total_m3_per_day_new - total_m3_per_day_old:.0f}", 
                f"{total_tonnage_per_day_new - total_tonnage_per_day_old:.0f}", 
                f"{total_trucks_per_day_new - total_trucks_per_day_old:.0f}"
            ],
            '% Difference': [
                '-', f"{(total_m3_per_day_new - total_m3_per_day_old) / total_m3_per_day_old * 100:.0f}%", 
                f"{(total_tonnage_per_day_new - total_tonnage_per_day_old) / total_tonnage_per_day_old * 100:.0f}%", 
                f"{(total_trucks_per_day_new - total_trucks_per_day_old) / total_trucks_per_day_old * 100:.0f}%"
            ]
        }
        
        # 10% Improved Cycle Time Simulation Data
        improved_cycle_data = {
            'Description': [
                 'Number of Swings', 'Total Volume (m³)', 
                'Total Tonnes', 'Total Trucks'
            ],
            'Old Bucket': [
                '1000', f"{total_m3_per_day_old:.0f}", f"{total_tonnage_per_day_old:.0f}", 
                f"{total_trucks_per_day_old:.0f}"
            ],
            'XMOR® Bucket': [
                '1100', f"{1.1 * total_m3_per_day_new:.0f}", f"{1.1 * total_tonnage_per_day_new:.0f}", 
                f"{1.1 * total_trucks_per_day_new:.0f}"
            ],
            'Difference': [
                '100', f"{1.1 * total_m3_per_day_new - total_m3_per_day_old:.0f}", 
                f"{1.1 * total_tonnage_per_day_new - total_tonnage_per_day_old:.0f}", 
                f"{1.1 * total_trucks_per_day_new - total_trucks_per_day_old:.0f}"
            ],
            '% Difference': [
                '10%', f"{(1.1 * total_m3_per_day_new - total_m3_per_day_old) / total_m3_per_day_old * 100:.0f}%", 
                f"{(1.1 * total_tonnage_per_day_new - total_tonnage_per_day_old) / total_tonnage_per_day_old * 100:.0f}%", 
                f"{(1.1 * total_trucks_per_day_new - total_trucks_per_day_old) / total_trucks_per_day_old * 100:.0f}%"
            ]
        }
        # Function to add a title row
        def add_section_title(title, data):
            # Create a DataFrame with the title row
            title_row = pd.DataFrame([[title] + [''] * (len(data.columns) - 1)], columns=data.columns)
            # Concatenate title row and the data
            return pd.concat([title_row, data], ignore_index=True)
        
        # Example DataFrames for each section
        side_by_side_df = pd.DataFrame(side_by_side_data)
        loadout_productivity_df = pd.DataFrame(loadout_productivity_data)
        swings_simulation_df = pd.DataFrame(swings_simulation_data)
        improved_cycle_df = pd.DataFrame(improved_cycle_data)
        
        # Adding section headers
        side_by_side_with_title = add_section_title("Side-by-Side Bucket Comparison", side_by_side_df)
        loadout_productivity_with_title = add_section_title("Loadout Productivity & Truck Pass Simulation", loadout_productivity_df)
        swings_simulation_with_title = add_section_title("1000 Swings Side-by-Side Simulation", swings_simulation_df)
        improved_cycle_with_title = add_section_title("10% Improved Cycle Time Simulation", improved_cycle_df)
        
        # Combine all sections into one DataFrame
        final_df = pd.concat([
            side_by_side_with_title,
            loadout_productivity_with_title,
            swings_simulation_with_title,
            improved_cycle_with_title
        ], ignore_index=True)
    
        if final_df is not None:
                st.title('XMOR® Productivity Comparison')
            
                # Call the function for each table with the appropriate title
                st.markdown(generate_html_table(side_by_side_data, "Side-by-Side Bucket Comparison"), unsafe_allow_html=True)
                st.markdown(generate_html_table(loadout_productivity_data, "Loadout Productivity & Truck Pass Simulation"), unsafe_allow_html=True)
                st.markdown(generate_html_table(swings_simulation_data, "1000 Swings Side-by-Side Simulation"), unsafe_allow_html=True)
                st.markdown(generate_html_table(improved_cycle_data, "10% Improved Cycle Time Simulation"), unsafe_allow_html=True)
            
                # Optional notes about dump truck fill factor
                if dump_truck_payload_new != dump_truck_payload:
                    st.write(f"*Dump Truck fill factor of {(100 * dump_truck_payload_new / dump_truck_payload):.1f}% applied for XMOR® Bucket pass matching.")
                if dump_truck_payload_old != dump_truck_payload:
                    st.write(f"*Dump Truck fill factor of {(100 * dump_truck_payload_old / dump_truck_payload):.1f}% applied for Old Bucket pass matching.")
            
                # Provide additional details for calculations
                st.write(f"Total Suspended Load (XMOR® Bucket): {optimal_bucket['total_bucket_weight']:.0f}kg")
                st.write(f"Safe Working Load at {user_data['reach']}m reach ({user_data['make']} {user_data['model']}): {swl:.0f}kg")
                st.write(f"Calculations based on the {user_data['make']} {user_data['model']} with a {user_data['boom_length']}m boom, {user_data['arm_length']}m arm, {user_data['cwt']}kg counterweight, {user_data['shoe_width']}mm shoes, operating at a reach of {user_data['reach']}m, and with a material density of {user_data['material_density']:.0f}kg/m³.")
                st.write(f"Dump Truck: {truck_brand} {truck_model}, Rated payload = {user_data['dump_truck_payload'] * 1000:.0f}kg")

        return side_by_side_df, loadout_productivity_df, swings_simulation_df, improved_cycle_df

def collect_email(sheet, user_data, optimal_bucket, comparison_df):
    """Collect the user's email and store it in the Google Sheet."""
    if 'email_form_submitted' not in st.session_state:
        st.session_state.email_form_submitted = False
    
    with st.form("email_form", clear_on_submit=True):
        email = st.text_input("Enter your email to receive your comparison!", placeholder="you@example.com")
        submit_button = st.form_submit_button("Yes Please!")
        
    if submit_button:
        if "@" in email and "." in email:  # Basic email validation
            
            # Generate PDF with the results
            pdf_file = generate_pdf(side_by_side_df, loadout_productivity_df, swings_simulation_df, improved_cycle_df)
            # Send the email with the PDF attached
            send_email_with_pdf(email, pdf_file)
            st.success("Please check your inbox!")
            
            try:
                sheet.append_row([email])  # Add email to the Google Sheet
                
                # Allow users to submit another email
                st.session_state.email_form_submitted = True
            except Exception as e:
                st.error(f"An error occurred: {e}")
        else:
            st.error("Please enter a valid email address.")
            
# Connect to Google Sheets
try:
    sheet = connect_to_google_sheet("bucket_email")  # Your Google Sheet name
except Exception as e:
    st.error(f"Unable to connect to Google Sheets: {e}")
    sheet = None
    
# **CALCULATION LOGIC**
# Run calculations only when the button is pressed
if 'calculate_button' not in st.session_state:
    st.session_state.calculate_button = False
    
# Button to trigger calculations
calculate_button = st.button('Calculate!', on_click=lambda: st.session_state.update({'calculate_button': True}))

if st.session_state.calculate_button:
    swl = find_matching_swl(user_data)  # Calculate matching SWL
    if swl:
        # Load selected bucket data
        selected_bucket_csv = bhc_bucket_csv if select_bhc else bucket_csv
        bucket_data = load_bucket_data(selected_bucket_csv)
        optimal_bucket = select_optimal_bucket(user_data, bucket_data, swl)
        
        if optimal_bucket:
            # Generate DataFrame for comparison
            comparison_df = generate_comparison_df(user_data, optimal_bucket, swl)
            side_by_side_df, loadout_productivity_df, swings_simulation_df, improved_cycle_df = comparison_df = generate_comparison_df(user_data, optimal_bucket, swl)
            #pdf_file = generate_pdf(side_by_side_df, loadout_productivity_df, swings_simulation_df, improved_cycle_df)
            
            # Ask for email after successful calculation
            if sheet:  # Ensure the sheet is connected
                st.subheader('Would you like a side-by-side comparison sent to your email?')
                collect_email(sheet, user_data, optimal_bucket, comparison_df)
        else:
            st.warning("No suitable bucket found within SWL limits.")
    else:
        st.warning("No matching excavator configuration found!")

# Run the Streamlit app
if __name__ == '__main__':
    app()
