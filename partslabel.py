import streamlit as st
import pandas as pd
import os
import re
import io
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph, PageBreak
from reportlab.lib.units import cm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER

# Configure Streamlit page
st.set_page_config(
    page_title="Part Label Generator",
    page_icon="🏷️",
    layout="wide"
)

# Style definitions
bold_style_v1 = ParagraphStyle(
    name='Bold_v1',
    fontName='Helvetica-Bold',
    fontSize=10,
    alignment=TA_LEFT,
    leading=20,
    spaceBefore=2,
    spaceAfter=2
)

bold_style_v2 = ParagraphStyle(
    name='Bold_v2',
    fontName='Helvetica-Bold',
    fontSize=10,
    alignment=TA_LEFT,
    leading=12,
    spaceBefore=0,
    spaceAfter=15,
)

desc_style = ParagraphStyle(
    name='Description',
    fontName='Helvetica',
    fontSize=20,
    alignment=TA_LEFT,
    leading=16,
    spaceBefore=2,
    spaceAfter=2
)

def format_part_no_v1(part_no):
    """Format part number with first 7 characters in 17pt font, rest in 22pt font."""
    if not part_no or not isinstance(part_no, str):
        part_no = str(part_no)

    if len(part_no) > 5:
        split_point = len(part_no) - 5
        part1 = part_no[:split_point]
        part2 = part_no[-5:]
        return Paragraph(f"<b><font size=17>{part1}</font><font size=22>{part2}</font></b>", bold_style_v1)
    else:
        return Paragraph(f"<b><font size=17>{part_no}</font></b>", bold_style_v1)

def format_part_no_v2(part_no):
    """Format part number with different font sizes to prevent overlapping."""
    if not part_no or not isinstance(part_no, str):
        part_no = str(part_no)

    if len(part_no) > 5:
        split_point = len(part_no) - 5
        part1 = part_no[:split_point]
        part2 = part_no[-5:]
        return Paragraph(f"<b><font size=34>{part1}</font><font size=40>{part2}</font></b><br/><br/>", bold_style_v2)
    else:
        return Paragraph(f"<b><font size=34>{part_no}</font></b><br/><br/>", bold_style_v2)

def format_description_v1(desc):
    """Format description text with dynamic font sizing based on length for v1."""
    if not desc or not isinstance(desc, str):
        desc = str(desc)
    
    # Dynamic font sizing based on description length
    desc_length = len(desc)
    
    if desc_length <= 30:
        font_size = 15
    elif desc_length <= 50:
        font_size = 13
    elif desc_length <= 70:
        font_size = 11
    elif desc_length <= 90:
        font_size = 10
    else:
        font_size = 9
        # Truncate very long descriptions to prevent overflow
        desc = desc[:100] + "..." if len(desc) > 100 else desc
    
    # Create a custom style for this description
    desc_style_v1 = ParagraphStyle(
        name='Description_v1',
        fontName='Helvetica',
        fontSize=font_size,
        alignment=TA_LEFT,
        leading=font_size + 2,
        spaceBefore=1,
        spaceAfter=1
    )
    
    return Paragraph(desc, desc_style_v1)

def format_description(desc):
    """Format description text with proper wrapping."""
    if not desc or not isinstance(desc, str):
        desc = str(desc)
    return Paragraph(desc, desc_style)

def parse_location_string_v1(location_str):
    """Parse a location string like "12M - LH -R-0-2-A-1" into its 7 components."""
    location_parts = [''] * 7

    if not location_str or not isinstance(location_str, str):
        return location_parts

    location_str = location_str.strip()
    pattern = r'([^_\s]+)'
    matches = re.findall(pattern, location_str)

    for i, match in enumerate(matches[:7]):
        location_parts[i] = match

    return location_parts

def parse_location_string_v2(location_str):
    """Parse a location string like "12M_ST-140_R_0_2_A_1" into its 7 components."""
    location_parts = [''] * 7

    if not location_str or not isinstance(location_str, str):
        return location_parts

    location_str = location_str.strip()
    pattern = r'([^_\s]+)'
    matches = re.findall(pattern, location_str)

    for i, match in enumerate(matches[:7]):
        location_parts[i] = match

    return location_parts

def generate_labels_from_excel_v1(df, progress_bar=None, status_text=None):
    """Generate labels using version 1 formatting."""
    
    # Create a BytesIO buffer to store the PDF
    buffer = io.BytesIO()
    
    # Set up key measurements
    part_no_height = 1.3 * cm
    desc_loc_height = 0.8 * cm

    # Identify column names in the file
    cols = df.columns.tolist()
    df.columns = [col.upper() for col in df.columns]
    cols = df.columns.tolist()

    # Find columns
    part_no_col = next((col for col in cols if 'PART' in col and ('NO' in col or 'NUM' in col or '#' in col)),
                      next((col for col in cols if col in ['PARTNO', 'PART']), None))
    desc_col = next((col for col in cols if 'DESC' in col), None)
    loc_col = next((col for col in cols if 'LOC' in col or 'POS' in col), None)

    if not part_no_col:
        part_no_col = cols[0]
    if not desc_col:
        desc_col = cols[1] if len(cols) > 1 else part_no_col
    if not loc_col:
        loc_col = cols[2] if len(cols) > 2 else desc_col

    if status_text:
        status_text.text(f"Using columns: Part No: {part_no_col}, Description: {desc_col}, Location: {loc_col}")

    # Group parts by location
    df_grouped = df.groupby(loc_col)
    total_locations = len(df_grouped)

    doc = SimpleDocTemplate(buffer, pagesize=A4)
    elements = []

    MAX_LABELS_PER_PAGE = 4
    label_count = 0

    for i, (location, group) in enumerate(df_grouped):
        try:
            if progress_bar:
                progress_value = int((i / total_locations) * 100)
                progress_bar.progress(progress_value)
            
            if status_text:
                status_text.text(f"Processing location {i+1}/{total_locations}: {location}")

            parts = group.head(2)

            if len(parts) < 2:
                if len(parts) == 1:
                    part1 = parts.iloc[0]
                    part2 = parts.iloc[0]
                else:
                    continue
            else:
                part1 = parts.iloc[0]
                part2 = parts.iloc[1]

            if label_count > 0 and label_count % MAX_LABELS_PER_PAGE == 0:
                elements.append(PageBreak())

            label_count += 1

            part_no_1 = str(part1[part_no_col])
            desc_1 = str(part1[desc_col])
            part_no_2 = str(part2[part_no_col])
            desc_2 = str(part2[desc_col])
            location_str = str(part1[loc_col])
            location_values = parse_location_string_v1(location_str)

            # Create tables for both parts with dynamic description formatting
            part_table = Table(
                [['Part No', format_part_no_v1(part_no_1)],
                 ['Description', format_description_v1(desc_1)]],
                colWidths=[4*cm, 11*cm],
                rowHeights=[part_no_height, desc_loc_height]
            )

            part_table.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('ALIGN', (0, 0), (0, -1), 'CENTRE'),
                ('ALIGN', (1, 0), (1, -1), 'LEFT'),
                ('VALIGN', (0, 0), (0, 0), 'MIDDLE'),
                ('VALIGN', (1, 0), (1, 0), 'MIDDLE'),
                ('VALIGN', (0, 1), (0, 1), 'MIDDLE'),
                ('VALIGN', (1, 1), (1, 1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 5),
                ('RIGHTPADDING', (0, 0), (-1, -1), 5),
                ('TOPPADDING', (0, 0), (-1, -1), 3),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                ('FONTNAME', (0, 0), (0, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (0, -1), 16),
            ]))

            part_table2 = Table(
                [['Part No', format_part_no_v1(part_no_2)],
                 ['Description', format_description_v1(desc_2)]],
                colWidths=[4*cm, 11*cm],
                rowHeights=[part_no_height, desc_loc_height]
            )

            part_table2.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('ALIGN', (0, 0), (0, -1), 'CENTRE'),
                ('ALIGN', (1, 0), (1, -1), 'LEFT'),
                ('VALIGN', (0, 0), (0, 0), 'MIDDLE'),
                ('VALIGN', (1, 0), (1, 0), 'MIDDLE'),
                ('VALIGN', (0, 1), (0, 1), 'MIDDLE'),
                ('VALIGN', (1, 1), (1, 1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 5),
                ('RIGHTPADDING', (0, 0), (-1, -1), 5),
                ('TOPPADDING', (0, 0), (-1, -1), 3),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                ('FONTNAME', (0, 0), (0, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (0, -1), 16),
            ]))

            # Location table
            location_data = [['Part Location'] + location_values]
            first_col_width = 4 * cm
            location_widths = [first_col_width]
            remaining_width = 11 * cm
            col_proportions = [1.8, 2.7, 1.3, 1.3, 1.3, 1.3, 1.3]
            total_proportion = sum(col_proportions)
            adjusted_widths = [w * remaining_width / total_proportion for w in col_proportions]
            location_widths.extend(adjusted_widths)

            location_table = Table(
                location_data,
                colWidths=location_widths,
                rowHeights=desc_loc_height
            )

            location_colors = [
                colors.HexColor('#E9967A'),
                colors.HexColor('#ADD8E6'),
                colors.HexColor('#90EE90'),
                colors.HexColor('#FFD700'),
                colors.HexColor('#ADD8E6'),
                colors.HexColor('#E9967A'),
                colors.HexColor('#90EE90')
            ]

            location_style = [
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (0, 0), 'TOP'),
                ('VALIGN', (1, 0), (-1, 0), 'TOP'),
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (0, 0), 16),
                ('FONTSIZE', (1, 0), (-1, -1), 14),
            ]

            for j, color in enumerate(location_colors):
                location_style.append(('BACKGROUND', (j+1, 0), (j+1, 0), color))

            location_table.setStyle(TableStyle(location_style))

            elements.append(part_table)
            elements.append(Spacer(1, 0.3 * cm))
            elements.append(part_table2)
            elements.append(location_table)
            elements.append(Spacer(1, 0.2 * cm))

        except Exception as e:
            if status_text:
                status_text.text(f"Error processing location {location}: {e}")
            continue

    if progress_bar:
        progress_bar.progress(100)

    if elements:
        if status_text:
            status_text.text("Building PDF document...")
        doc.build(elements)
        buffer.seek(0)
        return buffer
    else:
        if status_text:
            status_text.text("No labels were generated. Check if the Excel file has the expected columns.")
        return None

def generate_labels_from_excel_v2(df, progress_bar=None, status_text=None):
    """Generate labels using version 2 formatting."""
    
    buffer = io.BytesIO()
    
    # Set up key measurements
    part_no_height = 1.9 * cm
    desc_height = 2.1 * cm
    loc_height = 0.9 * cm

    # Identify column names
    cols = df.columns.tolist()
    df.columns = [col.upper() for col in df.columns]
    cols = df.columns.tolist()

    # Find columns
    part_no_col = next((col for col in cols if 'PART' in col and ('NO' in col or 'NUM' in col or '#' in col)),
                      next((col for col in cols if col in ['PARTNO', 'PART']), None))
    desc_col = next((col for col in cols if 'DESC' in col), None)
    loc_col = next((col for col in cols if 'LOC' in col or 'POS' in col), None)

    if not part_no_col:
        part_no_col = cols[0]
    if not desc_col:
        desc_col = cols[1] if len(cols) > 1 else part_no_col
    if not loc_col:
        loc_col = cols[2] if len(cols) > 2 else desc_col

    if status_text:
        status_text.text(f"Using columns: Part No: {part_no_col}, Description: {desc_col}, Location: {loc_col}")

    # Group parts by location
    df_grouped = df.groupby(loc_col)
    total_locations = len(df_grouped)

    doc = SimpleDocTemplate(buffer, pagesize=A4)
    elements = []

    MAX_LABELS_PER_PAGE = 4
    label_count = 0

    for i, (location, group) in enumerate(df_grouped):
        try:
            if progress_bar:
                progress_value = int((i / total_locations) * 100)
                progress_bar.progress(progress_value)
            
            if status_text:
                status_text.text(f"Processing location {i+1}/{total_locations}: {location}")

            parts = group.head(2)

            if len(parts) < 2:
                if len(parts) == 1:
                    part1 = parts.iloc[0]
                else:
                    continue
            else:
                part1 = parts.iloc[0]

            if label_count > 0 and label_count % MAX_LABELS_PER_PAGE == 0:
                elements.append(PageBreak())

            label_count += 1

            part_no = str(part1[part_no_col])
            desc = str(part1[desc_col])
            location_str = str(part1[loc_col])
            location_values = parse_location_string_v2(location_str)

            # Part table with enhanced formatting
            part_table = Table(
                [['Part No', format_part_no_v2(part_no)],
                 ['Description', format_description(desc)]],
                colWidths=[4*cm, 11*cm],
                rowHeights=[part_no_height, desc_height]
            )

            part_table.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('ALIGN', (0, 0), (0, -1), 'CENTER'),
                ('ALIGN', (1, 0), (1, 0), 'CENTER'),
                ('ALIGN', (1, 1), (1, -1), 'LEFT'),
                ('VALIGN', (0, 0), (0, 0), 'MIDDLE'),
                ('VALIGN', (1, 0), (1, 0), 'TOP'),
                ('VALIGN', (0, 1), (0, 1), 'MIDDLE'),
                ('VALIGN', (1, 1), (1, 1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 5),
                ('RIGHTPADDING', (0, 0), (-1, -1), 5),
                ('TOPPADDING', (1, 0), (1, 0), 10),
                ('BOTTOMPADDING', (1, 0), (1, 0), 5),
                ('FONTNAME', (0, 0), (0, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (0, -1), 16),
            ]))

            # Location table
            location_data = [['Part Location'] + location_values]
            location_widths = [4*cm]
            remaining_width = 11 * cm
            col_widths = [1.7, 2.9, 1.3, 1.2, 1.3, 1.3, 1.3]
            total_proportion = sum(col_widths)
            location_widths.extend([w * remaining_width / total_proportion for w in col_widths])

            location_table = Table(
                location_data,
                colWidths=location_widths,
                rowHeights=loc_height,
            )

            location_colors = [
                colors.HexColor('#E9967A'),
                colors.HexColor('#ADD8E6'),
                colors.HexColor('#90EE90'),
                colors.HexColor('#FFD700'),
                colors.HexColor('#ADD8E6'),
                colors.HexColor('#E9967A'),
                colors.HexColor('#90EE90')
            ]

            location_style = [
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (0, 0), 'TOP'),
                ('VALIGN', (1, 0), (-1, 0), 'TOP'),
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (0, 0), 16),
                ('FONTSIZE', (1, 0), (-1, -1), 16),
            ]

            for j, color in enumerate(location_colors):
                location_style.append(('BACKGROUND', (j+1, 0), (j+1, 0), color))

            location_table.setStyle(TableStyle(location_style))

            elements.append(part_table)
            elements.append(Spacer(1, 0.3 * cm))
            elements.append(location_table)
            elements.append(Spacer(1, 0.2 * cm))

        except Exception as e:
            if status_text:
                status_text.text(f"Error processing location {location}: {e}")
            continue

    if progress_bar:
        progress_bar.progress(100)

    if elements:
        if status_text:
            status_text.text("Building PDF document...")
        doc.build(elements)
        buffer.seek(0)
        return buffer
    else:
        if status_text:
            status_text.text("No labels were generated.")
        return None

def main():
    st.title("🏷️ Rack Label Generator")
    st.markdown(
        "<p style='font-size:18px; font-style:italic; margin-top:-10px; text-align:left;'>"
        "Designed and Developed by Agilomatrix</p>",
        unsafe_allow_html=True
    )

    st.markdown("---")

    # Sidebar for navigation
    st.sidebar.title("Label Generator Options")
    label_type = st.sidebar.selectbox(
        "Choose Label Type:",
        ["Enhanced Labels (v2)", "Standard Labels (v1)"]
    )

    # File upload
    uploaded_file = st.file_uploader(
        "Choose an Excel or CSV file",
        type=['xlsx', 'xls', 'csv'],
        help="Upload your Excel or CSV file containing part information"
    )

    if uploaded_file is not None:
        try:
            # Read the file
            if uploaded_file.name.lower().endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)

            st.success(f"✅ File loaded successfully! Found {len(df)} rows and {len(df.columns)} columns.")
            
            # Display file info
            with st.expander("📊 File Information", expanded=False):
                st.write("**Columns found:**", df.columns.tolist())
                st.write("**First few rows:**")
                st.dataframe(df.head(3))

            # Generate PDF button
            if st.button("🚀 Generate PDF Labels", type="primary"):
                
                # Create progress indicators
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                try:
                    # Generate PDF based on selected type
                    if label_type == "Single Part":
                        pdf_buffer = generate_labels_from_excel_v2(df, progress_bar, status_text)
                        filename = "singlepart_labels.pdf"
                    else:
                        pdf_buffer = generate_labels_from_excel_v1(df, progress_bar, status_text)
                        filename = "multiplepart_labels.pdf"

                    if pdf_buffer:
                        status_text.text("✅ PDF generated successfully!")
                        
                        # Provide download button
                        st.download_button(
                            label="📥 Download PDF",
                            data=pdf_buffer.getvalue(),
                            file_name=filename,
                            mime="application/pdf",
                            type="primary"
                        )
                        
                        st.success("🎉 Your PDF is ready for download!")
                        
                    else:
                        st.error("❌ Failed to generate PDF. Please check your file format and data.")
                        
                except Exception as e:
                    st.error(f"❌ An error occurred: {str(e)}")
                    st.info("Please ensure your file has the expected columns (Part No, Description, Location)")

        except Exception as e:
            st.error(f"❌ Error reading file: {str(e)}")
            st.info("Please ensure you've uploaded a valid Excel or CSV file.")

    else:
        # Show instructions when no file is uploaded
        st.info("👆 Please upload an Excel or CSV file to get started")
        
        with st.expander("📋 File Format Requirements", expanded=True):
            st.markdown("""
            **Your file should contain the following columns:**
            - **Part Number** (column names like: 'Part No', 'Part Number', 'PartNo', etc.)
            - **Description** (column names like: 'Description', 'Desc', etc.)
            - **Location** (column names like: 'Location', 'Loc', 'Position', 'Pos', etc.)
            
            **Supported file formats:**
            - Excel files (.xlsx, .xls)
            - CSV files (.csv)
            
            **Label Types:**
            - **Single Part**: One part assigned to a single location  
            - **Multiple Parts**: Multiple parts assigned to a single location
            """)

if __name__ == "__main__":
    main()
