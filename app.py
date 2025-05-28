import streamlit as st
import pandas as pd
import calendar
from io import BytesIO
from datetime import datetime
import xlsxwriter

st.title("ALS Field Services Timesheet")

sheet_type = st.selectbox("Select Timesheet Type", ["Personal Timesheet", "Equipment Timesheet"])

month = st.selectbox("Select Month", list(calendar.month_name)[1:])
field_name = st.text_input("Field Name")
well_name = st.text_input("Well Name")
customer_name = st.text_input("Client Name")
start_date = st.date_input("Starting Date")
end_date = st.date_input("Ending Date")

# === Dynamic Input Fields ===
technician_names = []
equipment_names = []

if sheet_type == "Personal Timesheet":
    if 'tech_count' not in st.session_state or sheet_type != st.session_state.get("last_sheet_type"):
        st.session_state.tech_count = 1
        st.session_state.last_sheet_type = sheet_type

    if st.button("+ Add Technician"):
        st.session_state.tech_count += 1

    for i in range(st.session_state.tech_count):
        label = "Installation & Commissioning Supervisor" if i == 0 else f"Technician {i + 1}"
        technician_names.append(st.text_input(label, key=f"tech_{i}"))

elif sheet_type == "Equipment Timesheet":
    equipment_options = [
        'BOP Can for 7" Rams',
        'BOP Can for 9 5/8" Rams',
        'ESP Welltest Toolbox Container c/w lifting tools',
        'ESP String (DHE) 300-1200 BPD',
        'ESP String (DHE) 1100-2500 BPD',
        'ESP String (DHE) 2300-4500 BPD',
        'Y-Tool Set: For 7" or 9 5/8" Casing',
        "Phoenix Multisensor '1', 257 deg F rated",
        'Generator',
        'Other (specify...)'
    ]

    if 'equip_count' not in st.session_state or sheet_type != st.session_state.get("last_sheet_type"):
        st.session_state.equip_count = 1
        st.session_state.last_sheet_type = sheet_type

    if st.button("+ Add Equipment"):
        if st.session_state.equip_count < 7:
            st.session_state.equip_count += 1

    for i in range(st.session_state.equip_count):
        selected_option = st.selectbox(f"Select Equipment {i + 1}", equipment_options, key=f"equip_{i}")
        if selected_option == 'Other (specify...)':
            custom_equipment = st.text_input(f"Enter Equipment {i + 1} name", key=f"custom_equip_{i}")
            equipment_names.append(custom_equipment)
        else:
            equipment_names.append(selected_option)

SLB_Representative = st.text_input("SLB Representative")

if st.button("Generate Timesheet"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        pd.DataFrame().to_excel(writer, index=False, sheet_name='Sheet1')
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        try:
            worksheet.insert_image('A1', 'logo.png', {
                'x_offset': 5, 'y_offset': 5,
                'x_scale': 0.5, 'y_scale': 0.5
            })
        except:
            pass

        title_fmt = workbook.add_format({'bold': True, 'font_size': 24, 'align': 'center', 'valign': 'vcenter'})
        header_fmt = workbook.add_format({'bold': True, 'align': 'center', 'border': 1})
        center = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
        bold_center = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
        subhead_fmt = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter'})
        dark_grey_fmt = workbook.add_format({'bg_color': '#595959'})

        year = start_date.year
        month_index = list(calendar.month_name).index(month)
        days_in_month = calendar.monthrange(year, month_index)[1]

        base_row = 3
        data_rows = len(technician_names) if sheet_type == "Personal Timesheet" else len(equipment_names)
        shift = data_rows - 2

        title_text = 'Field Services Timesheet' if sheet_type == "Personal Timesheet" else 'SLB Equipment Time Sheet'
        worksheet.merge_range('B2:AI2', title_text, title_fmt)
        worksheet.merge_range(f'B{10 + shift}:AI{10 + shift}',
                              'The above certifies and represents the number of days that lw Services have been provided at location',
                              subhead_fmt)

        worksheet.write(base_row, 0, '#', header_fmt)
        worksheet.write(base_row, 1, 'Month', header_fmt)
        worksheet.write(base_row, 2, 'ESP Crew' if sheet_type == "Personal Timesheet" else 'Equipment', header_fmt)
        worksheet.write(base_row, 3, 'Field Name', header_fmt)

        for i in range(1, days_in_month + 1):
            worksheet.write(base_row, 3 + i, str(i), header_fmt)
        worksheet.write(base_row, 4 + days_in_month, 'Total', header_fmt)

        work_days = list(range(start_date.day, end_date.day + 1))
        total_column_values = []

        for i in range(data_rows):
            row = base_row + 1 + i
            worksheet.write(row, 0, i + 1, center)

            if i == 0:
                if data_rows == 1:
                    worksheet.write(row, 1, month, center)
                    worksheet.write(row, 3, field_name, center)
                else:
                    worksheet.merge_range(row, 1, row + data_rows - 1, 1, month, center)
                    worksheet.merge_range(row, 3, row + data_rows - 1, 3, field_name, center)

            name = technician_names[i] if sheet_type == "Personal Timesheet" else equipment_names[i]
            worksheet.write(row, 2, name, center)

            workday_count = 0
            for day in range(1, days_in_month + 1):
                col = 3 + day
                if name.strip() and day in work_days:
                    worksheet.write(row, col, well_name, center)
                    workday_count += 1
                else:
                    worksheet.write(row, col, '', dark_grey_fmt)

            worksheet.write(row, 4 + days_in_month, workday_count, center)
            total_column_values.append(workday_count)

        date_label_row = base_row + data_rows + 1
        worksheet.write(date_label_row, 2, 'Starting Date', header_fmt)
        worksheet.write(date_label_row, 3, start_date.strftime("%Y-%m-%d"), center)
        worksheet.write(date_label_row + 1, 2, 'Ending Date', header_fmt)
        worksheet.write(date_label_row + 1, 3, end_date.strftime("%Y-%m-%d"), center)

        base_info_row = 12 + shift
        worksheet.merge_range(f'B{base_info_row}:D{base_info_row+3}', 'SLB Representative', bold_center)
        worksheet.merge_range(f'E{base_info_row}:J{base_info_row+3}', SLB_Representative, center)
        worksheet.merge_range(f'B{base_info_row+4}:D{base_info_row+5}', 'Date', bold_center)
        worksheet.merge_range(f'E{base_info_row+4}:J{base_info_row+5}', end_date.strftime("%Y-%m-%d"), center)
        worksheet.merge_range(f'B{base_info_row+6}:D{base_info_row+7}', 'Signature', bold_center)
        worksheet.merge_range(f'E{base_info_row+6}:J{base_info_row+7}', '', center)

        worksheet.merge_range(f'S{base_info_row}:X{base_info_row+1}', 'Client Name', bold_center)
        worksheet.merge_range(f'Y{base_info_row}:AF{base_info_row+1}', customer_name, center)
        worksheet.merge_range(f'S{base_info_row+2}:X{base_info_row+3}', 'Field Name', bold_center)
        worksheet.merge_range(f'Y{base_info_row+2}:AF{base_info_row+3}', field_name, center)
        worksheet.merge_range(f'S{base_info_row+4}:X{base_info_row+5}', 'Client Representative', bold_center)
        worksheet.merge_range(f'Y{base_info_row+4}:AF{base_info_row+5}', '', center)
        worksheet.merge_range(f'S{base_info_row+6}:X{base_info_row+7}', 'Client Rep. Signature', bold_center)
        worksheet.merge_range(f'Y{base_info_row+6}:AF{base_info_row+7}', '', center)
        worksheet.merge_range(f'S{base_info_row+8}:X{base_info_row+9}', 'Date', bold_center)
        worksheet.merge_range(f'Y{base_info_row+8}:AF{base_info_row+9}', '', center)

        frame_top = 0
        frame_bottom = base_info_row + 9
        frame_left = 0
        frame_right = 36  # Column AI

        # Format for specific sides
        top_border = workbook.add_format({'bottom': 2})
        bottom_border = workbook.add_format({'top': 2})
        left_border = workbook.add_format({'left': 2})
        right_border = workbook.add_format({'left': 2})
        top_left_corner = workbook.add_format({'bottom': 2, 'left': 2})
        #top_right_corner = workbook.add_format({'bottom': 2, 'right': 2})
        bottom_left_corner = workbook.add_format({'top': 2, 'left': 2})
        #bottom_right_corner = workbook.add_format({'top': 1, 'right': 1})

# Top edge
        worksheet.write_blank(frame_top, frame_left, None, top_left_corner)
        for col in range(frame_left, frame_right):
            worksheet.write_blank(frame_top, col, None, top_border)
        worksheet.write_blank(frame_top, frame_right, None)

# Bottom edge
        worksheet.write_blank(frame_bottom, frame_left, None, bottom_left_corner)
        for col in range(frame_left, frame_right):
            worksheet.write_blank(frame_bottom, col, None, bottom_border)
        worksheet.write_blank(frame_bottom, frame_right, None)

# Left edge (excluding corners already written)
        for row in range(frame_top + 1, frame_bottom):
            worksheet.write_blank(row, frame_left, None, left_border)

# Right edge (excluding corners already written)
        for row in range(frame_top + 1, frame_bottom):
            worksheet.write_blank(row, frame_right, None, right_border)

        # === Column widths ===
        worksheet.set_column('A:A', 5)
        worksheet.set_column('B:B', 8)
        worksheet.set_column('C:C', 39)
        worksheet.set_column('D:D', 12)
        worksheet.set_column('E:AI', 5)
        worksheet.set_column('AJ:AJ', 5)
        worksheet.set_column('S:AF', 5)

    st.success("✅ Excel file created!")

    st.download_button(
        label="📥 Download Timesheet",
        data=output.getvalue(),
        file_name=f"{sheet_type.replace(' ', '_')}_{well_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
