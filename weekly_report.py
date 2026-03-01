import openpyxl
from openpyxl.chart import LineChart, Reference

# Load the workbook
wb = openpyxl.load_workbook('weekly_report.xlsx')
# Select the active worksheet
ws = wb.active

# Prepare data and create a line chart
chart = LineChart()
chart.title = 'Weekly Report Data'

# Define data for the chart
data = Reference(ws, min_col=2, min_row=1, max_col=3, max_row=10)
chart.add_data(data, titles_from_data=True)

# Add chart to the worksheet
ws.add_chart(chart, "E5")

# Save the workbook
wb.save('weekly_report.xlsx')