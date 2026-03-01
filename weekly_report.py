import openpyxl
from openpyxl.chart import BarChart, Reference

# Load the workbook and select the active worksheet
workbook = openpyxl.load_workbook('weekly_report.xlsx')
worksheet = workbook.active

# Create a 2-D Clustered Bar Chart
chart = BarChart()
chart.type = 'col'
chart.style = 1
chart.title = 'Weekly Report Chart'
chart.y_axis.title = 'Values'
chart.x_axis.title = 'Categories'

# Define data for the chart
data = Reference(worksheet, min_col=1, min_row=3, max_col=2, max_row=4)
chart.add_data(data, titles_from_data=True)

# Set the colors for the columns
chart.series[0].graphicalProperties.solidFill = "FF0000FF"  # Blue
chart.series[1].graphicalProperties.solidFill = "FFFFA500"  # Orange

# Customize the chart
chart.height = 10  # Set height of the chart
chart.width = 20   # Set width of the chart
chart.style = 2    # Use a specified style
chart.showgridlines = False  # No gridlines
chart.datalabels.showVal = True  # Show values only

# Add the chart to the sheet
worksheet.add_chart(chart, 'D5')

# Save the workbook
workbook.save('weekly_report.xlsx')

# Close the workbook
workbook.close()