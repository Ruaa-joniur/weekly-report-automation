import matplotlib.pyplot as plt
import pandas as pd

# Sample data from Excel cells A3, B3, A4, B4
labels = ['Label 1', 'Label 2']
values = [
    pd.read_excel('file.xlsx', sheet_name='Sheet1', usecols='A', skiprows=2).iloc[0, 0],  # A3
    pd.read_excel('file.xlsx', sheet_name='Sheet1', usecols='B', skiprows=2).iloc[0, 0]   # B3
]

# Create a new figure for the plot
fig, ax = plt.subplots()

# Create a clustered bar chart
bars = ax.bar(labels, values, color=['#FF5733', '#33FF57'], edgecolor='black')  # Different colors

# Add data labels
for bar in bars:
    yval = bar.get_height()
    ax.text(bar.get_x() + bar.get_width()/2, yval, round(yval, 2), ha='center', va='bottom')

# Customize the chart
ax.set_title('Chart 1 Title')  # Title
ax.grid(False)  # No major gridlines
ax.set_ylim(0, max(values) * 1.1)  # Give some space above the bars

# Adjust plot area
plt.subplots_adjust(left=0.1, right=0.9, top=0.9, bottom=0.1)  # Smaller plot area for visible title

# Display the plot
plt.show()