import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.patches import Wedge
from matplotlib.colors import Normalize, LinearSegmentedColormap
from matplotlib.cm import ScalarMappable
from matplotlib.transforms import Affine2D

# Function to read data from Excel file
def read_data_from_excel(file_path):
    df = pd.read_excel(file_path)
    categories = df.iloc[:, 0].tolist()  # First column as categories
    sectors = df.columns[1:].tolist()    # Remaining columns as sectors
    data = df.iloc[:, 1:].values         # Data for sectors
    return sectors, categories, data

# Read data from Excel file
file_path = 'C:/Danial/Graphs/sample 1.xlsx'
sectors, categories, data = read_data_from_excel(file_path)

# Normalize data
data = np.array(data)

# Define custom colors
colors = ['red', '#FFAAAA', 'white', '#AAAAFF', 'blue']
cmap = LinearSegmentedColormap.from_list('custom_cmap', colors)

# Normalize the data for the colormap
norm = Normalize(vmin=data.min(), vmax=data.max())
color_mapper = ScalarMappable(norm=norm, cmap=cmap)

fig, ax = plt.subplots(figsize=(10, 10), subplot_kw={'aspect': 'equal'})

# Define parameters
gap_degrees = 35
angle_per_sector = (360 - gap_degrees) / len(sectors)
start_angle = gap_degrees / 2

inner_radius = 1.8  # Radius of the void space in the center
total_width = 1.9  # Fixed total width for the circular graph
sector_width = total_width / len(categories)

# Draw sectors with stacked bars
for i, sector_data in enumerate(data.T):  # Transpose data to iterate by sector
    sector_start_angle = start_angle + i * angle_per_sector
    sector_end_angle = sector_start_angle + angle_per_sector
    bottom = inner_radius

    for value in sector_data:
        color = color_mapper.to_rgba(value)
        wedge = Wedge(center=(0, 0), r=bottom + sector_width, theta1=sector_start_angle, theta2=sector_end_angle, width=sector_width, facecolor=color, edgecolor='black')
        ax.add_patch(wedge)
        bottom += sector_width

# Add an extra small sector to close the gap
extra_sector_angle = gap_degrees
extra_sector_start_angle = start_angle + len(sectors) * angle_per_sector
extra_sector_end_angle = extra_sector_start_angle + extra_sector_angle
bottom = inner_radius

for category in categories:
    wedge = Wedge(center=(0, 0), r=bottom + sector_width, theta1=extra_sector_start_angle, theta2=extra_sector_end_angle, width=sector_width, facecolor='white', edgecolor='none')
    ax.add_patch(wedge)
    bottom += sector_width

# Manually adjustable font size for stack bar names
stack_bar_font_size = 5  # Adjust this value to change the font size

# Calculate the font size based on the average length of category names
font_size = stack_bar_font_size

# Define the distance to move the text boxes away from the starting edge
text_box_distance = 0.1

# Add category names in text boxes outside the starting edge of the first sector
bottom = inner_radius
for i, category in enumerate(categories):
    radius = inner_radius + (i + 0.5) * sector_width
    text_radius = radius
    angle = start_angle
    
    x = text_radius * np.cos(np.radians(angle)) + text_box_distance
    y = text_radius * np.sin(np.radians(angle))
    
    trans_angle = Affine2D().rotate_deg_around(x, y, angle - 90)
    
    ax.text(x, y, category, ha='left', va='center', fontsize=font_size, rotation=angle - 90, rotation_mode='anchor', bbox=dict(facecolor='white', edgecolor='none', pad=1), transform=trans_angle + ax.transData)

# Add labels outside each sector along radial axis
label_distance = inner_radius + total_width + 0.5
for i, sector in enumerate(sectors):
    angle = start_angle + i * angle_per_sector + angle_per_sector / 2
    x = label_distance * np.cos(np.radians(angle))
    y = label_distance * np.sin(np.radians(angle))
    if 90 < angle < 270:
        rotation = angle - 180
        ha = 'right'
    else:
        rotation = angle
        ha = 'left'
    ax.text(x, y, sector, ha=ha, va='center', fontsize=12, rotation=rotation, rotation_mode='anchor')

# Create and add legend in the center void space
cbar_ax = fig.add_axes([0.4825, 0.41, 0.030, 0.15])  # [left, bottom, width, height]
cbar = fig.colorbar(color_mapper, cax=cbar_ax, orientation='vertical', aspect=6)

cbar.ax.set_title('Pvalue', pad=10, ha='center')

# Set limits and hide axes
ax.set_xlim(-5, 5)
ax.set_ylim(-5, 5)
ax.axis('off')

plt.show()
