#!/usr/bin/env python3
"""
Script to create a simple application icon for Word Document Merger
"""
import os
from PIL import Image, ImageDraw, ImageFont

def create_simple_icon():
    """Create a simple icon for the application"""
    # Create a 256x256 image with transparent background
    icon_size = 256
    icon = Image.new('RGBA', (icon_size, icon_size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(icon)
    
    # Draw a rounded rectangle for document shape
    border_radius = 20
    document_color = (56, 118, 187)  # Blue color for document
    
    # Main document shape
    draw.rounded_rectangle(
        [(30, 20), (icon_size - 30, icon_size - 20)],
        radius=border_radius,
        fill=document_color
    )
    
    # White area to represent paper
    draw.rounded_rectangle(
        [(50, 40), (icon_size - 50, icon_size - 40)],
        radius=border_radius,
        fill='white'
    )
    
    # Add some horizontal lines to represent text
    line_color = (200, 200, 200)
    line_y_positions = [80, 110, 140, 170, 200]
    for y in line_y_positions:
        draw.line([(70, y), (icon_size - 70, y)], fill=line_color, width=8)
    
    # Add a merge symbol
    merge_color = (216, 67, 57)  # Red color for merge
    draw.polygon(
        [(icon_size//2, 40), (icon_size//2 + 30, 80), (icon_size//2 - 30, 80)],
        fill=merge_color
    )
    
    # Save as icon file
    icon.save('app_icon.png')
    
    try:
        # Try to convert to ICO format if pillow supports it
        icon.save('app_icon.ico', format='ICO')
        print("Icon created: app_icon.ico")
        return True
    except Exception as e:
        print(f"Could not create ICO file: {e}")
        print("Using PNG format instead.")
        return False

if __name__ == "__main__":
    create_simple_icon()
    print("Icon file created: app_icon.png") 