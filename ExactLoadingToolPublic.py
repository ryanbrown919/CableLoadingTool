"""
Exact Loading Tool

This tool is to be used in conjuction with DEAT (4.2.4) and SQLite Import Tool.
This tool allows the user to assign exisitng loading, new loading and MPT tails to midspans.

Author: Ryan Brown
Email: Ryan.Brown4@ledcor.com | ryanmacbrown@gmail.com

Date Created: November 15th 2023
Last Modified: December 6th 2023

"""

import tkinter as tk
import pandas as pd
import utm
from tkinter import filedialog
from tkinter import ttk
import math
from openpyxl import load_workbook
import openpyxl
import os
import copy
from tkinter import *
import time

#Updated this if program is ever modified
version = ' Public Test'

# Load data from an Excel file with pole IDs read as strings
filePath = filedialog.askopenfilename() # Triggers the file dialog window
column_Types = {'Pole_': str}# Ensures the poleIDs are interprested as strings and not integers
SCD = pd.read_excel(filePath, sheet_name='SCD', dtype=str)  # Assign SCD worksheet data to an dataframe
PCD = pd.read_excel(filePath, sheet_name='PCD', dtype=column_Types) # Assign PCD worksheet data to a dataframe

# Correct column name if used for BAU
if 'VERIFIED' in PCD.columns:
    PCD.rename(columns={'VERIFIED': 'Verified'}, inplace=True)
    
# Declaring key data structures. these will be maniupulated throughout the program
# Notation: Data Structure | Data Type / Format | Description
spans = {}            # Dict  | Key = SpanID, Value = (FPID, LPID) | Dictionary containing spans and the poles they connect to
poles = {}            # Dict  | Key = PoleID, Value = (EAST, NORTH) | Dictionary containing poles and their locations
polesOriginal = {}    # Dict  | Key = PoleID, Value = (EAST, NORTH) | Dictionary containing poles and their locations
mptPaths = {}         # Dict  | Key = MPT PoleID, Value = [SpanIDs] | Dictionary containing poles with MPTs and the spanIDs between splice and MPT
mptCount = {}         # Dict  | Key = SpanID, Value = MPT Count | Dictionary containing spans and the number of MPT tails added in new loading
poleType = {}         # Dict  | Key = PoleID, Value = 'Splice/MPT' | Dictionary containing poles and a string describing if they are splice poles or MPTs poles. if MPT, formatting resembles 'MPT_<spliceID>'
spliceMPTs = {}       # Dict  | Key = Splice PoleID, Value = [MPT PoleIDs] | Dictionary containing poles with splices and the MPTs connected to each splice.
mptSpans = []         # List  | str | Temporary list with the spans between a splice and an MPT
mptPoles = []         # List  | str | List of poles with MPTs on them
splicePoles = []      # List  | str | List of poles with a splice on them
selectedSpans = []    # List  | str | Temporary list of spans selected for cable loading
selectedCables = []   # List  | str | Temporary list of cables selected for loading 
totalPoles = []       # List  | str | List of all possible poles
validPoles = []       # List  | str | List of all possible poles
connectedSpans = []   # List  | str | Temporary list of spans connected to a given pole
connectedPoles = []   # List  | str | Temporary list of poles adjacent to a given pole

# Declaring some variables with placeholder values
lastPoleID = ''       # Str   | Temporary string to track the most most recently selected pole
spliceID = ''         # Str   | Temporary string to track the poleID of a splice
zoomFactor = 1        # Int   | USE THIS FOR ZOOM?
scale = 1             # Float | Used for scaling elements as user zooms in and out of canvas
zoom_level = 1        # Int   | Used for determining the font size of the text when zooming in and out 

# Declaring some variables with logical values
exactLoadingExists = False  # Bool  | This is used to trackif exactLoading.xlsx is present in the same folder as DEAT, so past data can be reloaded if it is
firstClickTrue = True       # Bool  | This is used to track if a right click is a first click or not. This dictates starting the span selections or placing a splice
amDrawing = False           # Bool  | This is used to lock the zoom level when drawing. Becuase of the way canvas scaling works if the user zooms while drawing the line swill no longer match existing spans
drawingMPTs = False         # Bool  | This is used to determine if the 'Place Splice/MPTs' button is currently toggled

# Initialize some static variables with initial values
colourDict = {'Cable 1': '#696969','Cable 2': '#a9a9a9','Cable 3': '#7fffd4','Cable 4': '#2f4f4f','Cable 5': '#556b2f','Cable 6': '#6b8e23','Cable 7': '#a0522d','Cable 8': '#2e8b57','Cable 9': '#228b22','Cable 10': '#800000','Cable 11': '#191970','Cable 12': '#006400','Cable 13': '#708090','Cable 14': '#808000','Cable 15': '#483d8b','Cable 16': '#b22222','Cable 17': '#5f9ea0','Cable 18': '#3cb371','Cable 19': '#bc8f8f','Cable 20': '#663399','Cable 21': '#b8860b','Cable 22': '#bdb76b','Cable 23': '#008b8b','Cable 24': '#cd853f','Cable 25': '#4682b4','Cable 82': '#000080','Cable 26': '#d2691e','Cable 27': '#9acd32','Cable 28': '#20b2aa','Cable 29': '#cd5c5c','Cable 30': '#4b0082','Cable 31': '#32cd32','Cable 32': '#daa520','Cable 33': '#7f007f','Cable 34': '#8fbc8f','Cable 35': '#b03060','Cable 36': '#66cdaa','Cable 37': '#9932cc','Cable 38': '#ff0000','Cable 39': '#ff4500','Cable 40': '#ff8c00','Cable 41': '#ffa500','Cable 42': '#ffd700','Cable 43': '#ffff00','Cable 44': '#c71585','Cable 45': '#0000cd','Cable 46': '#deb887','Cable 47': '#40e0d0','Cable 48': '#7fff00','Cable 49': '#00ff00','Cable 50': '#ba55d3','Cable 51': '#00fa9a','Cable 52': '#00ff7f','Cable 53': '#4169e1','Cable 54': '#e9967a','Cable 55': '#dc143c','Cable 56': '#00ffff','Cable 57': '#00bfff','Cable 58': '#9370db','Cable 59': '#0000ff','Cable 60': '#a020f0','Cable 61': '#f08080','Cable 62': '#adff2f','Cable 63': '#ff6347','Cable 64': '#d8bfd8','Cable 65': '#b0c4de','Cable 66': '#ff7f50','Cable 67': '#ff00ff','Cable 68': '#1e90ff','Cable 69': '#db7093','Cable 70': '#f0e68c','Cable 71': '#eee8aa','Cable 72': '#ffff54','Cable 73': '#6495ed','Cable 74': '#dda0dd','Cable 75': '#87ceeb','Cable 76': '#ff1493','Cable 77': '#7b68ee','Cable 78': '#afeeee','Cable 79': '#ee82ee','Cable 80': '#98fb98', 'Cable 81': '#4287f5'}
allCableOptions = ['Cable 1','Cable 2','Cable 3','Cable 4','Cable 5','Cable 6','Cable 7','Cable 8','Cable 9','Cable 10','Cable 11','Cable 12','Cable 13','Cable 14','Cable 15','Cable 16','Cable 17','Cable 18','Cable 19','Cable 20','Cable 21','Cable 22','Cable 23','Cable 24','Cable 25','Cable 26','Cable 27','Cable 28','Cable 29','Cable 30','Cable 31','Cable 32','Cable 33','Cable 34','Cable 35','Cable 36','Cable 37','Cable 38','Cable 39','Cable 40','Cable 41','Cable 42','Cable 43','Cable 44','Cable 45','Cable 46','Cable 47','Cable 48','Cable 49','Cable 50','Cable 51','Cable 52','Cable 53','Cable 54','Cable 55','Cable 56','Cable 57','Cable 58','Cable 59','Cable 60','Cable 61','Cable 62','Cable 63','Cable 64','Cable 65','Cable 66','Cable 67','Cable 68','Cable 69','Cable 70','Cable 71','Cable 72','Cable 73','Cable 74','Cable 75','Cable 76','Cable 77','Cable 78','Cable 79','Cable 80','Cable 81', 'Cable 82']
options = ['Cable 1', 'Cable 2', 'Cable 3', 'Cable 40', 'Cable 53', 'Cable 65', 'Cable 71', 'Cable 8', 'Cable 34', 'Cable 10', 'Cable 66', 'Cable 12', 'Cable 13', 'Cable 42', 'Cable 15']
utm_eastings = PCD['EAST']  
utm_northings = PCD['NORTH']
totalPoles = PCD['Pole_'].tolist()
reference_utm_northing = min(utm_northings) # Used as a datum for pixel coordinates
span_data = SCD[['SPN_N', 'FPID', 'LPID', 'Verified']]

# Custom settings for if exactLoading.xlsx file is open at code execution
max_attempts = 5
attempt_delay_seconds = 5
attempts = 0

while attempts < max_attempts:

    try:
        # Load previous data if file is found
        exactLoading = pd.read_excel(os.path.dirname(filePath)+'/'+'exactLoading.xlsx', sheet_name = 'Exact Span Loading', dtype=column_Types)
        exactLoadingExists = True

        if 'Existing Loading' not in SCD.columns:
            # Initialize the columnds in PCD and SCD sheets
            SCD['Existing Loading'] = None
            SCD['New Loading'] = None
            SCD['# MPTs'] = 0
            PCD['Splice/MPT'] = None
            PCD['Path'] = None

        #Extract information from exactLoading.xlsx to PCD and SCD dataframes for easier access
        for index, row in exactLoading.iterrows():
            if not pd.isna(row['Existing Loading']):
                SCD.at[index, 'Existing Loading'] = row['Existing Loading']
            if not pd.isna(row['New Loading']):
                SCD.at[index, 'New Loading'] = row['New Loading']
            if not pd.isna(row['# MPTs']):
                SCD.at[index, '# MPTs'] = row['# MPTs']
            if not pd.isna(row['Splice/MPT']):
                PCD.at[index, 'Splice/MPT'] = row['Splice/MPT']
                poleType[str(row['Pole_'])] = row['Splice/MPT']
            if not pd.isna(row['Path']):
                PCD.at[index, 'Path'] = row['Path']
            if not pd.isna(row['Presets']):
                options[index] = row['Presets']
        break

    # If exactLoading.xlsx is open, this error occurs. Warning will be displayed to user in terminal window, and will check every 5 seconds for it being closed 
    except PermissionError:
        print(f"Please close exactLoading.xlsx, warning {attempts+1}/5")
        attempts += 1
        time.sleep(attempt_delay_seconds)

        
    # If the file does not exist in the first place, or some other error occurs, this will let the program continue
    except FileNotFoundError:
        SCD['Existing Loading'] = None
        SCD['New Loading'] = None
        SCD['# MPTs'] = 0
        PCD['Splice/MPT'] = None
        PCD['Path'] = None
   
        break
else: 
    # Code to execute if all attempts fail. Will terminate the program
    print(f"Unable to open the file after {max_attempts} attempts. Exiting program.")
    exit()

def convert_utm_to_pixel(utm_easting, utm_northing):
    """
    Convert UTM easting and northing to relative pixels on GUI canvas

    Args:
    float: UTM easting
    float: UTM northing

    Returns:
    float: x coordinate in pixel value
    float: y coordinate in pixel value
    """

    # Change these scale values to make it appear more zoomed in or out 
    scale_x = 2
    scale_y = 2
    pixel_x = (utm_easting - min(utm_eastings)) * scale_x
    pixel_y = (reference_utm_northing - utm_northing) * scale_y
    return (pixel_x), (pixel_y)

# Initializes some data in a format more applicable than previous declarations
for index, row in PCD.iterrows():
    poleType[str(row['Pole_'])] = ''
    poles.update({str(row['Pole_']): (convert_utm_to_pixel(row['EAST'], row['NORTH']))})

def on_closing():
    """
    Saves data to exactLoading.xlsx
    
    """
    global filePath
    global options
    excel_file = filePath
    
    if placeSpliceState.get():
        place_splice()
        
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title='Exact Span Loading'
    sheet['A1'] = 'SPN_N'
    sheet['B1'] = 'Existing Loading'
    sheet['C1'] = 'New Loading'
    sheet['D1'] = '# MPTs'
    sheet['E1'] = 'Pole_'
    sheet['F1'] = 'Splice/MPT'
    sheet['G1'] = 'Path'
    sheet['Z1'] = 'Presets'
    for index, item in enumerate(options):
        sheet[f'Z{index+2}'] = options[index]
        
    for index, row in SCD.iterrows():
        sheet[f'A{index+2}'] = row['SPN_N']
        sheet[f'B{index+2}'] = row['Existing Loading']
        sheet[f'C{index+2}'] = row['New Loading']
        sheet[f'D{index+2}'] = row['# MPTs']
        
    for index, row in PCD.iterrows():
        sheet[f'E{index+2}'] = row['Pole_']
        sheet[f'F{index+2}'] = PCD.at[index, 'Splice/MPT']
        sheet[f'G{index+2}'] = PCD.at[index, 'Path']
        
    workbook.save(filename=os.path.dirname(filePath)+'/'+'exactLoading.xlsx')
    root.destroy()


def on_canvas_click(event):
    """
    Event triggers left mouse button is clicked
    """
    canvas.scan_mark(event.x, event.y)


def on_canvas_release(event):
    """
    Pans the canvas on user mouse click drag
    """
    canvas.scan_dragto(event.x, event.y, gain=1)

def on_canvas_scroll(event):
    """
    Used to zoom in and out of canvas with scrool wheel, currently mouse wheel not bound as it disables span selection and recreation

    """
    x, y = canvas.canvasx(event.x), canvas.canvasy(event.y)
    global scale
    global zoomFactor
    scale = 1.1 if event.delta > 0 else 0.9  # Adjust the scaling factor as needed
    canvas.scale("all", x, y, scale, scale)
    zoomFactor *= scale
    updatePoles(zoomFactor)
    
def updatePoles(zoom):
    """
    Updates pole pixel locations based on zoom scaling 

    Args:
    float: zoom factor
    """
    global scale
    for poleID, values in polesOriginal.items():
        originalX, originalY = values
        scaledX = originalX * zoom
        scaledY = originalY * zoom
        poles[poleID] = (scaledX, scaledY)

def draw_elements():
    """
    Draws initial poles and spans on canvas

    """
    global validPoles
    zoom_level = canvas.canvasx(1) - canvas.canvasx(0)  # Get the current zoom level
    canvas.delete("elements")  # Clear previous elements
    
    # Draw spans
    for index, row in span_data.iterrows():
        fpid = str(row['FPID'])
        lpid = str(row['LPID'])

        # If the pole is set to ignore, it will appear with custom appearance
        if row['Verified'] != 'Ignore': 
            spans.update({row['SPN_N']: (fpid, lpid)})
        

        # Find the coordinates of the FPID and LPID locations
        if fpid in totalPoles and lpid in totalPoles:

            fpid_index = totalPoles.index(fpid)
            lpid_index = totalPoles.index(lpid)

            x1, y1 = convert_utm_to_pixel(utm_eastings[fpid_index], utm_northings[fpid_index])
            x2, y2 = convert_utm_to_pixel(utm_eastings[lpid_index], utm_northings[lpid_index])


            if row['Verified'] == 'Ignore': 
                new_line = canvas.create_line(x1, y1, x2, y2, fill="black", width=2, dash=(2, 1), tags=str(row['SPN_N'])+"_Init")  # Display ignore spans with a dash 
                canvas.tag_lower(new_line)
            else:
                new_line = canvas.create_line(x1, y1, x2, y2, fill="black", width=2, tags=str(row['SPN_N'])+"_Init")  # Display normal spans
                canvas.tag_lower(new_line)
    
    # Draw poles
    for i, row in PCD.iterrows():
        
        utm_easting = utm_eastings[i]
        utm_northing = utm_northings[i]
        location_id = totalPoles[i]
        pixel_x, pixel_y = convert_utm_to_pixel(utm_easting, utm_northing)

        # Adjust the font size based on the zoom level
        font_size = max(int(8 / zoom_level), 1)
        if row['Verified'] == 'Ignore':
            canvas.create_oval(pixel_x - 5, pixel_y - 5, pixel_x + 5, pixel_y + 5, fill="grey", tags="elements")  # Display locations as blue dots
            canvas.create_text(pixel_x, pixel_y + 10, text=location_id, font=("Helvetica", font_size), tags="elements")
        elif row['Code'] == 'Strain Pole': 
            canvas.create_oval(pixel_x - 5, pixel_y - 5, pixel_x + 5, pixel_y + 5, fill="red", tags="elements")  # Display locations as blue dots
            canvas.create_text(pixel_x, pixel_y + 10, text=location_id, font=("Helvetica", font_size), tags="elements")
            validPoles.append(location_id)

        else: 
            canvas.create_oval(pixel_x - 5, pixel_y - 5, pixel_x + 5, pixel_y + 5, fill="blue", tags="elements")  # Display locations as blue dots
            canvas.create_text(pixel_x, pixel_y + 10, text=location_id, font=("Helvetica", font_size), tags="elements")
            validPoles.append(location_id)
    
def updateFontSize():
    """
    Used to maintain text font size when zooming 

    """
    zoom_level = canvas.canvasx(1) - canvas.canvasx(0)
    font_size = max(8, 8 / zoom_level)
    

def find_closest_point(mouse_x, mouse_y, valid_keys):
    """
    Find the pole ID of the pole closest to the mouse when called

    Args:
    float: x coordinate of cursor on canvas
    float: y coordinate of cursor on canvas
    list: list of all possible poles that are not ignore

    Returns:
    string: poleID of the closest valid pole
    """
    
    closest_point = None
    closest_distance = float('inf')  # Initialize with positive infinity

    for key, (point_x, point_y) in poles.items():
        # Check if the key is in the valid keys list
        if key not in valid_keys:
            continue

        # Calculate the Euclidean distance between the mouse cursor and the point
        distance = math.sqrt((mouse_x - point_x) ** 2 + (mouse_y - point_y) ** 2)

        if distance < closest_distance:
            closest_distance = distance
            closest_point = key

    return closest_point

def find_related_keys(closest_point, second_dict):
    """
    Finds the adjacent poles to the given pole ID

    Args:
    string: PoleID of pole in question 
    dict: dictionary containing FPID and LPID information to find all connected poles

    Returns:
    list: related_keys is the span names connected to that pole
    list: related_poles is the pole IDs connected to that pole
    """
    related_keys = []
    related_poles = []

    for key, value in second_dict.items():
        if closest_point in value:
            related_keys.append(key)
            related_poles.append(value[0] if closest_point == value[1] else value[1])

    return related_keys, related_poles

def addSelectedSpan(connectedSpans, lastPoleID, nextPoleID):
    """
    This function adds spans to the list of selected spans very quickly by only allowing those directly connected to the pole

    Args:
    list: list of all the spans that connect to a given pole
    string: the pole selected right before the current one. Used to keep eliminate anoither option speeding up process
    string: the pole ID most recently selected

    Returns:
    string: spanID
    """

    for key in connectedSpans:
        if key in spans and isinstance(spans[key], tuple) and len(spans[key]) == 2:

            if lastPoleID == spans[key][0] and nextPoleID == spans[key][1]:
                selectedSpans.append(key)
                return key
            elif lastPoleID == spans[key][1] and nextPoleID == spans[key][0]:
                selectedSpans.append(key)
                return key
    return None

def recreate_lines(tag):
    """
    Finds the midspans with the passed tags and replaces them with new segmentation and loading for different colours 

    Args:
    string: indicator tag of which category of span to collect. Coloured spans have their span name as the tag, and the initial spans have the tag 'SpanID_Init'

    """
    global selectedSpans
    spansListed = []
    lines_to_recreate = []
    existingCables = []
    newCables = []
    exLoad = False
    newLoad = False
    

    for span in selectedSpans:
        midSection = True
        canvas.delete(span)
        spanRow = (SCD[SCD.iloc[:, 1] == span].index[0])
        if not pd.isna(SCD.loc[spanRow, 'Existing Loading']):
            existingCables = (str(SCD.loc[spanRow, 'Existing Loading'])).split(',')
            exLoad = True
        if not pd.isna(SCD.loc[spanRow, 'New Loading']):
            newCables = (str(SCD.loc[spanRow, 'New Loading'])).split(',')
            newLoad = True
        selectedCables = existingCables + newCables
        line_to_recreate = (canvas.find_withtag(span+tag))
        canvas.delete(span)

        coords = canvas.coords(line_to_recreate)
        
        # Create a new line with the updated coordinates and a color
    
        dx = (coords[2] - coords[0]) / len(selectedCables)
        dy = (coords[3] - coords[1]) / len(selectedCables)

        start_x = coords[0]
        start_y = coords[1]
        
        for index, cable in enumerate(selectedCables):

            segment_start_x = start_x + index * dx
            segment_start_y = start_y + index * dy
            segment_end_x = start_x + (index + 1) * dx
            segment_end_y = start_y + (index + 1) * dy            
            try:
                colour = colourDict[cable]
            except:
                colour = "red"

            scaled_dx, scaled_dy = scale_vector(dx, dy, 2)

            if cable in existingCables:
                new_line = canvas.create_line(segment_start_x, segment_start_y, segment_end_x, segment_end_y, fill=colour, width=4)
            else:
                if exLoad and newLoad and midSection:
                    new_midSection = canvas.create_line(segment_start_x - scaled_dx, segment_start_y - scaled_dy, segment_start_x + scaled_dx, segment_start_y + scaled_dy, fill="black", width=13)
                    midSection = False
                    canvas.tag_lower(new_midSection)
                    canvas.addtag_withtag(span, new_midSection)
                new_line = canvas.create_line(segment_start_x, segment_start_y, segment_end_x, segment_end_y, fill=colour, width=7)
                
            canvas.addtag_withtag(span, new_line)
            canvas.tag_raise(new_line)
            
            raisePoles()
        if exLoad and newLoad and not midSection:
            canvas.tag_raise(new_midSection)


def normalize_vector(dx, dy):
    """
    Normalizes the direction vectors to a unit value

    Args:
    float: x  vector
    float: y  vector
    

    Returns:
    float: x unit vector 
    float: y unit vector
    """
    magnitude = math.sqrt(dx**2 + dy**2)
    if magnitude == 0:
        return 0, 0  # To avoid division by zero
    normalized_dx = dx / magnitude
    normalized_dy = dy / magnitude
    return normalized_dx, normalized_dy

def scale_vector(dx, dy, new_magnitude):
    """
    Normalizes the direction vectors to a unit value

    Args:
    float: x vector
    float: y vector
    int: final desired magnitude of vector
    

    Returns:
    float: x scaled vector 
    float: y scaled vector
    """
    normalized_dx, normalized_dy = normalize_vector(dx, dy)
    scaled_dx = normalized_dx * new_magnitude
    scaled_dy = normalized_dy * new_magnitude
    return scaled_dx, scaled_dy
        
def recreate_Init_lines(tag):
    """
    Recreates the cable ladigns from saved excel file if present

    Args:
    string: tag that is just "_Init" to target the inital spans
    
    """
    # Find all lines with the "your_tag" tag
    global selectedSpans
    spansListed = []
    existingCables = []
    newCables = []
    lines_to_recreate = []

    for indexLine, row in SCD.iterrows():
        exLoad = False
        newLoad = False
        midSection = True

        
        if (not pd.isna(row['Existing Loading']) and row['Existing Loading'] != 0):
            existingCables = (str(row['Existing Loading'])).split(',')
            exLoad = True
        if (not pd.isna(row['New Loading']) and row['New Loading'] != 0):
            newCables = (str(row['New Loading'])).split(',')
            newLoad = True
        if (not pd.isna(row['Existing Loading']) and row['Existing Loading'] != 0) or (not pd.isna(row['New Loading']) and row['New Loading'] != 0):
        

            selectedCables = existingCables + newCables
            line_to_recreate = (canvas.find_withtag(row['SPN_N']+tag))
            coords = canvas.coords(line_to_recreate)
        
            # Create a new line with the updated coordinates and a random color
            dx = (coords[2] - coords[0]) / len(selectedCables)
            dy = (coords[3] - coords[1]) / len(selectedCables)

            start_x = coords[0]
            start_y = coords[1]
        
            for index, cable in enumerate(selectedCables):

                segment_start_x = start_x + index * dx
                segment_start_y = start_y + index * dy
                segment_end_x = start_x + (index + 1) * dx
                segment_end_y = start_y + (index + 1) * dy            

                try:
                    colour = colourDict[cable]
                except:
                    colour = "red"
                
                scaled_dx, scaled_dy = scale_vector(dx, dy, 2)

                if cable in existingCables:
                    new_line = canvas.create_line(segment_start_x, segment_start_y, segment_end_x, segment_end_y, fill=colour, width=4)
                else:
                    if exLoad and newLoad and midSection:
                        new_midSection = canvas.create_line(segment_start_x - scaled_dx, segment_start_y - scaled_dy, segment_start_x + scaled_dx, segment_start_y + scaled_dy, fill="black", width=13)
                        midSection = False
                        canvas.tag_lower(new_midSection)
                        canvas.addtag_withtag(row['SPN_N'], new_midSection)
                    new_line = canvas.create_line(segment_start_x, segment_start_y, segment_end_x, segment_end_y, fill=colour, width=7)
                
                canvas.addtag_withtag(row['SPN_N'], new_line)
            raisePoles()
            if exLoad and newLoad and not midSection:
                canvas.tag_raise(new_midSection)


            
        existingCables.clear()
        newCables.clear()
    lines_to_recreate.clear()
    spansListed.clear()            

def recreate_MPTs():
    """
    Adds MPTs and splices from existing data

    """

    
    global spliceMPTs

    
    for index, row in exactLoading.iterrows():
        if row['Splice/MPT'] == 'Splice':
            poleID = str(row['Pole_'])
            poleType[poleID] = 'Splice'
            xMPT1, yMPT1 = poles[poleID]
            splicePoles.append(poleID)
            
            canvas.create_oval(xMPT1, yMPT1 - 15, xMPT1 + 20, yMPT1 - 5, fill="red", tags=f"Splice_{poleID}")
            
            spliceID = poleID
            spliceMPTs[str(spliceID)] = []
            
    for index, row in exactLoading.iterrows(): 
        if row['Splice/MPT'] != '' and not pd.isna(row['Splice/MPT']) and not row['Splice/MPT'] == 'Splice':
            poleID = str(row['Pole_'])
            #print(f"added MPT at {row['Pole_']}")
            spliceID = str(row['Splice/MPT'])[4:]
            #print(f"spliceID: {spliceID}")
            spliceMPTs[str(spliceID)].append(poleID)
            #poleID = find_closest_point(canvas_x, canvas_y, totalPoles)
            xMPT1, yMPT1 = poles[poleID]
            poleType[poleID] = f'MPT_{spliceID}'
            mptPoles.append(poleID)
            drawMPT(mptPoles[-1], str(spliceID))
            drawHex(xMPT1, yMPT1, poleID)
            raiseMPTelements(splicePoles[-1])
            raiseMPTs()


def find_shortest_path(startPole, poleID):
    """
    Finds the shortest path between two poles along spans. Used for pathfinding tail when placing MPT.
    Essentially will spread along spans in all possible directions one span per iteration until the desired pole is located. 

    Args:
    string: pole with MPT
    string: pole with splice
    

    Returns:
    list: list of span IDs between the two poles
    """
    #print(f"finding new path between {startPole} & {poleID}")
    possibleSpans = []
    possiblePoles = []
    visited_lines = set()
    shortestPath = []
    spanPath = []
    polePath = []

    possibleSpans, possiblePoles = find_related_keys(poleID, spans)
    rows = len(possiblePoles)
    cols = 1

    # Initialize a 2D list with all elements set to a default value (e.g., 0)
    possiblePaths = [[None for _ in range(cols)] for _ in range(rows)]
    possiblePathSpans = [[None for _ in range(cols)] for _ in range(rows)]
    
    index = 0

    #initialize list of lists
    for poleIndex, pole in enumerate(possiblePoles):
        polelist = [pole]
        possiblePaths[(poleIndex)].append(pole)
        possiblePathSpans[(poleIndex)][0] = (possibleSpans[(poleIndex)])

    while True:
        #looping until condiiton is met
        #get adj poles, add adj spans to list
        #check adj poles for target, yes then break, no the continue
        
        #check if the latest entry is the right pole and retunr spans if it is
        for pathIndex, path in enumerate(possiblePaths):

            if not isinstance(path, list):
                path = [path]
            if not isinstance(possiblePathSpans[pathIndex], list):
                possiblePathSpans[pathIndex] = [possiblePathSpans[pathIndex]]
            if not isinstance(possiblePaths, list):
                possiblePaths[pathIndex] = [possiblePaths]
            if index > 500:
                return None
            if path[-1] == startPole:
                return possiblePathSpans[pathIndex]
            else:
                
                possibleSpans, possiblePoles = find_related_keys(path[-1], spans)
                possiblePoles, possibleSpans = nextPoles(possiblePoles, possibleSpans, path[-1], possiblePathSpans[pathIndex], path)
                if len(possiblePoles) == 0:
                       #no more paths
                       pass
                elif len(possiblePoles) == 1:
                    possiblePaths[pathIndex].append(possiblePoles[0])
                    possiblePathSpans[pathIndex].append(possibleSpans[0])
                else:
                    polePath = copy.deepcopy(path)
                    spanPath = copy.deepcopy(possiblePathSpans[pathIndex])
                    
                    for newPathIndex, newPole in enumerate(possiblePoles):
                        if newPathIndex == 0:
                            possiblePaths[pathIndex].append(possiblePoles[newPathIndex])
                            possiblePathSpans[pathIndex].append(possibleSpans[newPathIndex])
                            #replace OG path
                        else:
                            #add more paths:
                            polePath.append(newPole)

                            spanPath.append(possibleSpans[newPathIndex])
                            
                            possiblePaths.append(copy.deepcopy(polePath))
                            possiblePathSpans.append(copy.deepcopy(spanPath))

                            del polePath[-1]
                            del spanPath[-1]
           
            
                index += 1
        
        if index > 500:
            return None
        index +=1


def nextPoles(possiblePoles, possibleSpans, lastPoleID, recordedSpans, path):
    """
    Finds the possible connected poles not prviously inspected 

    Args:
    list: adjacent poles to a given pole
    list: connected spans to a given pole
    string: Pole ID for the last pole
    list: previous spans to that point
    list: previous poles to that point
    

    Returns:
    list: possible poles that could be the next ones
    list: the spans asscocietd with those poles
    """
    for index, span in enumerate(possibleSpans):
        target = (SCD[SCD.iloc[:, 1] == span].index[0])
        if SCD.loc[target, 'Verified'] == 'Ignore':
            possibleSpans.remove(span)
        elif span in recordedSpans:
            possibleSpans.remove(span)
    for index, pole in enumerate(possiblePoles):
        if pole in splicePoles:
            possiblePoles.remove(pole)
        elif pole == lastPoleID:
            possiblePoles.remove(pole)
            del possibleSpans[index]
        elif pole in path:
            possiblePoles.remove(pole)

    return possiblePoles, possibleSpans
    
def drawHex(center_x, center_y, poleID):
    """
    Finds the possible connected poles not prviously inspected 

    Args:
    int: x coordinate for the center of a hexagon
    int: y coordinate for the sente rof a hexagon
    string: pole ID asscoaited with the MPT
    
    """
    center_x += 10
    center_y -= 10
    radius = 8 # Change this to change the size of the hexagon
    
    hexagon_vertices = []
    for i in range(6):
        angle = 2 * math.pi / 6 * i
        x = center_x + radius * math.cos(angle)
        y = center_y + radius * math.sin(angle)
        hexagon_vertices.append((x, y))

    # Draw the hexagon on the canvas
    hexagon = canvas.create_polygon(hexagon_vertices, outline="blue", fill="white", width=2, tag=f"MPT_{poleID}")
    canvas.addtag_withtag('MPT', hexagon)

def raiseMPTs():
    """
    brings MPT hexagons forward so they are not obstructed by the lines
    """
    canvas.tag_raise('MPT')

def drawMPT(mptPoleID, splicePoleID):
    """
    Draws the lines from a splice to an MPT. If pole identified as having existing MPT data, path from exactLoading will be used, otherwise new path will be found

    Args:
    string: Pole ID of pole with MPT
    string: Pole ID for the splice it is connected to 
    
    """
    global mptPaths
    target = (PCD[PCD.iloc[:, 1] == mptPoleID].index[0])
    if pd.isna(PCD.at[target, 'Path']) or PCD.at[target, 'Path'] == '':
        path = find_shortest_path(mptPoleID, splicePoleID)
    else:
        path = (str(PCD.at[target, 'Path'])).split(',')
    mptPaths[mptPoleID] = path
    #print(mptPoleID)
    lines_to_recreate = []
    spansListed = []
    if not isinstance(path, list):
                path = [path]
    for span in path:
        lines_to_recreate.append(canvas.find_withtag(str(span).strip('{}')+"_Init"))
        spansListed.append(span)

    if len(lines_to_recreate) > 0:
        for indexLine, line in enumerate(lines_to_recreate):
            
            coords = canvas.coords(line)
            start_x = coords[0] + 10
            start_y = coords[1] - 10
            end_x = coords[2] + 10
            end_y = coords[3] - 10
                
            new_line = canvas.create_line(start_x, start_y, end_x, end_y, fill="blue", width=5)
            canvas.addtag_withtag(f"MPTline_{mptPoleID}", new_line)
        canvas.tag_lower(new_line)

def on_canvas_right_click(event):
    """
    Event triggered when right mouse button is clicked on canvas

    """
    
    canvas_x = canvas.canvasx(event.x)
    canvas_y = canvas.canvasy(event.y)
    amDrawing = True

    global firstClickTrue
    global connectedSpans
    global connectedPoles
    global lastPoleID
    global validPoles
    global scale
    global drawingMPTs
    global mptCount
    global poleType
    global spliceMPTs
    global mptPaths
    global spliceID
    global PCD

    #Check if /Place Plice/MPTs' is toggled
    if drawingMPTs:

        poleID = str(find_closest_point(canvas_x, canvas_y, validPoles))

        # Checks if splice/MPT already present on pole 
        if poleType[poleID] != '':
            target = (PCD[PCD.iloc[:, 1] == poleID].index[0])
            if firstClickTrue:
                if poleType[poleID] == 'Splice':
                    spliceID = poleID

                else:
                    spliceID = str(str(PCD.at[target, 'Splice/MPT'])[4:])

                
                firstClickTrue = False

            else:
                #clicking on a pole already with soemthing, need to delete MPT if mpt, delete all mpts and splice if splice
                if poleType[poleID] == 'Splice':
                    
                    canvas.delete(f"Splice_{poleID}")
                    poleType[(poleID)] = ''
                    firstClickTrue = True

                    for mptPole in spliceMPTs[(poleID)]:
                        canvas.delete(f"MPT_{mptPole}")
                        canvas.delete(f"MPTline_{mptPole}")
                        mptPaths[(mptPole)].clear()
                        poleType[(mptPole)] = ''

                    del spliceMPTs[(poleID)]

                elif not pd.isna(poleType[(poleID)]) and not poleType[(poleID)] == '':
                    canvas.delete(f"MPT_{poleID}")
                    canvas.delete(f"MPTline_{poleID}")
                    mptPaths[str(poleID)].clear()
                    poleType[(poleID)] = ''
                    poleType[(poleID)] = ''
                    PCD.at[target, 'Splice/MPT'] = ''

        else:
            if firstClickTrue:
                #placing a splice

                selectedSpans.clear()
                
                poleType[poleID] = 'Splice'
                xMPT1, yMPT1 = poles[poleID]
                splicePoles.append(poleID)
                
                canvas.create_oval(xMPT1, yMPT1 - 15, xMPT1 + 20, yMPT1 - 5, fill="red", tags=f"Splice_{poleID}")
                firstClickTrue = False
                spliceID = poleID
                spliceMPTs[str(spliceID)] = []
            elif not firstClickTrue:
                spliceMPTs[str(spliceID)].append(str(poleID))
                xMPT1, yMPT1 = poles[poleID]
                poleType[poleID] = f'MPT_{spliceID}'
                mptPoles.append(poleID)
                drawMPT(mptPoles[-1], splicePoles[-1])
                drawHex(xMPT1, yMPT1, poleID)
                raiseMPTelements(splicePoles[-1])
                raiseMPTs()
            
    # if 'Place Splice/MPTs' not toggled    
    else:
        if firstClickTrue:
            selectedSpans.clear()
            poleID = str(find_closest_point(canvas_x, canvas_y, validPoles))
            connectedSpans, connectedPoles = find_related_keys(poleID, spans)
            firstClickTrue = False
            lastPoleID = poleID
            
        elif not firstClickTrue:
            nextPoleID = find_closest_point(canvas_x, canvas_y, connectedPoles)
            connectedSpans, connectedPoles = find_related_keys(nextPoleID, spans)

            selectedSpan = addSelectedSpan(connectedSpans, lastPoleID, nextPoleID)


            x1, y1 = poles[lastPoleID]
            x2, y2 = poles[nextPoleID]
            
            canvas.create_line(x1*zoomFactor, y1*zoomFactor, x2*zoomFactor, y2*zoomFactor, fill="orange", width=4, tags=str(selectedSpan+"Temp"))
            lastPoleID = nextPoleID
            raisePoles()


def customize_button(button):
    """
    Changes appearance of button

    Args:
    object: button object
    """
    button.config(font=("Helvetica", 14))
    button.selectColor = "red"

# Function to handle right-click drag
def raisePoles():
    """
    brings poles forward to not be obscured by spans
    """
    canvas.tag_raise("elements")


def raiseMPTelements(spliceID):
    """
    Brings all elements asscietd with a certain splice

    Args:
    string: splice pole ID
    """
    
    canvas.tag_raise(f"MPT_{spliceID}")
    canvas.tag_raise(f"Splice_{spliceID}")


def number_press(index):
    """
    Increments whichever button was pressed indicating number of cables 
    
    Args:
    int: index out of 15 of which button was pressed
    
    """
    value=selected_option.get()
    numCables[index] += 1

    buttons[index].config(text="  "+str(numCables[index])+"  ")

def clear_selected():
    """
    Clears the selected spans, activate son button press

    """
    global selectedSpans
    global firstClickTrue
    firstClickTrue = True
    for span in selectedSpans:
        canvas.delete(span+"Temp")
    selectedSpans.clear()

def set_cables(columnName):
    """
    Saves selected cables for the selected spans to SCD dataframe under existing loading

    Args:
    string: either 'Exisitng Loading' or 'New Loading' depending on which button was pressed
    """
    global firstClickTrue
    global selectedSpans
    global selectedCables
    firstClickTrue = True

    for index, option in enumerate(options):
        for i in range(numCables[index]):
            selectedCables.append(dropdownsVar[index].get())
        numCables[index] = 0
        buttons[index].config(text="  0  ")
        
    for span in selectedSpans:
        spanRow = (SCD[SCD.iloc[:, 1] == span].index[0])
        
        cablesLoaded = ','.join(selectedCables)
    
        SCD.loc[spanRow, columnName] = cablesLoaded

    recreate_lines("_Init")
    
    for span in selectedSpans:
        canvas.delete(span+"Temp")


    selectedSpans.clear()
    selectedCables.clear()

def append_cables(columnName):
    """
    Takes data from SCD dataframe and appends new selection to new loading

    Args:
    string: either 'Exisitng Loading' or 'New Loading' depending on which button was pressed
    """
    selectedSpanCables = []
    newCables = []
    global selectedSpans
    selectedCables = []
    global firstClickTrue
    firstClickTrue = True

    tempSelectedSpans = copy.deepcopy(selectedSpans)
    for index, span in enumerate(tempSelectedSpans):
        selectedSpans.clear()
        selectedSpans.append(span)
        selectedCables.clear()
        
        spanRow = (SCD[SCD.iloc[:, 1] == span].index[0])
        if not pd.isna(SCD.loc[spanRow, columnName]):
            selectedCables = SCD.at[spanRow, columnName].split(',')
            
            for indexC, option in enumerate(options):

                for i in range(numCables[indexC]):
                    newCable = dropdownsVar[indexC].get()
                    selectedCables.append(newCable)

            SCD.at[spanRow, columnName] = ','.join(selectedCables)
            canvas.delete(span)
            recreate_lines("_Init")
            selectedCables.clear()
            canvas.delete(span+"Temp")
        else:
            for indexC, option in enumerate(options):

                for i in range(numCables[indexC]):
                    newCable = dropdownsVar[indexC].get()
                    selectedCables.append(newCable)

            SCD.at[spanRow, columnName] = ','.join(selectedCables)
            canvas.delete(span)
            recreate_lines("_Init")
            
            selectedCables.clear()
            canvas.delete(span+"Temp")
    for indexC, option in enumerate(options):
        numCables[indexC] = 0
        buttons[indexC].config(text="  0  ")

        
def delete_bundle():
    """
    Removes both exisitng and new loading from selected spans 

    """
    global selectedSpans
    global firstClickTrue
    firstClickTrue = True

    for span in selectedSpans:
        canvas.delete(span+"Temp")
        canvas.delete(span)
        spanRow = (SCD[SCD.iloc[:, 1] == span].index[0])
        SCD.at[spanRow, 'Existing Loading'] = None
        SCD.at[spanRow, 'New Loading'] = None
        
    selectedCables.clear()

def clear_bundle():
    """
    Returns all cable buttons to 0
    """
    for index, option in enumerate(options):
        buttons[index].config(text="  0  ")
        numCables[index] = 0

def set_selected_option(index):
    """
    Records the new cable in dropdown cell
    """
    selected_option = options[index]
    dropdownsVar[index].set(selected_option)

def on_dropdown_change(event):
    """
    Triggers when drop down is changed

    """
    global options
    for index, option in enumerate(options):
        
        newCableDropdown = dropdownsVar[index].get()
        options[index] = newCableDropdown
        buttons[index].config(bg=colourDict[newCableDropdown])

def place_splice():
    """
    Toggle for entering splice/MPT mode
    
    """
    global drawingMPTs
    global firstClickTrue
    global mptPaths
    global mptCount
    global PoleType
    
    if placeSpliceState.get():
        placeSplice.config(relief=tk.RAISED)  # Outdent the button when released
        placeSplice.config(text="Place Splice & MPTs")
        firstClickTrue = True
        drawingMPTs = False
        placeSpliceState.set(False)

        for index, row in SCD.iterrows():
            mptCount[row['SPN_N']] = 0
            
        for key, values in mptPaths.items():
            for value in values:
                mptCount[value] += 1
            poleRow = (PCD[PCD.iloc[:, 1] == key].index[0])    
            PCD.loc[poleRow, 'Path'] = ','.join(mptPaths[key])

        for key, value in mptCount.items():
            spanRow = (SCD[SCD.iloc[:, 1] == key].index[0])
            SCD.loc[spanRow, '# MPTs'] = value
        
        for key, value in poleType.items():
            poleRow = (PCD[PCD.iloc[:, 1] == key].index[0])
            PCD.loc[poleRow, 'Splice/MPT'] = value

    else:
        placeSplice.config(relief=tk.SUNKEN)  # Indent the button when pressed
        placeSplice.config(text="Placing splice / MPTs")
        drawingMPTs = True
        placeSpliceState.set(True)
    


# Create the main window
root = tk.Tk()
root.title(f"ExactLoadingToolR{version}")

# Create a canvas for displaying the locations
canvas = tk.Canvas(root, width=1200, height=800)

# Create a sidebar
sidebar = tk.Frame(root, width=300, height=800, bg="lightgray")

# Use pack geometry manager for both canvas and sidebar
canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
sidebar.pack(side=tk.RIGHT, fill=tk.BOTH, anchor="e")

checklist = tk.Frame(sidebar, width=300, bg='lightgray', height=800)
checklist.pack(pady=1, padx=10, expand=True)

rows = 20
columns = 3

# Create rows and columns with specific sizes
for i in range(rows):
    checklist.grid_rowconfigure(i, minsize=30)  # Set the height of each row

for j in range(columns):
    checklist.grid_columnconfigure(j, minsize=80)

# Add poles and spans
draw_elements()

if exactLoadingExists:
    #draw spans for exisitng loading
    recreate_Init_lines("_Init")
    recreate_MPTs()

buttons = []
dropdowns = []
dropdownsVar = []
numCables = []
i = 0

# Adding all cable type and amount buttons
for index, option in enumerate(options):
    
    numCables.append(0)
    button = tk.Button(checklist, text="  "+str(numCables[index])+"  ", font=("Helvetica", 10), bg=colourDict[option], command=lambda r = index: number_press(r))
    customize_button(button)
    button.grid(row=i, column=0, sticky="ew", padx=2, pady=1)
    buttons.append(button)
    selected_option = tk.StringVar()
    selected_option.set(options[index]) 

    dropdown = ttk.Combobox(checklist, textvariable=selected_option, values=allCableOptions, state="readonly")
    dropdown['font'] = ('Helvetica', 12)
    dropdown['background'] = 'lightgray'
    dropdown.bind("<<ComboboxSelected>>", on_dropdown_change)

    dropdown.grid(row=index, column=1, columnspan=2, sticky="ew", pady=1, padx=2)
    dropdowns.append(dropdown)
    dropdownsVar.append(selected_option)
    set_selected_option(index)
    
    i = i+1
    
# Adding all the other buttons
placeSpliceState = tk.BooleanVar()
placeSpliceState.set(False)
placeSplice = tk.Button(checklist, text="Place Splice & MPTs", font=("Helvetica", 12), command=place_splice)
placeSplice.grid(row=i, column=0, columnspan=3, sticky="ew", padx=2, pady=1)
clearBundle = tk.Button(checklist, text="Clear Cables", font=("Helvetica", 12), command=clear_bundle)
clearBundle.grid(row=i+1, column=0, columnspan=2, sticky="ew", padx=2, pady=1)
deleteBundle = tk.Button(checklist, text="Delete Cables", font=("Helvetica", 12), command=delete_bundle)
deleteBundle.grid(row=i+1, column=2, columnspan=1, sticky="ew", padx=2, pady=1)
clearSelected = tk.Button(checklist, text="Clear Selection of Spans", font=("Helvetica", 12), command=clear_selected)
clearSelected.grid(row=i+2, column=0, columnspan=3, sticky="ew", padx=2, pady=1)
enterExistingData = tk.Button(checklist, text="Set Existing Loading", font=("Helvetica", 12), command=lambda: set_cables('Existing Loading'))
enterExistingData.grid(row=i+3, column=0, columnspan=2, sticky="ew", padx=2, pady=1)
appendEx = tk.Button(checklist, text="Append", font=("Helvetica", 12), command=lambda: append_cables('Existing Loading'))
appendEx.grid(row=i+3, column=2, columnspan=1, sticky="ew", padx=2, pady=1)
enterNewData = tk.Button(checklist, text="Set New Loading", font=("Helvetica", 12), command=lambda: set_cables('New Loading'))
enterNewData.grid(row=i+4, column=0, columnspan=2, sticky="ew", padx=2, pady=1)
appendNew = tk.Button(checklist, text="Append", font=("Helvetica", 12), command=lambda: append_cables('New Loading'))
appendNew.grid(row=i+4, column=2, columnspan=1, sticky="ew", padx=2, pady=1)




# Bind mouse events to canvas
canvas.bind("<ButtonPress-1>", on_canvas_click)
canvas.bind("<B1-Motion>", on_canvas_release)
#canvas.bind("<MouseWheel>", on_canvas_scroll)  # Uncomment this to enable scroll zooming, warning drawing does not work with this enabled
canvas.bind("<Button-3>", on_canvas_right_click)
root.protocol("WM_DELETE_WINDOW", on_closing)


# Start the main loop
root.mainloop()
