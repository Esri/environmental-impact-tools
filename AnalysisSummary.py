"""
-------------------------------------------------------------------------------
 | Copyright 2016 Esri
 |
 | Licensed under the Apache License, Version 2.0 (the "License");
 | you may not use this file except in compliance with the License.
 | You may obtain a copy of the License at
 |
 |    http://www.apache.org/licenses/LICENSE-2.0
 |
 | Unless required by applicable law or agreed to in writing, software
 | distributed under the License is distributed on an "AS IS" BASIS,
 | WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 | See the License for the specific language governing permissions and
 | limitations under the License.
 ------------------------------------------------------------------------------
 """

import arcpy
import sys
import os
import locale

# Anticipated inputs ...
# 1. Area of Interest: Project location. Feature Class. Point, line or polygon.
# 2. Buffer shape (optional): pre-analyzed buffer for project area.
# 3. Reporting Units: based on the analysis layer shape type
# 4. Output table: the GDB path and name of the final results table
# Interim results will be written to the Scratch workspace

args = sys.argv

input_aoi = args[1]
input_buffer_layer = args[2]
reporting_units = args[3]
output_table = args[4]

output_workspace = arcpy.env.scratchWorkspace
arcpy.env.workspace = output_workspace


# --------------------------------------
# Function to create the output table
def create_output(out_table):
    try:

        out_path = os.path.dirname(os.path.abspath(out_table))
        out_name = out_table.split("\\")[-1]

        arcpy.CreateTable_management(out_path, out_name, None, None)

        arcpy.AddField_management(out_table, "ANALYSISPROP", "TEXT", "", "", 100, 'Analysis Details')
        arcpy.AddField_management(out_table, "ANALYSISDESC", "TEXT", "", "", 255, '  ')

        return out_table

    except Exception as error:
        arcpy.AddError("Create Output Table Error: {}".format(error))
        return "empty"


# --------------------------------------
# Function to calculate areas
def get_area(input_fc, units):
    try:
        geometries = arcpy.CopyFeatures_management(input_fc, arcpy.Geometry())
        area = 0

        for geometry in geometries:
            area += geometry.getArea('GEODESIC', units)

        return area

    except Exception as error:
        arcpy.AddError("Get Area Error: {}".format(error))


# --------------------------------------
# Function to count points
def get_point_info(input_fc):
    try:
        result = arcpy.GetCount_management(input_fc)
        point_count = int(result.getOutput(0))

        if point_count == 0:
            message = "No points in area of interest"
        elif point_count == 1:
            # for row in arcpy.da.SearchCursor(input_fc, ["SHAPE@XY"]):
            #     x, y = row[0]
            # message = "Point at {}, {}".format(x, y)
            message = "1 point"
        else:
            message = "{} points".format(point_count)

        return message
    except Exception as error:
        arcpy.AddError('Get Point Info Error: {}'.format(error))


# --------------------------------------
# Function to calculate lengths
def get_line_info(input_fc, units):
    try:
        geometries = arcpy.CopyFeatures_management(input_fc, arcpy.Geometry())
        length = 0

        for geometry in geometries:
            length += geometry.getLength('GEODESIC', units)

        return length

    except Exception as error:
        arcpy.AddError('Get Line Info Error: {}'.format(error))


# --------------------------------------
# Function to abbreviate units of measure for readability
def abbreviate_units(units):
    try:
        lower_units = units.lower()
        return {'acres': 'acres',
                'hectares': 'hectares',
                'squaremiles': 'sq mi',
                'squarekilometers': 'sq km',
                'squaremeters': 'sq m',
                'squarefeet': 'sq ft',
                'miles': 'mi',
                'kilometers': 'km',
                'meters': 'm',
                'feet': 'ft'
                }.get(lower_units, "none")

    except Exception as error:
        arcpy.AddError('Abbreviate Units Error: {}'.format(error))


# --------------------------------------
# Function to abbreviate units of measure for readability
def get_area_units(units):
    try:
        lower_units = units.lower()
        return {'miles': 'squaremiles',
                'kilometers': 'squarekilometers',
                'meters': 'squaremeters',
                'feet': 'squarefeet'
                }.get(lower_units, "none")

    except Exception as error:
        arcpy.AddError('Abbreviate Units Error: {}'.format(error))


# --------------------------------------
# Main

try:
    locale.setlocale(locale.LC_ALL, '')

    property_table = create_output(output_table)

    # Describe Area of Interest properties
    aoi_properties = arcpy.Describe(input_aoi)
    aoi_shape = aoi_properties.shapeType
    aoi_description = ""  # in the case the aoi is not a polygon
    buffer_description = ""
    aoi_length = 0
    area_units = reporting_units

    if aoi_shape == "Polygon":
        aoi_area = get_area(input_aoi, reporting_units)
    elif aoi_shape == "Polyline":
        aoi_area = 0
        aoi_length = get_line_info(input_aoi, reporting_units)
        # in the case of a polyline aoi, need to get complementary area units
        area_units = get_area_units(reporting_units)
        arcpy.AddMessage("Calculating buffer area as {}.".format(area_units))
    else:
        aoi_area = 0
        aoi_description = get_point_info(input_aoi)

    total_area = aoi_area

    # If a buffer has been used, get the area of the buffer
    if input_buffer_layer != "#":
        buffer_area = get_area(input_buffer_layer, area_units)
    else:
        buffer_area = 0
        buffer_description = "No buffer defined."
    total_area = aoi_area + buffer_area

    # format the numbers to be nicely readable
    if total_area != 0:
        total_area = round(total_area, 2)
        total_area = locale.format('%.2f', total_area, grouping=True)
    if aoi_area != 0:
        aoi_area = round(aoi_area, 2)
        aoi_area = locale.format('%.2f', aoi_area, grouping=True)
    if buffer_area != 0:
        buffer_area = round(buffer_area, 2)
        buffer_area = locale.format('%.2f', buffer_area, grouping=True)
    if aoi_length != 0:
        aoi_length = round(aoi_length, 2)
        aoi_length = locale.format('%.2f', aoi_length, grouping=True)

    # populate the output table with information
    with arcpy.da.InsertCursor(property_table, ["ANALYSISPROP", "ANALYSISDESC"]) as prop_cursor:
        prop_cursor.insertRow(["Analysis Shape Type", aoi_shape])

        # Next, insert the AOI descriptive information
        if aoi_shape == "Polygon":
            description = str(aoi_area) + " {}".format(abbreviate_units(reporting_units))
            prop_cursor.insertRow(["Analysis Area", description])
        elif aoi_shape == "Polyline":
            description = str(aoi_length) + " {}".format(abbreviate_units(reporting_units))
            prop_cursor.insertRow(["Analysis Length", description])
        else:
            prop_cursor.insertRow(["Analysis Length", aoi_description])

        # Next, insert the buffer descriptive information if one was entered
        if buffer_area != 0:
            description = str(buffer_area) + " {}".format(abbreviate_units(area_units))
            prop_cursor.insertRow(["Buffer Area", description])
        else:
            prop_cursor.insertRow(["Buffer Area", buffer_description])

        # Lastly, if a buffer was entered and the aoi was a polygon layer, provide a total area
        if aoi_area != 0 and buffer_area != 0:
            description = str(total_area) + " {}".format(abbreviate_units(reporting_units))
            prop_cursor.insertRow(["Total Area", description])
        else:
            prop_cursor.insertRow([" ", " "])

except Exception as main_error:
    arcpy.AddError("Analysis Summary Error: {}".format(main_error))
    if arcpy.Exists(output_table):
        arcpy.Delete_management(output_table)
