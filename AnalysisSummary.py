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

        arcpy.AddField_management(out_table, "PROPERTY", "TEXT", "", "", 100, 'Analysis Details')
        arcpy.AddField_management(out_table, "DESCRIPTION", "TEXT", "", "", 255, '  ')

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
# Function to calculate areas
def get_point_info(input_fc):
    try:
        result = arcpy.GetCount_management(input_fc)
        point_count = int(result.getOutput(0))

        return "{} points".format(point_count)

    except Exception as error:
        arcpy.AddError('Get Point Info Error: {}'.format(error))


# --------------------------------------
# Function to calculate areas
def get_line_info(input_fc):
    try:
        result = arcpy.GetCount_management(input_fc)
        line_count = int(result.getOutput(0))

        return "{} lines".format(line_count)

    except Exception as error:
        arcpy.AddError('Get Line Info Error: {}'.format(error))


# --------------------------------------
# Main

try:
    property_table = create_output(output_table)

    # Describe Area of Interest properties
    aoi_properties = arcpy.Describe(input_aoi)
    aoi_shape = aoi_properties.shapeType
    aoi_description = ""  # in the case the aoi is not a polygon
    buffer_description = ""

    if aoi_shape == "Polygon":
        aoi_area = get_area(input_aoi, reporting_units)
    else:
        aoi_area = 0
        if aoi_shape == "Point":  # get some point info
            aoi_description = get_point_info(input_aoi)
        elif aoi_shape == "Polyline":  # get some line info
            aoi_description = get_line_info(input_aoi)

    total_area = aoi_area

    # If a buffer has been used, get the area of the buffer
    if input_buffer_layer != "#":
        buffer_area = get_area(input_buffer_layer, reporting_units)
    else:
        buffer_area = 0
        buffer_description = "No buffer defined."
    total_area = aoi_area + buffer_area

    # format the numbers to be nicely readable
    if total_area != 0:
        total_area = round(total_area, 2)
    if aoi_area != 0:
        aoi_area = round(aoi_area, 2)
    if buffer_area !=0:
        buffer_area = round(buffer_area, 2)

    rows = arcpy.InsertCursor(property_table)
    row = rows.newRow()
    row.setValue("PROPERTY", "Analysis Shape Type")
    row.setValue("DESCRIPTION", aoi_shape)
    rows.insertRow(row)

    if aoi_area != 0:
        row = rows.newRow()
        row.setValue("PROPERTY", "Analysis Area")
        row.setValue("DESCRIPTION", str(aoi_area) + " {}".format(reporting_units.lower()))
        rows.insertRow(row)
    else:
        row = rows.newRow()
        row.setValue("PROPERTY", "Analysis Area")
        row.setValue("DESCRIPTION", aoi_description)
        rows.insertRow(row)
    if buffer_area != 0:
        row = rows.newRow()
        row.setValue("PROPERTY", "Buffer Area")
        row.setValue("DESCRIPTION", str(buffer_area) + " {}".format(reporting_units.lower()))
        rows.insertRow(row)
    else:
        row = rows.newRow()
        row.setValue("PROPERTY", "Buffer Area")
        row.setValue("DESCRIPTION", buffer_description)
        rows.insertRow(row)
    row = rows.newRow()
    row.setValue("PROPERTY", "Total Area")
    row.setValue("DESCRIPTION", str(total_area) + " {}".format(reporting_units.lower()))
    rows.insertRow(row)

    del row
    del rows

except Exception as main_error:
    arcpy.AddError("Analysis Summary Error: {}".format(main_error))
    if arcpy.Exists(output_table):
        arcpy.Delete_management(output_table)


