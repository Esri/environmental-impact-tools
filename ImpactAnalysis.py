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
# 1. Analysis Type: Choose from: Basic Proximity, Feature Comparison, Distance
# 2. Analysis Layer: Feature class to be analyzed
# 3. Output fields: from the selected input Analysis layer, which fields should be included in the output result
# 4. Group like records:  merge records in the results so that only
#                         unique combinations of attributes appear in the output
# 5. Area of Interest: Project location. Feature Class. Point, line or polygon.
# 6. Buffer shape (optional): pre-analyzed buffer for project area.
# 7. Reporting Units: based on the analysis layer shape type
# 8. Output table: the GDB path and name of the final results table
# Interim results will be written to the Scratch workspace

args = sys.argv

analysis_type = args[1]
input_analysis_layer = args[2]
output_fields = args[3]
#group_output_records = args[4]
group_output_records = True  # This should just always happen for the released script
input_aoi = args[5]
input_buffer_layer = args[6]
reporting_units = args[7]
output_table = args[8]

# Termporary feature classes created during script execution
interim_output_aoi = "interim_result_aoi"
interim_output_buffer = "interim_result_buffer"
interim_output_merged = "interim_result"
interim_aoi_lines = "interim_aoi_lines"

# arcpy.env.overwriteOutput = True  # should be set by the user in Geoprocessing Options
output_workspace = arcpy.env.scratchWorkspace
arcpy.env.workspace = output_workspace

arcpy.AddMessage("Analysis Type: {}".format(analysis_type))
arcpy.AddMessage("Input analysis layer: {}".format(input_analysis_layer))
arcpy.AddMessage("Output fields: {}".format(output_fields))
arcpy.AddMessage("Group like records: {}".format(group_output_records))
arcpy.AddMessage("Input AOI: {}".format(input_aoi))
arcpy.AddMessage("Input buffer: {}".format(input_buffer_layer))
arcpy.AddMessage("Reporting units: {}".format(reporting_units))
arcpy.AddMessage("Output table: {}".format(output_table))

# Get units of measure squared away
area_units = reporting_units
if reporting_units in ["Meters", "Kilometers"]:  # then we need a different unit for area calculations
    area_units = "SQUAREKILOMETERS"
elif reporting_units in ["Feet", "Miles"]:
    area_units = "ACRES"


# --------------------------------------
# Function to validate script inputs
def validate_inputs():

    result = True
    reason = ""

    if analysis_type in ["Feature Comparison", "Basic Proximity", "Distance"]:
        result = True
    else:
        result = False
        reason = 'Analysis type must be one of: "Feature Comparison", "Basic Proximity", "Distance"'

    if arcpy.Exists(input_analysis_layer):
        fields = arcpy.ListFields(input_analysis_layer)
        field_names = []
        for field in fields:
            field_names.append(field.name.upper())

        check_fields = output_fields.split(';')
        for item in check_fields:
            if item.upper() in field_names:
                continue
            else:
                result = False
                reason += "\n No field {} found in {}".format(item, input_analysis_layer)
    else:
        result = False
        reason += '\n {} does not exist'.format(input_analysis_layer)

    if arcpy.Exists(input_aoi):
        a = 1
    else:
        result = False
        reason += '\n {} does not exist'.format(input_aoi)

    if input_buffer_layer != "#":
        if arcpy.Exists(input_buffer_layer):
            a = 1
        else:
            result = False
            reason += '\n {} does not exist'.format(input_buffer_layer)

    return [result, reason]


# --------------------------------------
# Basic Proximity Analysis Type
def basic_proximity(analysis_layer, select_by_layer, out_layer_name, layer_type):
    try:
        out_layer = output_workspace + "\\" + out_layer_name

        arcpy.MakeFeatureLayer_management(analysis_layer, out_layer_name)
        arcpy.SelectLayerByLocation_management(out_layer_name, 'intersect', select_by_layer)

        match_count = int(arcpy.GetCount_management(out_layer_name)[0])

        if match_count == 0:
            arcpy.AddMessage(("No features found in {0}".format(analysis_layer)))
            return "empty"

        else:
            arcpy.AddMessage(("{0} features found in {1}".format(match_count, analysis_layer)))

            arcpy.CopyFeatures_management(out_layer_name, out_layer)
            arcpy.AddField_management(out_layer, "ANALYSISTYPE", "TEXT", "", "", 10, "Analysis Result Type")
            exp = "'{}'".format(layer_type)

            arcpy.CalculateField_management(out_layer, "ANALYSISTYPE", exp, "PYTHON_9.3", None)

        return out_layer

    except Exception as error:
        arcpy.AddError("Basic Proximity Error: {}".format(error))


# --------------------------------------
# Feature Comparison Analysis Type
def feature_comparison(analysis_layer, clip_layer, out_layer_name, layer_type, clip_area):
    try:
        xy_tolerance = ""
        out_layer = output_workspace + "\\" + out_layer_name

        arcpy.Clip_analysis(analysis_layer, clip_layer, out_layer, xy_tolerance)

        match_count = int(arcpy.GetCount_management(out_layer_name)[0])

        if match_count == 0:
            arcpy.AddMessage(("No {0} features found in {1}.".format(analysis_layer, clip_layer)))
            return "empty"

        else:
            arcpy.AddMessage(("Processing {0} features in {1} within {2}.".format(match_count, analysis_layer, layer_type)))

            arcpy.AddField_management(out_layer, "ANALYSISTYPE", "TEXT", "", "", 10, "Analysis Result Type")
            exp = "'{}'".format(layer_type)
            arcpy.CalculateField_management(out_layer, "ANALYSISTYPE", exp, "PYTHON_9.3", None)

            desc = arcpy.Describe(out_layer)
            shape_type = desc.shapeType

            if shape_type == "Polygon":
                arcpy.AddField_management(out_layer, "ANALYSISAREA", "DOUBLE", "", "", "",
                                          "Total Area ({})".format(reporting_units))
                exp = '!shape.area@{}!'.format(reporting_units)
                arcpy.CalculateField_management(out_layer, "ANALYSISAREA", exp, "PYTHON_9.3", None)

                arcpy.AddField_management(out_layer, "ANALYSISPERCENT", "DOUBLE", "", "", "", "Percent of Area")
                exp = '(!ANALYSISAREA! / ' + str(clip_area) + ')*100'
                arcpy.CalculateField_management(out_layer, "ANALYSISPERCENT", exp, "PYTHON_9.3", None)

            elif shape_type == "Polyline":
                arcpy.AddField_management(out_layer, "ANALYSISLEN", "DOUBLE", "", "", "",
                                          "Total Length ({})".format(reporting_units))
                exp = '!shape.length@{}!'.format(reporting_units)
                arcpy.CalculateField_management(out_layer, "ANALYSISLEN", exp, "PYTHON_9.3", None)

            elif shape_type == "Point":
                arcpy.AddField_management(out_layer, "ANALYSISCOUNT", "SHORT", "", "", "", "Count of Features")
                exp = '1'
                arcpy.CalculateField_management(out_layer, "ANALYSISCOUNT", exp, "PYTHON_9.3", None)

            else:
                arcpy.AddMessage("Shape type not supported: {}".format(shape_type))

            return out_layer
    except Exception as error:
        arcpy.AddError("Feature Comparison Error: {}".format(error))


# --------------------------------------
# Distance Analysis Type -- for the Area of Interest layer
def distance_analysis_aoi(near_layer, aoi_layer, out_layer_name):
    try:
                    # Find everything intersecting the AOI
            arcpy.MakeFeatureLayer_management(near_layer, "near_layer")
            arcpy.SelectLayerByLocation_management("near_layer", 'intersect', aoi_layer)

            match_count = int(arcpy.GetCount_management("near_layer")[0])

            if match_count == 0:
                arcpy.AddMessage(("No features found in {0}".format(aoi_layer)))
                return "empty"

            else:
                arcpy.AddMessage(("{0} features found in {1}. Calculating distances.".format(match_count, aoi_layer)))
                arcpy.CopyFeatures_management("near_layer", out_layer_name)

                desc = arcpy.Describe(aoi_layer)
                input_shape_type = desc.shapeType

                if input_shape_type == "Polygon":
                    arcpy.PolygonToLine_management(aoi_layer, interim_aoi_lines)
                    arcpy.Near_analysis(out_layer_name, interim_aoi_lines, None, "NO_LOCATION", "ANGLE", "GEODESIC")
                else:
                    arcpy.Near_analysis(out_layer_name, aoi_layer, None, "NO_LOCATION", "ANGLE", "GEODESIC")

                arcpy.AddField_management(out_layer_name, "ANALYSISLOC", "TEXT", "", "", 100, "Location")

                exp = "getValue(!NEAR_DIST!)"
                codeblock = """def getValue(dist):
                        dist = round(dist, 4)
                        if dist == 0:
                            return 'Intersecting with AOI boundary'
                        else:
                            return 'Within AOI'"""
                arcpy.CalculateField_management(out_layer_name, "ANALYSISLOC", exp, "PYTHON_9.3", codeblock)

                exp = "getValue(!NEAR_DIST!, !NEAR_ANGLE!)"
                codeblock = """def getValue(dist, angle):
                        dist = round(dist, 4)
                        if dist == 0:
                            return 0
                        else:
                            return angle"""
                arcpy.CalculateField_management(out_layer_name, "NEAR_ANGLE", exp, "PYTHON_9.3", codeblock)

                return out_layer_name

    except Exception as error:
        arcpy.AddError("Distance Analysis (AOI) Error: {}".format(error))


# --------------------------------------
# Distance Analysis Type -- for the Buffer layer
def distance_analysis_buffer(near_layer, aoi_layer, buffer_layer, out_layer_name):
    try:
        # Find everything intersecting the buffer
        arcpy.MakeFeatureLayer_management(near_layer, "near_layer")
        arcpy.SelectLayerByLocation_management("near_layer", 'intersect', buffer_layer)

        # Remove from the selection everything intersecting the AOI (these have already been accounted for)
        arcpy.SelectLayerByLocation_management("near_layer", 'intersect', aoi_layer, None, "REMOVE_FROM_SELECTION")
        match_count = int(arcpy.GetCount_management("near_layer")[0])

        if match_count == 0:
            arcpy.AddMessage(("No features found in {0}".format(buffer_layer)))
            return "empty"

        else:
            arcpy.AddMessage(("{0} features found in {1}. Calculating distances.".format(match_count, buffer_layer)))
            arcpy.CopyFeatures_management("near_layer", out_layer_name)

            desc = arcpy.Describe(aoi_layer)
            input_shape_type = desc.shapeType

            desc_near_layer = arcpy.Describe(out_layer_name).spatialReference
            global reporting_units
            reporting_units = desc_near_layer.linearUnitName

            if input_shape_type == "Polygon":
                arcpy.PolygonToLine_management(aoi_layer, interim_aoi_lines)
                arcpy.Near_analysis(out_layer_name, interim_aoi_lines, None, "NO_LOCATION", "ANGLE", "GEODESIC")
            else:
                arcpy.Near_analysis(out_layer_name, aoi_layer, None, "NO_LOCATION", "ANGLE", "GEODESIC")

            arcpy.AddField_management(out_layer_name, "ANALYSISLOC", "TEXT", "", "", 100, "Location")

            exp = "getValue(!NEAR_DIST!)"
            codeblock = """def getValue(dist):
                if dist > 0:
                    return 'Outside AOI, within buffer'
                else:
                    return 'Unexpected result'"""

            arcpy.CalculateField_management(out_layer_name, "ANALYSISLOC", exp, "PYTHON_9.3", codeblock)

            return out_layer_name

    except Exception as error:
        arcpy.AddError("Distance Analysis (Buffer) Error: {}".format(error))


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
# Function to format outputs
def format_outputs(output_layer, out_fields):

    try:

        field_mapper = arcpy.FieldMappings()

        fields = arcpy.ListFields(output_layer)
        compare_fields = []
        stat_fields = ""
        for field in fields:

            if field.name in out_fields:
                # Keep Me
                fm = arcpy.FieldMap()
                fm.addInputField(output_layer, field.name)
                field_mapper.addFieldMap(fm)
                compare_fields.append(field.name)

            elif field.name in ['ANALYSISPERCENT', 'ANALYSISAREA', 'ANALYSISLEN', 'ANALYSISCOUNT']:
                # Keep Me -- these are statistics fields
                fm = arcpy.FieldMap()
                fm.addInputField(output_layer, field.name)
                field_mapper.addFieldMap(fm)
                if stat_fields == "":
                    stat_fields = field.name + " SUM"
                else:
                    stat_fields = stat_fields + ";" + field.name + " SUM"

            elif field.name in ['ANALYSISTYPE', 'ANALYSISLOC', 'NEAR_DIST', 'NEAR_ANGLE']:
                # Keep Me -- these are informational/descriptive fields
                fm = arcpy.FieldMap()
                fm.addInputField(output_layer, field.name)
                field_mapper.addFieldMap(fm)
                compare_fields.append(field.name)

            elif field.type.upper() in ["OID"]:
                # Keep Me
                fm = arcpy.FieldMap()
                fm.addInputField(output_layer, field.name)
                field_mapper.addFieldMap(fm)

            else:
                # Delete Me
                # arcpy.AddMessage("{}: Delete Me".format(field.name))
                continue

        out_path = os.path.dirname(os.path.abspath(output_table))
        out_name = output_table.split("\\")[-1]

        # Clean up the results by deleting or merging identical records
        if analysis_type == "Basic Proximity":
            # Nothing special, just remove duplicates
            arcpy.TableToTable_conversion(output_layer, out_path, out_name, field_mapping=field_mapper)
            if group_output_records:
                arcpy.DeleteIdentical_management(output_table, compare_fields)
        elif analysis_type == "Distance":
            # Don't remove duplicates, if there are 5 eagle nests, i want to know there
            #                          are 5 and their unique distances from the project
            arcpy.TableToTable_conversion(output_layer, out_path, out_name, field_mapping=field_mapper)

            # Update field aliases to be readable and have distance units embedded
            arcpy.AlterField_management(output_table, "NEAR_DIST", None,
                                        "Distance ({})".format(reporting_units))
            arcpy.AlterField_management(output_table, "NEAR_ANGLE", None, "Direction")

        else:
            # Feature Comparison -- summarize the output layer based on the 'keep fields'
            # Possible Stat Fields .. 'ANALYSISPERCENT', 'ANALYSISAREA', 'ANALYSISLEN', 'ANALYSISCOUNT'
            if group_output_records:
                arcpy.Statistics_analysis(output_layer, output_table, stat_fields, compare_fields)
                arcpy.DeleteField_management(output_table, ["FREQUENCY"])
                if "ANALYSISPERCENT" in stat_fields:
                    arcpy.AlterField_management(output_table, "SUM_ANALYSISPERCENT", "ANALYSISPERCENT", "Percent of Area")
                    arcpy.AlterField_management(output_table, "SUM_ANALYSISAREA", "ANALYSISAREA",
                                                "Total Area ({})".format(reporting_units))
                elif "ANALYSISLEN" in stat_fields:
                    arcpy.AlterField_management(output_table, "SUM_ANALYSISLEN", "ANALYSISLEN",
                                                "Total Length ({})".format(reporting_units))
                else:
                    arcpy.AlterField_management(output_table, "SUM_ANALYSISCOUNT", "ANALYSISCOUNT", "Count of Features")
            else:
                arcpy.TableToTable_conversion(output_layer, out_path, out_name, field_mapping=field_mapper)

        return True
    except Exception as error:
        arcpy.AddError("Format Outputs Error: {}".format(error))


# --------------------------------------
# Create the output table with information about there being no results
def create_empty_output(out_table):
    try:

        out_path = os.path.dirname(os.path.abspath(out_table))
        out_name = out_table.split("\\")[-1]

        arcpy.CreateTable_management(out_path, out_name, None, None)
        arcpy.AddField_management(out_table, "ANALYSISNONE", "TEXT", "", "", 100, 'Analysis Result Message')

        rows = arcpy.InsertCursor(out_table)
        row = rows.newRow()
        row.setValue("ANALYSISNONE", "No features intersect the area of interest or buffer.")
        rows.insertRow(row)

        del row
        del rows
    except Exception as error:
        arcpy.AddError("Create Empty Output Error: {}".format(error))

# --------------------------------------
# Main
try:
    arcpy.AddMessage("Begin Analysis: {}".format(analysis_type))

    valid_inputs = validate_inputs()
    if not valid_inputs[0]:
        arcpy.AddError("Invalid inputs: {}".format(valid_inputs[1]))
        exit()

    if analysis_type == "Feature Comparison":

        aoi_layer_properties = arcpy.Describe(input_aoi)
        aoi_shape = aoi_layer_properties.shapeType

        analysis_layer_properties = arcpy.Describe(input_analysis_layer)
        analysis_shape = analysis_layer_properties.shapeType

        if aoi_shape in ["Point", "Polyline"]:
            # check if a buffer shape was provided

            if input_buffer_layer == "#":
                arcpy.AddWarning("For point and polyline areas of interest, a buffer layer required for "
                                 "Feature Comparison analyses")
            else:
                aoi_out = "empty"
                if analysis_shape == "Polygon":
                    buffer_area = get_area(input_buffer_layer, area_units)
                else:
                    buffer_area = 0  # for point and polyline anlaysis layers, the buffer_area is not used
                buffer_out = feature_comparison(input_analysis_layer, input_buffer_layer, interim_output_buffer,
                                                'Buffer', buffer_area)

        else:  # polygon analysis layer

            if analysis_shape == "Polygon":
                aoi_area = get_area(input_aoi, area_units)
            else:
                aoi_area = 0
            aoi_out = feature_comparison(input_analysis_layer, input_aoi, interim_output_aoi, 'AOI', aoi_area)

            if input_buffer_layer != "#":
                if analysis_shape == "Polygon":
                    buffer_area = get_area(input_buffer_layer, area_units)
                else:
                    buffer_area = 0
                buffer_out = feature_comparison(input_analysis_layer, input_buffer_layer, interim_output_buffer,
                                                'Buffer', buffer_area)
            else:
                arcpy.AddMessage("No buffer layer provided.")
                buffer_out = "empty"

    elif analysis_type == "Basic Proximity":
        aoi_out = basic_proximity(input_analysis_layer, input_aoi, interim_output_aoi, 'AOI')
        if input_buffer_layer != "#":
            buffer_out = basic_proximity(input_analysis_layer, input_buffer_layer, interim_output_buffer, 'Buffer')
        else:
            arcpy.AddMessage("No buffer layer provided.")
            buffer_out = "empty"
    else:
        aoi_out = distance_analysis_aoi(input_analysis_layer, input_aoi, interim_output_aoi)
        if input_buffer_layer != "#":
            buffer_out = distance_analysis_buffer(input_analysis_layer, input_aoi,
                                                  input_buffer_layer, interim_output_buffer)
        else:
            arcpy.AddMessage("No buffer layer provided.")
            buffer_out = "empty"

    arcpy.AddMessage("Creating output table.")
    if (aoi_out == "empty") & (buffer_out == "empty"):
        arcpy.AddMessage("No results found in the Area of Interest or Buffer locations.")
        create_empty_output(output_table)

    elif aoi_out == "empty":
        arcpy.AddMessage("No results found in the Area of Interest locations.")
        format_outputs(buffer_out, output_fields)

    elif buffer_out == "empty":
        arcpy.AddMessage("No results found in the Buffer location or no buffer provided.")
        format_outputs(aoi_out, output_fields)

    else:
        # got results from both AOI and Buffer results, merge them together before formatting...
        merged_output = output_workspace + "\\" + interim_output_merged
        arcpy.Merge_management([aoi_out, buffer_out], merged_output)
        arcpy.AddMessage("Results found for both the Area of Interest and Buffer locations.")
        format_outputs(merged_output, output_fields)

except Exception as err:
    arcpy.AddError("Impact Analysis Error: {}".format(err))

finally:
    # Clean up temporary files
    arcpy.AddMessage("Cleaning up interim results.")

    if arcpy.Exists(interim_output_aoi):
        arcpy.Delete_management(interim_output_aoi)
    if arcpy.Exists(interim_output_buffer):
        arcpy.Delete_management(interim_output_buffer)
    if arcpy.Exists(interim_output_merged):
        arcpy.Delete_management(interim_output_merged)
    if arcpy.Exists(interim_aoi_lines):
        arcpy.Delete_management(interim_aoi_lines)

