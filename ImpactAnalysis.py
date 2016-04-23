# python
# author: jess neuner

import arcpy
import sys
import os

# Anticipated inputs ...
# 1. Analysis Type: Choose from: Basic Proximity, Feature Comparison, Distance
# 2. Analysis Layer: Feature class to be analyzed
# 3. Output fields: from the selected input Analysis layer, which fields should be included in the output result
# 4. Area of Interest: Project location. Feature Class. Point, line or polygon.
# 5. Buffer shape (optional): pre-analyzed buffer for project area.
# 6. Output table: the GDB path and name of the final results table
# 7. Scratch workspace: folder to write interim results to

args = sys.argv

analysis_type = args[1]
input_analysis_layer = args[2]
output_fields = args[3]
input_aoi = args[4]
input_buffer_layer = args[5]
output_table = args[6]

interim_output_aoi = "interim_result_aoi"
interim_output_buffer = "interim_result_buffer"
interim_output_merged = "interim_result"

arcpy.env.overwriteOutput = True
output_workspace = arcpy.env.scratchWorkspace
arcpy.env.workspace = output_workspace


# --------------------------------------
# Function to validate script inputs
def validate_inputs():

    result = True
    reason = "Pass"

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

        for item in output_fields:
            if item.upper() in field_names:
                continue
            else:
                reason += "\n No field {} found in {}".format(item, input_analysis_layer)
    else:
        result = False
        reason += '\n {} does not exist'.format(input_analysis_layer)

    if arcpy.Exists(input_aoi):
        a = 1
    else:
        result = False
        reason += '\n {} does not exist'.format(input_aoi)

    if input_buffer_layer != "":
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
            arcpy.AddMessage(("no features found in {0}".format(analysis_layer)))
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
        arcpy.AddMessage("Analyzing layer: {}".format(analysis_layer))

        xy_tolerance = ""
        out_layer = output_workspace + "\\" + out_layer_name

        arcpy.Clip_analysis(analysis_layer, clip_layer, out_layer, xy_tolerance)

        match_count = int(arcpy.GetCount_management(out_layer_name)[0])

        if match_count == 0:
            arcpy.AddMessage(("No {0} features found in {1}.".format(analysis_layer, clip_layer)))
            return "empty"

        else:
            arcpy.AddMessage(("{0} features found in {1}. Adding more information.".format(match_count, analysis_layer)))

            arcpy.AddField_management(out_layer, "ANALYSISTYPE", "TEXT", "", "", 10, "Analysis Result Type")
            exp = "'{}'".format(layer_type)
            arcpy.CalculateField_management(out_layer, "ANALYSISTYPE", exp, "PYTHON_9.3", None)

            desc = arcpy.Describe(out_layer)
            shape_type = desc.shapeType

            if shape_type == "Polygon":
                arcpy.AddMessage("   Adding information for polygons")
                arcpy.AddField_management(out_layer, "ANALYSISAREA", "DOUBLE", "", "", "", "Total Area (acres)")
                exp = 'round(!shape.area@acres!,2)'
                arcpy.CalculateField_management(out_layer, "ANALYSISAREA", exp, "PYTHON_9.3", None)

                arcpy.AddField_management(out_layer, "ANALYSISPERCENT", "DOUBLE", "", "", "", "Percent of Area")
                exp = '(!ANALYSISAREA! / ' + str(clip_area) + ')*100'
                arcpy.CalculateField_management(out_layer, "ANALYSISPERCENT", exp, "PYTHON_9.3", None)

            elif shape_type == "Polyline":
                arcpy.AddMessage("Adding information for lines")
                arcpy.AddField_management(out_layer, "ANALYSISLEN", "DOUBLE", "", "", "", "Total Length")
                exp = 'round(!shape.length@miles!,2)'
                arcpy.CalculateField_management(out_layer, "ANALYSISLEN", exp, "PYTHON_9.3", None)

            elif shape_type == "Point":
                arcpy.AddMessage("   Adding information for points")
                arcpy.AddField_management(out_layer, "ANALYSISCOUNT", "SHORT", "", "", "", "Count of Features")
                exp = '1'
                arcpy.CalculateField_management(out_layer, "ANALYSISCOUNT", exp, "PYTHON_9.3", None)

            else:
                arcpy.AddMessage("type not supported")

            return out_layer
    except Exception as error:
        arcpy.AddError("Feature Comparison Error: {}".format(error))


# --------------------------------------
# Distance Analysis Type
def distance_analysis_aoi(near_layer, aoi_layer, out_layer_name):
    try:
                    # Find everything intersecting the AOI
            arcpy.MakeFeatureLayer_management(near_layer, "near_layer")
            arcpy.SelectLayerByLocation_management("near_layer", 'intersect', aoi_layer)

            match_count = int(arcpy.GetCount_management("near_layer")[0])

            if match_count == 0:
                arcpy.AddMessage(("no features found in {0}".format(aoi_layer)))
                return "empty"

            else:
                arcpy.AddMessage(("{0} features found in {1}".format(match_count, aoi_layer)))
                arcpy.CopyFeatures_management("near_layer", out_layer_name)

                desc = arcpy.Describe(aoi_layer)
                input_shape_type = desc.shapeType

                if input_shape_type == "Polygon":
                    arcpy.PolygonToLine_management(aoi_layer, "aoi_lines")
                    arcpy.Near_analysis(out_layer_name, "aoi_lines", None, "NO_LOCATION", "ANGLE", "GEODESIC")
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
# Distance Analysis Type
def distance_analysis_buffer(near_layer, aoi_layer, buffer_layer, out_layer_name):
    try:
        # Find everything intersecting the buffer
        arcpy.MakeFeatureLayer_management(near_layer, "near_layer")
        arcpy.SelectLayerByLocation_management("near_layer", 'intersect', buffer_layer)

        # Remove from the selection everything intersecting the AOI (these have already been accounted for)
        arcpy.SelectLayerByLocation_management("near_layer", 'intersect', aoi_layer, None, "REMOVE_FROM_SELECTION")
        match_count = int(arcpy.GetCount_management("near_layer")[0])

        if match_count == 0:
            arcpy.AddMessage(("no features found in {0}".format(buffer_layer)))
            return "empty"

        else:
            arcpy.AddMessage(("{0} features found in {1}".format(match_count, buffer_layer)))
            arcpy.CopyFeatures_management("near_layer", out_layer_name)

            desc = arcpy.Describe(aoi_layer)
            input_shape_type = desc.shapeType

            if input_shape_type == "Polygon":
                arcpy.PolygonToLine_management(aoi_layer, "aoi_lines")
                arcpy.Near_analysis(out_layer_name, "aoi_lines", None, "NO_LOCATION", "ANGLE", "GEODESIC")
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

        arcpy.AddMessage(("Total for {0}:\t {2}: {1}".format(input_fc, area, units)))

        return area
    except Exception as error:
        arcpy.AddError("Get Area Error: {}".format(error))


# --------------------------------------
# Function to format outputs
def format_outputs(output_layer, out_fields):

    try:
        arcpy.AddMessage("Only output fields requested... ")
        field_mapper = arcpy.FieldMappings()

        fields = arcpy.ListFields(output_layer)
        for field in fields:

            if field.name in out_fields:
                # Keep Me
                fm = arcpy.FieldMap()
                fm.addInputField(output_layer, field.name)
                field_mapper.addFieldMap(fm)

            elif field.name in ['ANALYSISTYPE', 'ANALYSISLOC', 'ANALYSISPERCENT', 'ANALYSISAREA', 'ANALYSISLEN', 'ANALYSISCOUNT', 'NEAR_DIST', 'NEAR_ANGLE']:
                # Keep Me
                fm = arcpy.FieldMap()
                fm.addInputField(output_layer, field.name)
                field_mapper.addFieldMap(fm)

            elif (field.type.upper() in ["OID"]):
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

        # arcpy.AddMessage("path: {}    ....   name: {}".format(out_path, out_name))
        arcpy.TableToTable_conversion(output_layer, out_path, out_name, field_mapping=field_mapper)

        return True
    except Exception as error:
        arcpy.AddError("Format Outputs Error: {}".format(error))


# --------------------------------------
# Create the output table with information about there being no results
def create_empty_output(out_table):
    try:
        arcpy.AddMessage("Creating empty output table.")

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
    arcpy.AddMessage("{} Analysis started".format(analysis_type))

    valid_inputs = validate_inputs()
    if not valid_inputs[0]:
        arcpy.AddError("Invalid inputs: {}".format(valid_inputs[1]))
        exit()

    if analysis_type == "Feature Comparison":

        layer_properties = arcpy.Describe(input_aoi)
        shape = layer_properties.shapeType
        if shape in ["Point", "Polyline"]:
            # check if a buffer shape was provided
            if input_buffer_layer == "":
                arcpy.AddWarning("For point and polyline areas of interest, a buffer is also required for Feature Comparison analyses")
            else:
                aoi_out = "empty"
                buffer_area = get_area(input_buffer_layer, 'ACRES')
                buffer_out = feature_comparison(input_analysis_layer, input_buffer_layer, interim_output_buffer, 'Buffer',
                                        buffer_area)
        else:
            aoi_area = get_area(input_aoi, 'ACRES')
            aoi_out = feature_comparison(input_analysis_layer, input_aoi, interim_output_aoi, 'AOI', aoi_area)

            buffer_area = get_area(input_buffer_layer, 'ACRES')
            buffer_out = feature_comparison(input_analysis_layer, input_buffer_layer, interim_output_buffer, 'Buffer',
                                    buffer_area)

    elif analysis_type == "Basic Proximity":
        aoi_out = basic_proximity(input_analysis_layer, input_aoi, interim_output_aoi, 'AOI')
        buffer_out = basic_proximity(input_analysis_layer, input_buffer_layer, interim_output_buffer, 'Buffer')

    else:
        aoi_out = distance_analysis_aoi(input_analysis_layer, input_aoi, interim_output_aoi)
        buffer_out = distance_analysis_buffer(input_analysis_layer, input_aoi, input_buffer_layer, interim_output_buffer)

    arcpy.AddMessage("Consolidate results")
    if (aoi_out == "empty") & (buffer_out == "empty"):
        arcpy.AddMessage("create empty output")
        create_empty_output(output_table)

    elif aoi_out == "empty":
        arcpy.AddMessage("Format Report Outputs:")
        format_outputs(buffer_out, output_fields)
        arcpy.Delete_management(interim_output_buffer)

    elif buffer_out == "empty":
        arcpy.AddMessage("Format Report Outputs:")
        format_outputs(aoi_out, output_fields)
        arcpy.Delete_management(interim_output_aoi)

    else:
        # got results from both AOI and Buffer results, merge them together before formatting...
        merged_output = output_workspace + "\\" + interim_output_merged
        arcpy.Merge_management([aoi_out, buffer_out], merged_output)

        arcpy.AddMessage("Format Report Outputs:")
        format_outputs(merged_output, output_fields)

        arcpy.AddMessage("Clean Up")
        arcpy.Delete_management(interim_output_aoi)
        arcpy.Delete_management(interim_output_buffer)

    arcpy.AddMessage("{} Analysis completed".format(analysis_type))

except Exception as err:
    arcpy.AddError("Impact Analysis Error: {}".format(err))



