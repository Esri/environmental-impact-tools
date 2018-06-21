# Esri start of added imports
import sys, os, arcpy
# Esri end of added imports

# Esri start of added variables
g_ESRI_variable_1 = 'interim_result_aoi'
g_ESRI_variable_2 = 'interim_result_buffer'
g_ESRI_variable_3 = 'interim_result'
g_ESRI_variable_4 = 'interim_aoi_lines'
g_ESRI_variable_5 = 'interim_analysis_layer'
g_ESRI_variable_6 = 'interim_result_related'
g_ESRI_variable_7 = 'interim_output_intersect'
g_ESRI_variable_8 = 'basic_proximity'
g_ESRI_variable_9 = "'{}'"
g_ESRI_variable_10 = 'ANALYSISTYPE'
g_ESRI_variable_11 = 'ANALYSISAREA'
g_ESRI_variable_12 = 'ANALYSISPERCENT'
g_ESRI_variable_13 = 'ANALYSISLEN'
g_ESRI_variable_14 = 'ANALYSISCOUNT'
g_ESRI_variable_15 = 'near_layer'
g_ESRI_variable_16 = 'GEODESIC'
g_ESRI_variable_17 = 'ANALYSISLOC'
g_ESRI_variable_18 = 'NEAR_ANGLE'
g_ESRI_variable_19 = 'ANALYSISTYPE COUNT'
g_ESRI_variable_20 = 'NEAR_DIST'
g_ESRI_variable_21 = 'SUM_ANALYSISPERCENT'
g_ESRI_variable_22 = 'SUM_ANALYSISAREA'
g_ESRI_variable_23 = 'SUM_ANALYSISLEN'
g_ESRI_variable_24 = 'SUM_ANALYSISCOUNT'
g_ESRI_variable_25 = '#'
# Esri end of added variables

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
# 4. Related Table (optional): If the Analysis Layer has related tables,
#                         optionally choose a table to report information from
# 5. Related Field (optional): If a Related Table is selected, one field from the table to report values from
# -. Group like records:  merge records in the results so that only
#                         unique combinations of attributes appear in the output
# 6. Area of Interest: Project location. Feature Class. Point, line or polygon.
# 7. Buffer shape (optional): pre-analyzed buffer for project area.
# 8. Reporting Units: based on the analysis layer shape type
# 9. Output table: the GDB path and name of the final results table
# Interim results will be written to the Scratch workspace

args = sys.argv

analysis_type = args[1]
input_analysis_layer = args[2]
output_fields = args[3]
related_table = args[4]
related_field = args[5]
# group_output_records = args[4]
group_output_records = True  # This should just always happen for the released script
input_aoi = args[6]
input_buffer_layer = args[7]
reporting_units = args[8]
output_table = args[9]

# Termporary feature classes created during script execution
interim_output_aoi = g_ESRI_variable_1
interim_output_buffer = g_ESRI_variable_2
interim_output_merged = g_ESRI_variable_3
interim_aoi_lines = g_ESRI_variable_4
interim_analysis_key = g_ESRI_variable_5
interim_related_result = g_ESRI_variable_6
interim_output_intersect = g_ESRI_variable_7

# arcpy.env.overwriteOutput = True  # should be set by the user in Geoprocessing Options
output_workspace = arcpy.env.scratchWorkspace
arcpy.env.workspace = output_workspace

arcpy.AddMessage("Analysis Type: {}".format(analysis_type))
arcpy.AddMessage("Input analysis layer: {}".format(input_analysis_layer))
arcpy.AddMessage("Output fields: {}".format(output_fields))
arcpy.AddMessage("Related table: {}".format(related_table))
arcpy.AddMessage("Related field: {}".format(related_field))
arcpy.AddMessage("Group like records: {}".format(group_output_records))
arcpy.AddMessage("Input AOI: {}".format(input_aoi))
arcpy.AddMessage("Input buffer: {}".format(input_buffer_layer))
arcpy.AddMessage("Reporting units: {}".format(reporting_units))
arcpy.AddMessage("Output table: {}".format(output_table))
arcpy.AddMessage("- - - - - - - - - - - - - - - - - - - ")

# Get units of measure squared away
area_units = reporting_units
if reporting_units in ["Meters", "Kilometers"]:  # then we need a different unit for area calculations
    area_units = "SQUAREKILOMETERS"
elif reporting_units in ["Feet", "Miles"]:
    area_units = "ACRES"

# Related Table Variables
foreign_key = ""
primary_key = ""
aoi_key_values = ""
buffer_key_values = ""
origin_OID_field = ""
related_table_full_path = ""


# --------------------------------------
# Function to validate script inputs
def validate_inputs():

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

    if not arcpy.Exists(input_aoi):
        result = False
        reason += '\n {} does not exist'.format(input_aoi)

    if input_buffer_layer != g_ESRI_variable_25:
        if not arcpy.Exists(input_buffer_layer):
            result = False
            reason += '\n {} does not exist'.format(input_buffer_layer)

    return [result, reason]


# --------------------------------------
# Get related table information
def check_related_records(input_fc, rel_table):
    try:
        global foreign_key
        global primary_key
        global origin_OID_field
        global related_table_full_path

        arcpy.AddMessage("Collecting related table information")

        fc_desc = arcpy.Describe(input_fc)
        for j,rel in enumerate(fc_desc.relationshipClassNames):
            rel_desc = arcpy.Describe(os.path.dirname(fc_desc.catalogPath) + "\\" + rel)
            if rel_desc.isAttachmentRelationship:
                continue

            destination = rel_desc.destinationClassNames
            if rel_table in destination:
                for key_name, key_role, unk in rel_desc.originClassKeys:
                    # print "Name: {}      Role: {}".format(key_name, key_role)
                    if key_role == "OriginForeign":
                        foreign_key = key_name
                    elif key_role == "OriginPrimary":
                        primary_key = key_name

        # get the full path for the related table
        dir_name = os.path.dirname(fc_desc.catalogPath)
        dir_desc = arcpy.Describe(dir_name)
        if hasattr(dir_desc, "datasetType") and dir_desc.datasetType == "FeatureDataset":
            dir_name = os.path.dirname(dir_name)
        related_table_full_path = dir_name + "\\" + related_table

        # check if the primary key is the OID, if it is, we need to kick it back
        if fc_desc.hasOID:
            origin_OID_field = fc_desc.OIDFieldName
            if fc_desc.OIDFieldName.lower() == primary_key.lower():
                # need to skip this
                arcpy.AddWarning("Check Related Records: The primary key for the selected related table and "
                                 "field is the ObjectID. ObjectID field types are not supported as primary keys "
                                 "for {} Analysis. Related records will not be analyzed".format(analysis_type))
                foreign_key = ""
                primary_key = ""
                return ""

        # check if the primary key is a GlobalID, if it is, we need to copy the dataset preserving the global id value
        primary_key_field = arcpy.ListFields(input_fc, primary_key)[0]
        if primary_key_field.type == "GlobalID":
            field_mapping = arcpy.FiledMappings()
            for field in arcpy.ListFields(input_fc):
                if field.type != "OID" and field.type != "GlobalID":
                    fm = arcpy.FieldMap()
                    fm.addInputField(input_fc, field.name)
                    field_mapping.addFieldMap(fm)
            # create a mapping to map the GlobalID into
            fm = arcpy.FieldMap()
            fm.addInputField(input_fc, primary_key_field.name)
            fm.mergeRule = 'First'
            f_name = fm.outputField
            f_name.name = "ANALYSISKEY"
            f_name.aliasName = "Relationship Key Field"
            f_name.type = "Guid"
            fm.outputField = f_name
            field_mapping.addFieldMap(fm)

            arcpy.CopyFeatures_management(input_fc, interim_analysis_key)
            # Update the primary key field to the copy --
            primary_key = "ANALYSISKEY"
            return output_workspace + "\\" + interim_analysis_key

        else:
            return input_fc

    except Exception as error:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        arcpy.AddError("Check Related Records: Line {}: {}".format(exc_tb.tb_lineno, error))


# --------------------------------------
# Get key values for related records - currently not required
def get_key_values(selected_layer, analysis_layer, layer_type):
    try:
        global aoi_key_values
        global buffer_key_values

        arcpy.AddMessage("Collecting related record key values ({}).".format(layer_type))

        selected_layer = arcpy.Describe(selected_layer)
        selected_ids = selected_layer.FIDset
        selected_ids = selected_ids.split(';')
        selected_ids = ", ".join(selected_ids)
        # arcpy.AddMessage("Selected IDs: {}".format(selected_ids))
        if layer_type == "AOI":
            aoi_key_values = [row[0] for row in arcpy.da.SearchCursor(analysis_layer, primary_key,
                                                                      "{0} IN ({1})".format(origin_OID_field,
                                                                                            selected_ids))]
            # arcpy.AddMessage("Selected AOI primary key values: {}".format(aoi_key_values))
        else:
            buffer_key_values = [row[0] for row in arcpy.da.SearchCursor(analysis_layer, primary_key,
                                                                         "{0} IN ({1})".format(origin_OID_field,
                                                                                               selected_ids))]
            # arcpy.AddMessage("Selected Buffer primary key values: {}".format(buffer_key_values))

    except Exception as error:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        arcpy.AddError("Get Key Values: Line {}: {}".format(exc_tb.tb_lineno, error))


# --------------------------------------
# Basic Proximity Analysis Type
def basic_proximity(analysis_layer, select_by_layer, out_layer_name, layer_type):
    try:
        out_layer = output_workspace + "\\" + out_layer_name

        arcpy.MakeFeatureLayer_management(analysis_layer, g_ESRI_variable_8)
        arcpy.SelectLayerByLocation_management(g_ESRI_variable_8, 'intersect', select_by_layer)

        match_count = int(arcpy.GetCount_management(g_ESRI_variable_8)[0])

        if match_count == 0:
            arcpy.AddMessage(("No features found in {0}".format(analysis_layer)))
            return g_ESRI_variable_25

        else:
            # Removed Get Key Values call - if primary key is persisted through the result, this shouldn't be needed
            # # If a field from a related table has been identified, then get a list of key values
            # if primary_key:
            #     get_key_values(out_layer_name, analysis_layer, layer_type)

            arcpy.AddMessage(("{0} features found in {1}".format(match_count, analysis_layer)))

            arcpy.CopyFeatures_management(g_ESRI_variable_8, out_layer)
            arcpy.AddField_management(out_layer, "ANALYSISTYPE", "TEXT", "", "", 10, "Analysis Result Type")
            exp = g_ESRI_variable_9.format(layer_type)

            arcpy.CalculateField_management(out_layer, g_ESRI_variable_10, exp, "PYTHON_9.3", None)

        return out_layer

    except Exception as error:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        arcpy.AddError("Basic Proximity Error: Line {}: {}".format(exc_tb.tb_lineno, error))


# --------------------------------------
# Feature Comparison Analysis Type
def feature_comparison(analysis_layer, clip_layer, out_layer_name, layer_type, clip_area, aoi_shape_type):
    try:
        xy_tolerance = ""
        out_layer = output_workspace + "\\" + out_layer_name

        # at pro 2.2 and possibly back to 2.1.3 if input features do not intersect the extent of the clip feature
        # the clip tool will fail rather than throw the expected warning about empty output
        # try/except were added to avoid the failure
        # intersect does not throw the same failure so it is checked to verify that no features intersect the clip feature
        # so we can avoid being concerened about why the clip failed
        try:
          arcpy.Clip_analysis(analysis_layer, clip_layer, out_layer, xy_tolerance)
        except Exception as error:
          out_intersect_layer = interim_output_intersect
          r = arcpy.Intersect_analysis([analysis_layer, clip_layer], out_intersect_layer, "ONLY_FID", None, "INPUT")
          warnings = r.getMessages(1)
          # 000117 is the code for empty output
          if warnings and '000117' in warnings:
            pass

        # at Pro 1.3, if nothing intersects with the clip layer, no result is generated - account for this
        if not arcpy.Exists(out_layer_name):
            arcpy.AddMessage(("No {0} features found in {1} (No clip result found).".format(analysis_layer, clip_layer)))
            return g_ESRI_variable_25

        match_count = int(arcpy.GetCount_management(out_layer_name)[0])

        if match_count == 0:
            arcpy.AddMessage(("No {0} features found in {1}.".format(analysis_layer, clip_layer)))
            return g_ESRI_variable_25

        else:
            arcpy.AddMessage(("Processing {0} features in {1} within {2}.".format(match_count, analysis_layer, layer_type)))

            desc = arcpy.Describe(out_layer)
            analysis_shape_type = desc.shapeType

            if analysis_shape_type == "Polygon":

                arcpy.AddField_management(out_layer, "ANALYSISAREA", "DOUBLE", "", "", "",
                                          "Total Area ({})".format(reporting_units))
                exp = '!shape.geodesicArea@{}!'.format(reporting_units)
                arcpy.CalculateField_management(out_layer, g_ESRI_variable_11, exp, "PYTHON_9.3", None)

                arcpy.AddField_management(out_layer, "ANALYSISPERCENT", "DOUBLE", "", "", "", "Percent of Area")
                exp = '(!ANALYSISAREA! / ' + str(clip_area) + ')*100'
                arcpy.CalculateField_management(out_layer, g_ESRI_variable_12, exp, "PYTHON_9.3", None)

            elif analysis_shape_type == "Polyline":
                arcpy.AddField_management(out_layer, "ANALYSISLEN", "DOUBLE", "", "", "",
                                          "Total Length ({})".format(reporting_units))
                exp = '!shape.geodesicLength@{}!'.format(reporting_units)
                arcpy.CalculateField_management(out_layer, g_ESRI_variable_13, exp, "PYTHON_9.3", None)

            elif analysis_shape_type == "Point":
                arcpy.AddField_management(out_layer, "ANALYSISCOUNT", "SHORT", "", "", "", "Count of Features")
                exp = '1'
                arcpy.CalculateField_management(out_layer, g_ESRI_variable_14, exp, "PYTHON_9.3", None)

            else:
                arcpy.AddMessage("Shape type not supported: {}".format(analysis_shape_type))

            if aoi_shape_type == "Polygon":
                # If the AOI is a point or line, then the only results will be with the buffer,
                # don't need to differentiate between AOI or buffer in this case
                arcpy.AddField_management(out_layer, "ANALYSISTYPE", "TEXT", "", "", 10, "Analysis Result Type")
                exp = g_ESRI_variable_9.format(layer_type)
                arcpy.CalculateField_management(out_layer, g_ESRI_variable_10, exp, "PYTHON_9.3", None)

            return out_layer
    except Exception as error:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        arcpy.AddError("Feature Comparison Error: Line {}: {}".format(exc_tb.tb_lineno, error))


# --------------------------------------
# Distance Analysis Type -- for the Area of Interest layer
def distance_analysis_aoi(near_layer, aoi_layer, out_layer_name):
    try:
        # Find everything intersecting the AOI
        arcpy.MakeFeatureLayer_management(near_layer, g_ESRI_variable_15)
        arcpy.SelectLayerByLocation_management(g_ESRI_variable_15, 'intersect', aoi_layer)

        match_count = int(arcpy.GetCount_management(g_ESRI_variable_15)[0])

        if match_count == 0:
            arcpy.AddMessage(("No features found in {0}".format(aoi_layer)))
            return g_ESRI_variable_25

        else:
            arcpy.AddMessage(("{0} features found in {1}. Calculating distances.".format(match_count, aoi_layer)))
            arcpy.CopyFeatures_management(g_ESRI_variable_15, out_layer_name)

            desc = arcpy.Describe(aoi_layer)
            input_shape_type = desc.shapeType

            desc_near_layer = arcpy.Describe(out_layer_name).spatialReference
            global reporting_units
            reporting_units = desc_near_layer.linearUnitName

            if input_shape_type == "Polygon":
                arcpy.PolygonToLine_management(aoi_layer, interim_aoi_lines)
                arcpy.Near_analysis(out_layer_name, interim_aoi_lines, None, "NO_LOCATION", "ANGLE", g_ESRI_variable_16)
            else:
                arcpy.Near_analysis(out_layer_name, aoi_layer, None, "NO_LOCATION", "ANGLE", g_ESRI_variable_16)

            arcpy.AddField_management(out_layer_name, "ANALYSISLOC", "TEXT", "", "", 100, "Location")

            exp = "getValue(!NEAR_DIST!)"
            codeblock = """def getValue(dist):
                    dist = round(dist, 4)
                    if dist == 0:
                        return 'Intersecting with AOI boundary'
                    else:
                        return 'Within AOI'"""
            arcpy.CalculateField_management(out_layer_name, g_ESRI_variable_17, exp, "PYTHON_9.3", codeblock)

            exp = "getValue(!NEAR_DIST!, !NEAR_ANGLE!)"
            codeblock = """def getValue(dist, angle):
                    dist = round(dist, 4)
                    if dist == 0:
                        return 0
                    else:
                        return angle"""
            arcpy.CalculateField_management(out_layer_name, g_ESRI_variable_18, exp, "PYTHON_9.3", codeblock)

            return out_layer_name

    except Exception as error:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        arcpy.AddError("Distance Analysis (AOI) Error: Line {}: {}".format(exc_tb.tb_lineno, error))


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
                'meter': 'm',
                'us-foot': 'ft',
                'foot_us': 'ft',
                'feet': 'ft'
                }.get(lower_units, units)

    except Exception as error:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        arcpy.AddError("Abbreviate Units Error: Line {}: {}".format(exc_tb.tb_lineno, error))


# --------------------------------------
# Distance Analysis Type -- for the Buffer layer
def distance_analysis_buffer(near_layer, aoi_layer, buffer_layer, out_layer_name):
    try:
        # Find everything intersecting the buffer
        arcpy.MakeFeatureLayer_management(near_layer, g_ESRI_variable_15)
        arcpy.SelectLayerByLocation_management(g_ESRI_variable_15, 'intersect', buffer_layer)

        # Remove from the selection everything intersecting the AOI (these have already been accounted for)
        arcpy.SelectLayerByLocation_management(g_ESRI_variable_15, 'intersect', aoi_layer, None, "REMOVE_FROM_SELECTION")
        match_count = int(arcpy.GetCount_management(g_ESRI_variable_15)[0])

        if match_count == 0:
            arcpy.AddMessage(("No features found in {0}".format(buffer_layer)))
            return g_ESRI_variable_25

        else:
            arcpy.AddMessage(("{0} features found in {1}. Calculating distances.".format(match_count, buffer_layer)))
            arcpy.CopyFeatures_management(g_ESRI_variable_15, out_layer_name)

            desc = arcpy.Describe(aoi_layer)
            input_shape_type = desc.shapeType

            desc_near_layer = arcpy.Describe(out_layer_name).spatialReference
            global reporting_units
            reporting_units = desc_near_layer.linearUnitName

            if input_shape_type == "Polygon":
                arcpy.PolygonToLine_management(aoi_layer, interim_aoi_lines)
                arcpy.Near_analysis(out_layer_name, interim_aoi_lines, None, "NO_LOCATION", "ANGLE", g_ESRI_variable_16)
            else:
                arcpy.Near_analysis(out_layer_name, aoi_layer, None, "NO_LOCATION", "ANGLE", g_ESRI_variable_16)

            arcpy.AddField_management(out_layer_name, "ANALYSISLOC", "TEXT", "", "", 100, "Location")

            exp = "getValue(!NEAR_DIST!)"
            codeblock = """def getValue(dist):
                if dist > 0:
                    return 'Outside AOI, within buffer'
                else:
                    return 'Unexpected result'"""

            arcpy.CalculateField_management(out_layer_name, g_ESRI_variable_17, exp, "PYTHON_9.3", codeblock)

            return out_layer_name

    except Exception as error:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        arcpy.AddError("Distance Analysis (Buffer) Error: Line {}: {}".format(exc_tb.tb_lineno, error))


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
        exc_type, exc_obj, exc_tb = sys.exc_info()
        arcpy.AddError("Get Area Error: Line {}: {}".format(exc_tb.tb_lineno, error))


# --------------------------------------
# Function to format outputs
def format_outputs(output_layer, out_fields):

    try:

        field_mapper = arcpy.FieldMappings()
        fields = arcpy.ListFields(output_layer)
        compare_fields = []
        stat_fields = ""
        out_fields_lower = out_fields.lower()

        for field in fields:

            field_name_lower = field.name
            field_name_lower = field_name_lower.lower()
            if field_name_lower in out_fields_lower and field.type != "Geometry":
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

            elif field.name in primary_key:
                # Keep Me -- need this for related records
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

        # if related records need to be included -- break here and create another interim result before the final
        if primary_key:
            out_path = output_workspace
            out_name = interim_related_result
        else:
            out_path = os.path.dirname(os.path.abspath(output_table))
            out_name = output_table.split("\\")[-1]

        out_table = out_path + "\\" + out_name

        abbreviated_units = abbreviate_units(reporting_units)
        # Clean up the results by deleting or merging identical records
        if analysis_type == "Basic Proximity":
            # Nothing special, just remove duplicates
            if group_output_records:
                # Using the summary statistics tool because Remove Duplicates requires an Advanced License
                arcpy.Statistics_analysis(output_layer, out_table, g_ESRI_variable_19, compare_fields)
                arcpy.DeleteField_management(out_table, ["FREQUENCY"])
                arcpy.DeleteField_management(out_table, ["COUNT_ANALYSISTYPE"])
            else:
                arcpy.TableToTable_conversion(output_layer, out_path, out_name, field_mapping=field_mapper)
        elif analysis_type == "Distance":
            # Don't remove duplicates, if there are 5 eagle nests, i want to know there
            #                          are 5 and their unique distances from the project
            arcpy.TableToTable_conversion(output_layer, out_path, out_name, field_mapping=field_mapper)

            # Update field aliases to be readable and have distance units embedded
            arcpy.AlterField_management(out_table, g_ESRI_variable_20, None,
                                        "Distance ({})".format(abbreviated_units))
            arcpy.AlterField_management(out_table, g_ESRI_variable_18, None, "Direction")

        else:
            # Feature Comparison -- summarize the output layer based on the 'keep fields'
            # Possible Stat Fields .. 'ANALYSISPERCENT', 'ANALYSISAREA', 'ANALYSISLEN', 'ANALYSISCOUNT'
            if group_output_records:
                arcpy.Statistics_analysis(output_layer, out_table, stat_fields, compare_fields)
                arcpy.DeleteField_management(out_table, ["FREQUENCY"])
                if "ANALYSISPERCENT" in stat_fields:
                    arcpy.AlterField_management(out_table, g_ESRI_variable_21, "ANALYSISPERCENT", "Percent of Area")
                    arcpy.AlterField_management(out_table, g_ESRI_variable_22, "ANALYSISAREA",
                                                "Total Area ({})".format(abbreviated_units))
                elif "ANALYSISLEN" in stat_fields:
                    arcpy.AlterField_management(out_table, g_ESRI_variable_23, "ANALYSISLEN",
                                                "Total Length ({})".format(abbreviated_units))
                else:
                    arcpy.AddField_management(out_table, "ANALYSISCOUNT", "SHORT", "", "", "", "Count of Features")
                    exp = '!SUM_ANALYSISCOUNT!'
                    arcpy.CalculateField_management(out_table, g_ESRI_variable_14, exp, "PYTHON_9.3", None)
                    arcpy.DeleteField_management(out_table, g_ESRI_variable_24)
            else:
                arcpy.TableToTable_conversion(output_layer, out_path, out_name, field_mapping=field_mapper)

        # if related records need to be included -- add related values to result table
        if primary_key:
            arcpy.AddMessage("Adding related table values to result.")
            query_related_fields = []

            # create a new table with same fields as the interim result
            out_path = os.path.dirname(os.path.abspath(output_table))
            out_name = output_table.split("\\")[-1]
            # arcpy.AddMessage("New output table: {}\\{}".format(out_path, out_name))
            # arcpy.AddMessage("Template table: {}".format(interim_related_result))
            arcpy.CreateTable_management(out_path, out_name, interim_related_result, None)

            # Get the field properties from the related table, must get them all
            field_props = arcpy.ListFields(related_table_full_path)

            #  Add a field to the result table to contain the selected related table field values
            for field in field_props:
                # arcpy.AddMessage("CURRENT field: {}".format(field.name))
                if field.name in related_field:
                    # arcpy.AddMessage("found a related field -- adding it to output")
                    query_related_fields.append(field.name)
                    arcpy.AddField_management(output_table, field.name, field.type,
                                              field.precision, field.scale,
                                              field.length, field.aliasName,
                                              "NULLABLE", "NON_REQUIRED", None)

            # prep an insert cursor for the final results table
            result_fields = arcpy.ListFields(output_table)
            result_field_names = []
            empty_result_value = []
            for field in result_fields:
                # if the field is the OID skip it
                if field.type != "OID":
                    result_field_names.append(field.name)
                    # if the current field is the analysis type field or related table field -
                    #                                   then do not add an empty value for it
                    if field.name != "ANALYSISTYPE" and field.name not in related_field:
                        empty_result_value.append(None)

            # arcpy.AddMessage("Result fields: {}".format(result_field_names))
            # Prep parameters to be used during the process
            foreign_key_formatted = arcpy.AddFieldDelimiters(arcpy.Describe(related_table_full_path).path, foreign_key)
            foreign_key_type = arcpy.ListFields(related_table_full_path, foreign_key)[0].type

            # cycle through the interim result table
            with arcpy.da.SearchCursor(interim_related_result, "*") as records:
                fields = records.fields
                key_field_index = fields.index(primary_key)

                if "ANALYSISTYPE" in result_field_names:
                    analysis_type_index = fields.index("ANALYSISTYPE")

                for record in records:
                    # query the related table to get the list of related values
                    current_key = record[key_field_index]

                    if "ANALYSISTYPE" in result_field_names:
                        current_analysis_type = record[analysis_type_index]
                    if foreign_key_type == "String":
                        current_key = g_ESRI_variable_9.format(current_key)

                    whereclause = "{} in ({})".format(foreign_key_formatted, current_key)

                    # strip OID
                    origin_values = list(record)
                    origin_values.pop(0)
                    # arcpy.AddMessage("Current record: {}".format(origin_values))
                    # arcpy.AddMessage("Related fields for query: {}".format(query_related_fields))
                    with arcpy.da.SearchCursor(related_table_full_path, query_related_fields, whereclause) as related_values:
                        # arcpy.AddMessage("result values: ".format(related_values))
                        with arcpy.da.InsertCursor(output_table, result_field_names) as result_cursor:
                            # Loop through the related values list and insert them into the final results table
                            for i, value in enumerate(related_values):
                                related_value_only = list(empty_result_value)
                                # arcpy.AddMessage(related_value_only)
                                if i == 0:
                                    # add a record into the result table for the 'parent' record + one or many related values
                                    origin_values.extend(value)
                                    # arcpy.AddMessage("Values list: {}".format(origin_values))
                                    # arcpy.AddMessage("Origin Values: {}".format(origin_values))
                                    result_cursor.insertRow(origin_values)
                                else:
                                    # add a record into the result table for each additional related value
                                    value_list = list(value)
                                    if "ANALYSISTYPE" in result_field_names:
                                        # the index for the analysis type field will be off by 1 because it's based
                                        #                                  on the record that includes an OID field
                                        related_value_only.insert((analysis_type_index - 1), current_analysis_type)
                                    related_value_only.extend(value_list)
                                    # arcpy.AddMessage(related_value_only)
                                    result_cursor.insertRow(related_value_only)

        return True
    except Exception as error:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        arcpy.AddError("Format Outputs Error: Line {}: {}".format(exc_tb.tb_lineno, error))


# --------------------------------------
# Create the output table with information about there being no results
def create_empty_output(out_table, message_overwrite):
    try:

        out_path = os.path.dirname(os.path.abspath(out_table))
        out_name = out_table.split("\\")[-1]

        arcpy.CreateTable_management(out_path, out_name, None, None)
        arcpy.AddField_management(out_table, "ANALYSISNONE", "TEXT", "", "", 150, 'Analysis Result')

        rows = arcpy.InsertCursor(out_table)

        row = rows.newRow()

        if message_overwrite != "":
            row.setValue("ANALYSISNONE", message_overwrite)
        elif input_buffer_layer != g_ESRI_variable_25:
            row.setValue("ANALYSISNONE", "No features intersect the area of interest or buffer.")
        else:
            row.setValue("ANALYSISNONE", "No features intersect the area of interest.")
        rows.insertRow(row)

        del row
        del rows
    except Exception as error:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        arcpy.AddError("Create Empty Output Error: Line {}: {}".format(exc_tb.tb_lineno, error))

# --------------------------------------
# Main
try:
    arcpy.AddMessage("Begin Analysis: {}".format(analysis_type))
    valid_inputs = validate_inputs()
    output_message = ""

    if not valid_inputs[0]:
        arcpy.AddError("Invalid inputs: {}".format(valid_inputs[1]))
        exit()

    # if related fields are chosen, then we need to ensure the primary key is preserved
    if related_field != g_ESRI_variable_25 and related_field != "" and related_field != None:
        input_analysis_layer = check_related_records(input_analysis_layer, related_table)

    if analysis_type == "Feature Comparison":

        aoi_layer_properties = arcpy.Describe(input_aoi)
        aoi_shape = aoi_layer_properties.shapeType

        analysis_layer_properties = arcpy.Describe(input_analysis_layer)
        analysis_shape = analysis_layer_properties.shapeType

        if aoi_shape in ["Point", "Polyline"]:
            # check if a buffer shape was provided

            if input_buffer_layer == g_ESRI_variable_25:
                arcpy.AddWarning("For point and polyline areas of interest, a buffer layer is required for "
                                 "Feature Comparison analyses.")
                aoi_out = g_ESRI_variable_25
                buffer_out = g_ESRI_variable_25
                output_message = "For point and polyline areas of interest, a buffer layer is required for " \
                                 "Feature Comparison analyses."
            else:
                aoi_out = g_ESRI_variable_25
                if analysis_shape == "Polygon":
                    buffer_area = get_area(input_buffer_layer, area_units)
                else:
                    buffer_area = 0  # for point and polyline anlaysis layers, the buffer_area is not used
                buffer_out = feature_comparison(input_analysis_layer, input_buffer_layer, interim_output_buffer,
                                                'Buffer', buffer_area, aoi_shape)

        else:  # polygon area of interest layer layer

            if analysis_shape == "Polygon":
                aoi_area = get_area(input_aoi, area_units)
            else:
                aoi_area = 0
            aoi_out = feature_comparison(input_analysis_layer, input_aoi, interim_output_aoi, 'AOI', aoi_area, aoi_shape)
            if input_buffer_layer != g_ESRI_variable_25:
                if analysis_shape == "Polygon":
                    buffer_area = get_area(input_buffer_layer, area_units)
                else:
                    buffer_area = 0
                buffer_out = feature_comparison(input_analysis_layer, input_buffer_layer, interim_output_buffer,
                                                'Buffer', buffer_area, aoi_shape)
            else:
                arcpy.AddMessage("No buffer layer provided.")
                buffer_out = g_ESRI_variable_25

    elif analysis_type == g_ESRI_variable_25:
        aoi_out = basic_proximity(input_analysis_layer, input_aoi, interim_output_aoi, 'AOI')
        if input_buffer_layer != g_ESRI_variable_25:
            buffer_out = basic_proximity(input_analysis_layer, input_buffer_layer, interim_output_buffer, 'Buffer')
        else:
            arcpy.AddMessage("No buffer layer provided.")
            buffer_out = g_ESRI_variable_25
    else:
        aoi_out = distance_analysis_aoi(input_analysis_layer, input_aoi, interim_output_aoi)
        if input_buffer_layer != g_ESRI_variable_25:
            buffer_out = distance_analysis_buffer(input_analysis_layer, input_aoi,
                                                  input_buffer_layer, interim_output_buffer)
        else:
            arcpy.AddMessage("No buffer layer provided.")
            buffer_out = g_ESRI_variable_25

    arcpy.AddMessage("Creating output table.")
    if (aoi_out == g_ESRI_variable_25) & (buffer_out == g_ESRI_variable_25):
        arcpy.AddMessage("No results found in the Area of Interest or Buffer (if provided) locations.")
        create_empty_output(output_table, output_message)
    elif aoi_out == g_ESRI_variable_25:
        arcpy.AddMessage("No results found in the Area of Interest locations.")
        format_outputs(buffer_out, output_fields)
    elif buffer_out == g_ESRI_variable_25:
        arcpy.AddMessage("No results found in the Buffer location or no buffer provided.")
        format_outputs(aoi_out, output_fields)
    else:
        # got results from both AOI and Buffer results, merge them together before formatting...
        merged_output = output_workspace + "\\" + interim_output_merged
        arcpy.Merge_management([aoi_out, buffer_out], merged_output)
        arcpy.AddMessage("Results found for both the Area of Interest and Buffer locations.")
        format_outputs(merged_output, output_fields)

except Exception as err:
    exc_m_type, exc_m_obj, exc_m_tb = sys.exc_info()
    arcpy.AddError("Impact Analysis Error: Line {}: {}".format(exc_m_tb.tb_lineno, err))

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
    if arcpy.Exists(interim_analysis_key):
        arcpy.Delete_management(interim_analysis_key)
    if arcpy.Exists(interim_related_result):
        arcpy.Delete_management(interim_related_result)
    if arcpy.Exists(interim_output_intersect):
        arcpy.Delete_management(interim_output_intersect)

