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
import os, sys, json, datetime, textwrap, math

ANALYSIS_SUMMARY = 'analysis_summary'

TOTAL_VALUE = 'TOTAL: '

SPLIT_FIELD = 'ANALYSISTYPE'

AREA_FIELD = 'ANALYSISAREA'
LEN_FIELD = 'ANALYSISLEN'
COUNT_FIELD = 'ANALYSISCOUNT'
PERCENT_FIELD = 'ANALYSISPERCENT'
SUM_FIELDS = [AREA_FIELD, LEN_FIELD, COUNT_FIELD, PERCENT_FIELD]

NO_AOI_FIELD_NAME = 'Analysis Result'
NO_AOI_VALUE = 'No features intersect the area of interest.'

SPLIT_VAL_BUFFER = 'Buffer'
SPLIT_VAL_AOI = 'AOI'

BUFFER_TITLE = ' (within buffer)'

Y_MARGIN = .025
X_MARGIN = .06

NUM_DIGITS = '{0:.2f}'

class MockField:
    def __init__(self, name, alias, type):
        self.name = name
        self.aliasName = alias
        self.type = type 
    def name(self):
        return self.name
    def aliasName(self):
        self.aliasName
    def type(self):
        return self.type

class Table:
    def __init__(self, title, rows, fields):
        self.rows = rows
        self.fields = fields
        self.title = title
        self.row_count = len(self.rows)

        self.field_widths = []
        self.row_heights = []
        self.auto_adjust = []
        self.max_vals = []
        self.field_name_lengths = []
        self.adjust_header_columns = {}

        self.row_background = None
        self.content_display = None
        self.report_title = None
        self.table_title = None
        self.table_header_background = None
        self.field_name = None
        self.field_value = None

        self.overflow_rows = None
        self.is_overflow = False
        self.full_overflow = False
        self.remaining_height = None
        
        self.buffer_rows = None
        self.has_buffer_rows = False

        self.total_row_index = None

    def calc_widths(self):
        self.auto_adjust = []
        self.field_name_lengths = []
        
        for f in self.fields:
            self.field_name.text = f.aliasName
            w = self.field_name.elementWidth + (X_MARGIN * 3)
            self.field_name_lengths.append(w)
            self.field_widths.append(w)

        self.max_vals = self.get_max_vals(self.rows)

        x = 0
        for v in self.max_vals:
            length = self.field_name_lengths[x]
            self.field_value.text = v if not is_float(v) else NUM_DIGITS.format(float(v))
            potential_length = self.field_value.elementWidth + (X_MARGIN * 3)
            if potential_length > length:
                self.field_widths[x] = potential_length
            x += 1

        self.row_width = sum(self.field_widths)

        if self.row_width > self.content_display.elementWidth:
            self.auto_adjust = self.adjust_row_widths()        

    def get_max_vals(self, rows):
        vals = []
        indexes = []
        fr = True
        for r in rows:
            x = 0
            for v in r:
                if fr:
                    vals.append(v)
                    if x + 1 == len(r):
                        fr = False                    
                else:
                    if len(v) > len(vals[x]):
                        vals[x] = v
                x += 1  
        return vals

    def adjust_row_widths(self):
        auto_adjust = []
        display_width = self.content_display.elementWidth
        overflow_width = display_width - self.row_width
        large_field_widths = []
        idx = 0
        even_width = display_width / len(self.field_widths)
        for column_width in self.field_widths:
            if column_width > even_width:
                large_field_widths.append(column_width)
            else:
                large_field_widths.append(0)
                display_width -= self.field_widths[idx]
            idx += 1

        idx = 0
        sum_widths = sum(large_field_widths)
        pos_widths = [x for x in large_field_widths if x != 0 ]
        small_widths = {}
        for w in large_field_widths:
            if float(w)/float(sum_widths) > 0:
                new_w = display_width * (1/len(pos_widths))
                if new_w < w:
                    self.field_widths[idx] = new_w
                    auto_adjust.append(idx)
                else:
                    small_widths[idx] = new_w - w
            idx += 1

        x=0
        if len(small_widths) > 0:
            t=0
            for small_w in small_widths:
                large_field_widths[small_w] = 0
                t += small_widths[small_w]
            pos_widths = [x for x in large_field_widths if x != 0 ]
            new_adjust = t/len(pos_widths)
            idx = 0
            for w in large_field_widths:
                if w > 0:
                    self.field_widths[idx] += new_adjust

        self.row_width = sum(self.field_widths)
        return auto_adjust

    def calc_num_chars(self, fit_width, v):
        self.field_value.text = ""
        x = 0
        while self.field_value.elementWidth < fit_width and len(v) > x:
            self.field_value.text += str(v)[x]
            x += 1
        return x

    def calc_heights(self):
        h_height = self.field_name.elementHeight
        self.header_height = h_height + (Y_MARGIN * 2)

        c_height = self.field_value.elementHeight 
        self.row_height = c_height + (Y_MARGIN * 2)

        self.adjust_header_columns = {}
        adjust_header_columns_list = [] 

        row_heights = []
        if len(self.auto_adjust) > 0:
            row_heights = [self.row_height] * len(self.rows)
            for column_index in self.auto_adjust:
                col_width = self.field_widths[column_index]
                field_name_width = self.field_name_lengths[column_index]
                fit_width = col_width - (X_MARGIN * 3)
                if field_name_width > fit_width:
                    adjust_header_columns_list.append(column_index)
                long_val = self.max_vals[column_index]
                max_chars = self.calc_num_chars(fit_width, long_val)
                x = 0
                for row in self.rows:
                    v = str(row[column_index])
                    if len(v) > max_chars:
                        v = v if not is_float(v) else NUM_DIGITS.format(float(v))
                        wrapped_val = textwrap.wrap(v, max_chars)
                        wrapped_height = (len(wrapped_val) * (self.row_height - Y_MARGIN))
                        if wrapped_height > row_heights[x]:
                            row_heights[x] = wrapped_height
                        row[column_index] = '\n'.join(wrapped_val)
                    x += 1
        else:
            table_height = self.row_count * self.row_height
            row_heights = [self.row_height] * len(self.rows)

        if not self.is_overflow:
            header_height = self.header_height
            if len(adjust_header_columns_list) > 0:
                field_names = [f.aliasName for f in self.fields]
                for ac in adjust_header_columns_list:
                    field_name = field_names[ac]
                    col_width = self.field_widths[ac]
                    fit_width = col_width - (X_MARGIN * 3)
                    max_chars = self.calc_num_chars(fit_width, field_name)
                    wrapped_val = textwrap.wrap(field_name, max_chars)
                    wrapped_height = (len(wrapped_val) * (self.header_height - Y_MARGIN))
                    if wrapped_height > header_height:
                        header_height = wrapped_height
                    if wrapped_height > self.header_height:
                        self.adjust_header_columns[ac] = '\n'.join(wrapped_val)
            row_heights.insert(0, header_height)  
        table_height = sum(row_heights)
        self.row_heights = row_heights

        if self.remaining_height == None:
            self.remaining_height = self.content_display.elementHeight
        if not self.is_overflow:
            self.remaining_height -= (self.table_header_background.elementHeight + Y_MARGIN)
        if table_height > self.remaining_height:
            num_rows = 0
            if self.remaining_height > 0:
                if len(self.auto_adjust) > 0:
                    sum_height = 0
                    for height in self.row_heights:
                        if sum_height + height < self.remaining_height:
                            sum_height += height
                            num_rows += 1
                        else:
                            break
                    self.table_height = sum_height
                else:
                    num_rows = int(self.remaining_height / self.row_height)
                    self.table_height = self.row_height * num_rows
            if num_rows <= 1:
                self.full_overflow = True
                self.overflow_rows = self.rows
                self.rows = []
                self.row_count = len(self.rows)
            else:
                if self.is_overflow:
                    self.overflow_rows = self.rows[num_rows:]
                    self.rows = self.rows[:num_rows]
                else:
                    self.overflow_rows = self.rows[num_rows - 1:]
                    self.rows = self.rows[:num_rows - 1]
                self.row_count = len(self.rows)
                self.remaining_height -= self.table_height
                self.row_heights = self.row_heights[:num_rows]
            return True
        else:
            self.table_height = table_height
            self.remaining_height -= self.table_height
            return False
    
    def init_elements(self, elements, layout_type):
        #Get key elelments
        #TODO make sure we always check the underlying table name rather than the table_view name or alias name when we do that
        if layout_type == 'map' and self.title == ANALYSIS_SUMMARY:
            self.field_name = elements['FieldNameA']
            self.field_value = elements['FieldValueA']
            self.row_background = elements['EvenRowBackgroundA']
        else:
            self.field_name = elements['FieldName']
            self.field_value = elements['FieldValue']
            self.row_background = elements['EvenRowBackground']
        self.table_header_background = elements['TableHeaderBackground']
        self.table_title = elements['TableTitle']
        self.content_display = elements['ContentDisplayArea']

        #Map only elements
        if layout_type == 'map':
            self.report_title = elements['ReportTitle']
    
    def check_result_type(self):
        field_names = [f.name for f in self.fields]
        pop_split_field = True
        if SPLIT_FIELD in field_names:
            split_field_idx = field_names.index(SPLIT_FIELD)
            buffer_rows = []
            aoi_rows = []
            for r in self.rows:
                v = r[split_field_idx]
                if v == SPLIT_VAL_BUFFER:
                    r.pop(split_field_idx)
                    buffer_rows.append(r)
                elif v == SPLIT_VAL_AOI:
                    r.pop(split_field_idx)
                    aoi_rows.append(r)
            if len(buffer_rows) > 0:
                self.has_buffer_rows = True
                self.buffer_rows = buffer_rows
            if len(aoi_rows) == 0:
                if len(buffer_rows) == 0:
                    aoi_rows = self.rows
                else:
                    aoi_rows = [[NO_AOI_VALUE]]
                self.p_fields = self.fields
                if SPLIT_FIELD in field_names:
                    self.p_fields.pop(field_names.index(SPLIT_FIELD))
                self.fields = [MockField(NO_AOI_FIELD_NAME, NO_AOI_FIELD_NAME, 'String')]
                pop_split_field = False
            
            self.rows = aoi_rows
            if SPLIT_FIELD in field_names and pop_split_field:
                self.fields.pop(field_names.index(SPLIT_FIELD))

    def calc_totals(self):
        #ANALYSIS_FIELDS
        sum_indexes = []
        percent_idx = None
        for f in self.fields:
            if f.name in SUM_FIELDS:
                idx = self.fields.index(f)
                sum_indexes.append(idx)
                if f.name == PERCENT_FIELD:
                    percent_idx = sum_indexes.index(idx)
        sums = {}
        first_row = True
        if len(sum_indexes) > 0:
            for r in self.rows:
                for i in sum_indexes:
                    v = r[i]
                    new_v = float(v) if is_float(v) else int(v)
                    if first_row:
                        sums[i] = new_v
                    else:
                        sums[i] += new_v
                    if is_float(new_v):
                        r[i] = str(NUM_DIGITS.format(new_v))
                first_row = False
        num_sums = len(sums)
        if num_sums > 0:           
            total_row = [''] * (len(self.fields) - (num_sums + 1))
            total_row.append(TOTAL_VALUE)
            self.total_row_index = total_row.index(TOTAL_VALUE)
            i = 0
            for sum in sums:
                v = sums[sum]
                if not percent_idx == None and i == percent_idx:
                    v = str(NUM_DIGITS.format(float(v)))
                total_row.append(str(v))
                i += 1
            self.rows.append(total_row)
            
    def init_table(self, elements, remaining_height, layout_type):
        #locate placeholder elements
        self.init_elements(elements, layout_type)

        self.remaining_height = remaining_height

        self.check_result_type()

        if not self.is_overflow and not self.full_overflow:
            self.calc_totals()

        #Calculate the column/row widths and the row/table heights
        if len(self.field_widths) == 0:
            self.calc_widths()
        overflow = self.calc_heights()
        #self.row_heights = [math.ceil(x*100)/100 for x in self.row_heights]
        return overflow

class Report:
    def __init__(self, report_title, sub_title, logo, map, scale_unit, report_type, map_template, overflow_template):
        self.tables = []
        self.pdf_paths = []
        self.pdfs = []
        self.temp_files = []
        self.idx = 0
        self.cur_x = 0
        self.cur_y = 0
        self.base_y = None
        self.base_x = None
        self.remaining_height = None
        self.overflow_row = None
        self.overflow = False
        self.place_holder = None
        self.map_pagx = None
        self.overflow_pagx = None
        self.layout_type = None

        self.report_title = report_title 
        self.sub_title = sub_title
        self.logo = logo
        self.map = map
        self.scale_unit = scale_unit
        self.report_type = report_type
        self.map_template = map_template
        self.overflow_template = overflow_template
        self.page_num = 0
    
        self.aprx = arcpy.mp.ArcGISProject('CURRENT')
        self.temp_dir = self.aprx.homeFolder

        self.map_element_names =  ['horzLine', 'vertLine', 'FieldValue', 
                                    'FieldName', 'ReportType', 'TableHeaderBackground', 
                                    'TableTitle', 'EvenRowBackground', 'ContentDisplayArea',
                                    'ReportSubTitle', 'ReportTitle', 'ScaleBarM', 'ScaleBarKM',
                                    'Logo', 'PageNumber', 'ReportTitleFooter', 'MapFrame', 'Legend', 
                                    'LegendPlaceholder', 'Text', 'horzLineA', 'vertLineA', 'FieldValueA', 
                                    'FieldNameA', 'EvenRowBackgroundA']

        self.overflow_element_names =  ['horzLine', 'vertLine', 'FieldValue', 
                                        'FieldName', 'TableHeaderBackground', 'TableTitle', 
                                        'EvenRowBackground', 'ContentDisplayArea', 'ReportTitle',
                                        'Logo', 'PageNumber', 'ReportTitleFooter']
        
        self.optional_element_names = ['Logo']

        self.init_layouts(False)

    def update_time_stamp(self):
        self.time_stamp = datetime.datetime.now().strftime("%m%d%Y-%H%M%S.%f")

    def add_table(self, title, rows, fields):
        self.tables.append(Table(title, rows, fields))

    def __iter__(self):       
        return self

    def __next__(self):
        try:
            t = self.tables[self.idx]
        except IndexError:
            self.idx = 0
            raise StopIteration
        self.idx = self.idx + 1
        return t

    def init_layouts(self, is_overflow):
        self.elements = {}
        self.update_time_stamp()
        if not is_overflow:
            self.init_map_layout(self.map_template)
            self.elements[self.map_layout_name] = self.find_elements('map')
        self.init_overflow_layout(self.overflow_template)
        self.elements[self.overflow_layout_name] = self.find_elements('overflow')
    
    def init_map_layout(self, template):
        pagx = None
        if template in ['', '#', ' ', None]:
            self.map_layout_name = 'zzMapLayout{0}'.format(self.time_stamp)
            map_layout_json = '{"layoutDefinition": {"uRI": "CIMPATH=layout/maplayout2.xml", "name": "' + self.map_layout_name + '", "sourceModifiedTime": {"type": "TimeInstant"}, "type": "CIMLayout", "elements": [{"visible": true, "name": "ContentDisplayArea", "anchor": "TopLeftCorner", "rotationCenter": {"x": 0.5000000000000018, "y": 2.329098039215686}, "type": "CIMGraphicElement", "graphic": {"placement": "Unspecified", "symbol": {"symbol": {"symbolLayers": [{"miterLimit": 10, "type": "CIMSolidStroke", "width": 1, "color": {"values": [0, 0, 0, 100], "type": "CIMRGBColor"}, "lineStyle3D": "Strip", "enable": true, "joinStyle": "Round", "capStyle": "Round"}], "type": "CIMPolygonSymbol"}, "type": "CIMSymbolReference"}, "polygon": {"hasZ": true, "rings": [[[0.5000000000000018, 2.329098039215686, null], [8.000000000000005, 2.329098039215686, null], [8.000000000000005, 0.8125, null], [0.5000000000000018, 0.8125, null], [0.5000000000000018, 2.329098039215686, null]]]}, "blendingMode": "Alpha", "type": "CIMPolygonGraphic"}}, {"visible": true, "name": "horzLineA", "anchor": "TopLeftCorner", "rotationCenter": {"x": 0.5850836166495381, "y": 1.7078296192178088}, "type": "CIMGraphicElement", "graphic": {"placement": "Unspecified", "symbol": {"symbol": {"symbolLayers": [{"miterLimit": 10, "type": "CIMSolidStroke", "width": 0.5, "color": {"values": [0, 0, 0, 0], "type": "CIMRGBColor"}, "lineStyle3D": "Strip", "enable": true, "joinStyle": "Round", "capStyle": "Round"}], "type": "CIMLineSymbol"}, "type": "CIMSymbolReference"}, "blendingMode": "Alpha", "line": {"hasZ": true, "paths": [[[0.5850836166495381, 1.7078296192178088, null], [2.585083616649538, 1.7078296192178088, null]]]}, "type": "CIMLineGraphic"}}, {"visible": true, "name": "vertLineA", "anchor": "TopLeftCorner", "rotationCenter": {"x": 2.5850518767017094, "y": 1.707829619217808}, "type": "CIMGraphicElement", "graphic": {"placement": "Unspecified", "symbol": {"symbol": {"symbolLayers": [{"miterLimit": 10, "type": "CIMSolidStroke", "width": 0.5, "color": {"values": [0, 0, 0, 0], "type": "CIMRGBColor"}, "lineStyle3D": "Strip", "enable": true, "joinStyle": "Round", "capStyle": "Round"}], "type": "CIMLineSymbol"}, "type": "CIMSymbolReference"}, "blendingMode": "Alpha", "line": {"hasZ": true, "paths": [[[2.5850518767017094, 1.707829619217808, null], [2.585083616649538, 0.7083182510693704, null]]]}, "type": "CIMLineGraphic"}}, {"visible": true, "name": "FieldNameA", "anchor": "TopLeftCorner", "rotationCenter": {"x": 0.6661620507831318, "y": 1.659085561951601}, "lockedAspectRatio": true, "type": "CIMGraphicElement", "graphic": {"placement": "Unspecified", "symbol": {"symbol": {"extrapolateBaselines": true, "verticalGlyphOrientation": "Right", "symbol": {"symbolLayers": [{"color": {"values": [0, 0, 0, 0], "type": "CIMRGBColor"}, "enable": true, "type": "CIMSolidFill"}], "type": "CIMPolygonSymbol"}, "type": "CIMTextSymbol", "hinting": "Default", "lineGapType": "ExtraLeading", "textCase": "Normal", "fontEncoding": "Unicode", "fontType": "TTOpenType", "depth3D": 1, "horizontalAlignment": "Left", "fontFamilyName": "Arial", "fontStyleName": "Bold", "haloSize": 1, "ligatures": true, "height": 6, "billboardMode3D": "FaceNearPlane", "fontEffects": "Normal", "symbol3DProperties": {"scaleZ": 0, "dominantSizeAxis3D": "Z", "scaleY": 0, "rotationOrder3D": "XYZ", "type": "CIM3DSymbolProperties"}, "kerning": true, "letterWidth": 100, "textDirection": "LTR", "wordSpacing": 100, "verticalAlignment": "Top", "blockProgression": "TTB"}, "type": "CIMSymbolReference"}, "shape": {"x": 0.6661620507831318, "y": 1.659085561951601}, "blendingMode": "Alpha", "type": "CIMTextGraphic", "text": "Field Name"}}, {"visible": true, "name": "FieldValueA", "anchor": "TopLeftCorner", "rotationCenter": {"x": 0.6661620507831318, "y": 1.4493416074708056}, "lockedAspectRatio": true, "type": "CIMGraphicElement", "graphic": {"placement": "Unspecified", "symbol": {"symbol": {"extrapolateBaselines": true, "verticalGlyphOrientation": "Right", "symbol": {"symbolLayers": [{"color": {"values": [0, 0, 0, 100], "type": "CIMRGBColor"}, "enable": true, "type": "CIMSolidFill"}], "type": "CIMPolygonSymbol"}, "type": "CIMTextSymbol", "hinting": "Default", "lineGapType": "ExtraLeading", "textCase": "Normal", "fontEncoding": "Unicode", "fontType": "TTOpenType", "depth3D": 1, "horizontalAlignment": "Left", "fontFamilyName": "Arial", "fontStyleName": "Regular", "haloSize": 1, "ligatures": true, "height": 10, "billboardMode3D": "FaceNearPlane", "fontEffects": "Normal", "symbol3DProperties": {"scaleZ": 0, "dominantSizeAxis3D": "Z", "scaleY": 0, "rotationOrder3D": "XYZ", "type": "CIM3DSymbolProperties"}, "kerning": true, "letterWidth": 100, "textDirection": "LTR", "wordSpacing": 100, "verticalAlignment": "Top", "blockProgression": "TTB"}, "type": "CIMSymbolReference"}, "shape": {"x": 0.6661620507831318, "y": 1.4493416074708056}, "blendingMode": "Alpha", "type": "CIMTextGraphic", "text": "Field Value"}}, {"visible": true, "name": "MapFrame", "graphicFrame": {"backgroundSymbol": {"symbol": {"symbolLayers": [{"color": {"values": [255, 255, 255, 0], "type": "CIMRGBColor"}, "enable": true, "type": "CIMSolidFill"}], "effects": [{"method": "Square", "offset": 0, "option": "Fast", "count": 1, "type": "CIMGeometricEffectOffset"}], "type": "CIMPolygonSymbol"}, "type": "CIMSymbolReference"}, "borderSymbol": {"symbol": {"symbolLayers": [{"miterLimit": 10, "type": "CIMSolidStroke", "width": 0.7, "color": {"values": [0, 0, 0, 100], "type": "CIMRGBColor"}, "lineStyle3D": "Strip", "enable": true, "joinStyle": "Round", "capStyle": "Round"}], "effects": [{"method": "Square", "offset": 0, "option": "Fast", "count": 1, "type": "CIMGeometricEffectOffset"}], "type": "CIMLineSymbol"}, "type": "CIMSymbolReference"}, "type": "CIMGraphicFrame"}, "anchor": "BottomLeftCorner", "rotationCenter": {"x": 0.5184697855750491, "y": 3.2710412154489505}, "type": "CIMMapFrame", "frame": {"rings": [[[0.5184697855750491, 3.2710412154489505], [0.5184697855750491, 9.223456432964337], [5.537037037037043, 9.223456432964337], [5.537037037037043, 3.2710412154489505], [0.5184697855750491, 3.2710412154489505]]]}, "view": {"viewableObjectPath": "CIMPATH=map/map.xml", "timeDisplay": {"defaultTimeIntervalUnits": "esriTimeUnitsUnknown", "timeRelation": "esriTimeRelationOverlaps", "type": "CIMMapTimeDisplay"}, "viewingMode": "Map", "camera": {"x": 473990.30952676287, "scale": 636915.4936921077, "y": 610101.8446561435, "pitch": -90, "type": "CIMViewCamera"}, "type": "CIMMapView", "verticalExaggerationScaleFactor": 1}, "autoCamera": {"source": "None", "autoCameraType": "Extent", "marginType": "Percent", "type": "CIMAutoCamera"}}, {"visible": true, "name": "ReportHeaderBackground", "anchor": "BottomLeftCorner", "rotationCenter": {"x": 0.49618708959743785, "y": 9.792000000000002}, "type": "CIMGraphicElement", "graphic": {"placement": "Unspecified", "symbol": {"symbol": {"symbolLayers": [{"miterLimit": 10, "type": "CIMSolidStroke", "width": 1, "color": {"values": [0, 0, 0, 0], "type": "CIMRGBColor"}, "lineStyle3D": "Strip", "enable": true, "joinStyle": "Round", "capStyle": "Round"}, {"color": {"values": [11, 121.99610900878906, 192, 100], "type": "CIMRGBColor"}, "enable": true, "type": "CIMSolidFill"}], "type": "CIMPolygonSymbol"}, "type": "CIMSymbolReference"}, "polygon": {"hasZ": true, "rings": [[[0.49618708959743785, 10.5, null], [8.000000000000002, 10.5, null], [8.000000000000002, 9.792000000000002, null], [0.49618708959743785, 9.792000000000002, null], [0.49618708959743785, 10.5, null]]]}, "blendingMode": "Alpha", "type": "CIMPolygonGraphic"}}, {"visible": true, "name": "ReportType", "anchor": "BottomLeftCorner", "rotationCenter": {"x": 1.2780402526863508, "y": 10.193474522079494}, "lockedAspectRatio": true, "type": "CIMGraphicElement", "graphic": {"placement": "Unspecified", "symbol": {"symbol": {"extrapolateBaselines": true, "verticalGlyphOrientation": "Right", "symbol": {"symbolLayers": [{"color": {"values": [255, 255, 255, 255], "type": "CIMRGBColor"}, "enable": true, "type": "CIMSolidFill"}], "type": "CIMPolygonSymbol"}, "type": "CIMTextSymbol", "hinting": "Default", "lineGapType": "ExtraLeading", "textCase": "Normal", "fontEncoding": "Unicode", "fontType": "Unspecified", "depth3D": 1, "horizontalAlignment": "Left", "fontFamilyName": "Tahoma", "fontStyleName": "Regular", "haloSize": 1, "ligatures": true, "height": 14, "billboardMode3D": "FaceNearPlane", "fontEffects": "Normal", "symbol3DProperties": {"scaleZ": 0, "dominantSizeAxis3D": "Z", "scaleY": 0, "rotationOrder3D": "XYZ", "type": "CIM3DSymbolProperties"}, "kerning": true, "letterWidth": 100, "textDirection": "LTR", "wordSpacing": 100, "verticalAlignment": "Bottom", "blockProgression": "TTB"}, "type": "CIMSymbolReference"}, "shape": {"x": 1.2780402526863508, "y": 10.193474522079494}, "blendingMode": "Alpha", "type": "CIMTextGraphic", "text": "Report Type"}}, {"visible": true, "name": "ReportTitle", "anchor": "BottomLeftCorner", "rotationCenter": {"x": 1.2780402526863508, "y": 9.864047438746155}, "lockedAspectRatio": true, "type": "CIMGraphicElement", "graphic": {"placement": "Unspecified", "symbol": {"symbol": {"extrapolateBaselines": true, "verticalGlyphOrientation": "Right", "symbol": {"symbolLayers": [{"color": {"values": [255, 255, 255, 255], "type": "CIMRGBColor"}, "enable": true, "type": "CIMSolidFill"}], "type": "CIMPolygonSymbol"}, "type": "CIMTextSymbol", "hinting": "Default", "lineGapType": "ExtraLeading", "textCase": "Normal", "fontEncoding": "Unicode", "fontType": "Unspecified", "depth3D": 1, "horizontalAlignment": "Left", "fontFamilyName": "Tahoma", "fontStyleName": "Bold", "haloSize": 1, "ligatures": true, "height": 18, "billboardMode3D": "FaceNearPlane", "fontEffects": "Normal", "symbol3DProperties": {"scaleZ": 0, "dominantSizeAxis3D": "Z", "scaleY": 0, "rotationOrder3D": "XYZ", "type": "CIM3DSymbolProperties"}, "kerning": true, "letterWidth": 100, "textDirection": "LTR", "wordSpacing": 100, "verticalAlignment": "Bottom", "blockProgression": "TTB"}, "type": "CIMSymbolReference"}, "shape": {"x": 1.2780402526863508, "y": 9.864047438746155}, "blendingMode": "Alpha", "type": "CIMTextGraphic", "text": "Report Title"}}, {"visible": true, "name": "ReportSubTitle", "anchor": "BottomLeftCorner", "rotationCenter": {"x": 0.5526366601581318, "y": 9.372692066420667}, "lockedAspectRatio": true, "type": "CIMGraphicElement", "graphic": {"placement": "Unspecified", "symbol": {"symbol": {"extrapolateBaselines": true, "verticalGlyphOrientation": "Right", "symbol": {"symbolLayers": [{"color": {"values": [0, 0, 0, 255], "type": "CIMRGBColor"}, "enable": true, "type": "CIMSolidFill"}], "type": "CIMPolygonSymbol"}, "type": "CIMTextSymbol", "hinting": "Default", "lineGapType": "ExtraLeading", "textCase": "Normal", "fontEncoding": "Unicode", "fontType": "Unspecified", "depth3D": 1, "horizontalAlignment": "Left", "fontFamilyName": "Tahoma", "fontStyleName": "Regular", "haloSize": 1, "ligatures": true, "height": 16, "billboardMode3D": "FaceNearPlane", "fontEffects": "Normal", "symbol3DProperties": {"scaleZ": 0, "dominantSizeAxis3D": "Z", "scaleY": 0, "rotationOrder3D": "XYZ", "type": "CIM3DSymbolProperties"}, "kerning": true, "letterWidth": 100, "textDirection": "LTR", "wordSpacing": 100, "verticalAlignment": "Bottom", "blockProgression": "TTB"}, "type": "CIMSymbolReference"}, "shape": {"x": 0.5526366601581318, "y": 9.372692066420669}, "blendingMode": "Alpha", "type": "CIMTextGraphic", "text": "Report Sub-Title"}}, {"groupGap": 5, "type": "CIMLegend", "defaultPatchHeight": 12, "frame": {"rings": [[[5.605715670750312, 3.271041215448945], [5.605715670750312, 9.203861007744178], [7.906777877514093, 9.203861007744178], [7.906777877514093, 3.271041215448945], [5.605715670750312, 3.271041215448945]]]}, "horizontalPatchGap": 5, "anchor": "BottomLeftCorner", "name": "Legend", "horizontalItemGap": 5, "mapFrame": "MapFrame", "verticalPatchGap": 5, "verticalItemGap": 5, "autoFonts": true, "headingGap": 5, "autoAdd": false, "defaultPatchWidth": 24, "layerNameGap": 5, "scaleSymbols": true, "visible": true, "showTitle": true, "graphicFrame": {"backgroundSymbol": {"symbol": {"symbolLayers": [{"color": {"values": [255, 255, 255, 0], "type": "CIMRGBColor"}, "enable": true, "type": "CIMSolidFill"}], "effects": [{"method": "Square", "offset": 0, "option": "Fast", "count": 1, "type": "CIMGeometricEffectOffset"}], "type": "CIMPolygonSymbol"}, "type": "CIMSymbolReference"}, "borderSymbol": {"symbol": {"symbolLayers": [{"miterLimit": 10, "type": "CIMSolidStroke", "width": 1, "color": {"values": [255, 255, 255, 0], "type": "CIMRGBColor"}, "lineStyle3D": "Strip", "enable": true, "joinStyle": "Round", "capStyle": "Round"}], "effects": [{"method": "Square", "offset": 0, "option": "Fast", "count": 1, "type": "CIMGeometricEffectOffset"}], "type": "CIMLineSymbol"}, "type": "CIMSymbolReference"}, "type": "CIMGraphicFrame"}, "autoReorder": false, "rotationCenter": {"x": 5.605715670750312, "y": 3.271041215448945}, "titleGap": 5, "fittingStrategy": "AdjustColumnsAndSize", "title": "Legend", "autoVisibility": true, "textGap": 5, "minFontSize": 4}, {"lineSymbol": {"symbolName": "Symbol_1148", "symbol": {"symbolLayers": [{"miterLimit": 10, "type": "CIMSolidStroke", "width": 1, "color": {"values": [0, 0, 0, 100], "type": "CIMRGBColor"}, "lineStyle3D": "Strip", "enable": true, "joinStyle": "Round", "capStyle": "Round"}], "type": "CIMLineSymbol"}, "type": "CIMSymbolReference"}, "divisionMarkHeight": 7, "anchor": "BottomLeftCorner", "labelSymbol": {"symbolName": "Symbol_1149", "symbol": {"extrapolateBaselines": true, "verticalGlyphOrientation": "Right", "symbol": {"symbolLayers": [{"color": {"values": [0, 0, 0, 100], "type": "CIMRGBColor"}, "enable": true, "type": "CIMSolidFill"}], "type": "CIMPolygonSymbol"}, "fontEffects": "Normal", "hinting": "Default", "lineGapType": "ExtraLeading", "textCase": "Normal", "fontEncoding": "Unicode", "compatibilityMode": true, "fontType": "Unspecified", "depth3D": 1, "horizontalAlignment": "Center", "fontFamilyName": "Arial", "fontStyleName": "Regular", "haloSize": 1, "ligatures": false, "shadowColor": {"values": [0, 0, 0, 100], "type": "CIMRGBColor"}, "height": 10, "billboardMode3D": "FaceNearPlane", "type": "CIMTextSymbol", "kerning": true, "letterWidth": 100, "textDirection": "LTR", "drawSoftHyphen": true, "verticalAlignment": "Baseline", "blockProgression": "TTB", "wordSpacing": 100}, "type": "CIMSymbolReference"}, "type": "CIMScaleLine", "frame": {"rings": [[[0.5184697855750491, 2.7088648048181616], [0.5184697855750491, 3.0029186242626063], [5.486523985239858, 3.0029186242626063], [5.486523985239858, 2.7088648048181616], [0.5184697855750491, 2.7088648048181616]]]}, "divisionMarkSymbol": {"symbolName": "Symbol_1151", "symbol": {"symbolLayers": [{"miterLimit": 10, "type": "CIMSolidStroke", "width": 1, "color": {"values": [0, 0, 0, 100], "type": "CIMRGBColor"}, "lineStyle3D": "Strip", "enable": true, "joinStyle": "Round", "capStyle": "Round"}], "type": "CIMLineSymbol"}, "type": "CIMSymbolReference"}, "unitLabel": "Miles", "unitLabelPosition": "AfterLabels", "name": "ScaleBarM", "unitLabelSymbol": {"symbolName": "Symbol_1150", "symbol": {"extrapolateBaselines": true, "verticalGlyphOrientation": "Right", "symbol": {"symbolLayers": [{"color": {"values": [0, 0, 0, 100], "type": "CIMRGBColor"}, "enable": true, "type": "CIMSolidFill"}], "type": "CIMPolygonSymbol"}, "fontEffects": "Normal", "hinting": "Default", "lineGapType": "ExtraLeading", "textCase": "Normal", "fontEncoding": "Unicode", "compatibilityMode": true, "fontType": "Unspecified", "depth3D": 1, "horizontalAlignment": "Left", "fontFamilyName": "Arial", "fontStyleName": "Regular", "haloSize": 1, "ligatures": false, "shadowColor": {"values": [0, 0, 0, 100], "type": "CIMRGBColor"}, "height": 10, "billboardMode3D": "FaceNearPlane", "type": "CIMTextSymbol", "kerning": true, "letterWidth": 100, "textDirection": "LTR", "drawSoftHyphen": true, "verticalAlignment": "Center", "blockProgression": "TTB", "wordSpacing": 100}, "type": "CIMSymbolReference"}, "mapFrame": "MapFrame", "markPosition": "Above", "units": {"uwkid": 9093}, "labelGap": 3, "labelPosition": "Above", "labelFrequency": "DivisionsAndFirstMidpoint", "divisions": 2, "divisionsBeforeZero": 0, "numberFormat": {"useSeparator": true, "alignmentWidth": 12, "roundingOption": "esriRoundNumberOfDecimals", "type": "CIMNumericFormat", "roundingValue": 2, "alignmentOption": "esriAlignLeft"}, "division": 7.5, "visible": true, "subdivisions": 4, "graphicFrame": {"backgroundSymbol": {"symbol": {"symbolLayers": [{"color": {"values": [255, 255, 255, 0], "type": "CIMRGBColor"}, "enable": true, "type": "CIMSolidFill"}], "effects": [{"method": "Square", "offset": 0, "option": "Fast", "count": 1, "type": "CIMGeometricEffectOffset"}], "type": "CIMPolygonSymbol"}, "type": "CIMSymbolReference"}, "borderSymbol": {"symbol": {"symbolLayers": [{"miterLimit": 10, "type": "CIMSolidStroke", "width": 1, "color": {"values": [255, 255, 255, 0], "type": "CIMRGBColor"}, "lineStyle3D": "Strip", "enable": true, "joinStyle": "Round", "capStyle": "Round"}], "effects": [{"method": "Square", "offset": 0, "option": "Fast", "count": 1, "type": "CIMGeometricEffectOffset"}], "type": "CIMLineSymbol"}, "type": "CIMSymbolReference"}, "type": "CIMGraphicFrame"}, "subdivisionMarkSymbol": {"symbolName": "Symbol_1152", "symbol": {"symbolLayers": [{"miterLimit": 10, "type": "CIMSolidStroke", "width": 1, "color": {"values": [0, 0, 0, 100], "type": "CIMRGBColor"}, "lineStyle3D": "Strip", "enable": true, "joinStyle": "Round", "capStyle": "Round"}], "type": "CIMLineSymbol"}, "type": "CIMSymbolReference"}, "rotationCenter": {"x": 0.5184697855750491, "y": 2.7088648048181616}, "fittingStrategy": "AdjustDivision", "subdivisionMarkHeight": 5, "unitLabelGap": 3, "markFrequency": "DivisionsAndSubdivisions"}, {"visible": true, "name": "TableHeaderBackground", "anchor": "TopLeftCorner", "rotationCenter": {"x": 0.5688520734581082, "y": 2.05841128981763}, "type": "CIMGraphicElement", "graphic": {"placement": "Unspecified", "symbol": {"symbol": {"symbolLayers": [{"miterLimit": 10, "type": "CIMSolidStroke", "width": 1, "color": {"values": [0, 0, 0, 0], "type": "CIMRGBColor"}, "lineStyle3D": "Strip", "enable": true, "joinStyle": "Round", "capStyle": "Round"}, {"color": {"values": [11, 121.99610900878906, 192, 0], "type": "CIMRGBColor"}, "enable": true, "type": "CIMSolidFill"}], "type": "CIMPolygonSymbol"}, "type": "CIMSymbolReference"}, "polygon": {"hasZ": true, "rings": [[[0.5688520734581082, 2.05841128981763, null], [7.957740962346998, 2.05841128981763, null], [7.957740962346998, 1.6590855619516027, null], [0.5688520734581082, 1.6590855619516027, null], [0.5688520734581082, 2.05841128981763, null]]]}, "blendingMode": "Alpha", "type": "CIMPolygonGraphic"}}, {"visible": true, "name": "TableTitle", "anchor": "TopLeftCorner", "rotationCenter": {"x": 0.6661620507831318, "y": 2.0234756706046673}, "lockedAspectRatio": true, "type": "CIMGraphicElement", "graphic": {"placement": "Unspecified", "symbol": {"symbol": {"extrapolateBaselines": true, "verticalGlyphOrientation": "Right", "symbol": {"symbolLayers": [{"color": {"values": [11, 121.99610900878906, 192, 100], "type": "CIMRGBColor"}, "enable": true, "type": "CIMSolidFill"}], "type": "CIMPolygonSymbol"}, "type": "CIMTextSymbol", "hinting": "Default", "lineGapType": "ExtraLeading", "textCase": "Normal", "fontEncoding": "Unicode", "fontType": "TTOpenType", "depth3D": 1, "horizontalAlignment": "Left", "fontFamilyName": "Arial", "fontStyleName": "Regular", "haloSize": 1, "ligatures": true, "height": 14, "billboardMode3D": "FaceNearPlane", "fontEffects": "Normal", "symbol3DProperties": {"scaleZ": 0, "dominantSizeAxis3D": "Z", "scaleY": 0, "rotationOrder3D": "XYZ", "type": "CIM3DSymbolProperties"}, "kerning": true, "letterWidth": 100, "textDirection": "LTR", "wordSpacing": 100, "verticalAlignment": "Top", "blockProgression": "TTB"}, "type": "CIMSymbolReference"}, "shape": {"x": 0.6661620507831318, "y": 2.0234756706046673}, "blendingMode": "Alpha", "type": "CIMTextGraphic", "text": "Table Title"}}, {"visible": true, "name": "EvenRowBackgroundA", "anchor": "TopLeftCorner", "rotationCenter": {"x": 0.5947562356971572, "y": 1.208073935143589}, "type": "CIMGraphicElement", "graphic": {"placement": "Unspecified", "symbol": {"symbol": {"symbolLayers": [{"miterLimit": 10, "type": "CIMSolidStroke", "width": 1, "color": {"values": [0, 0, 0, 0], "type": "CIMRGBColor"}, "lineStyle3D": "Strip", "enable": true, "joinStyle": "Round", "capStyle": "Round"}, {"color": {"values": [243.311279296875, 0], "type": "CIMGrayColor"}, "enable": true, "type": "CIMSolidFill"}], "type": "CIMPolygonSymbol"}, "type": "CIMSymbolReference"}, "polygon": {"hasZ": true, "rings": [[[0.5947562356971572, 1.208073935143589, null], [2.5850836166495377, 1.208073935143589, null], [2.5850836166495377, 0.9441322702057002, null], [0.5947562356971572, 0.9441322702057002, null], [0.5947562356971572, 1.208073935143589, null]]]}, "blendingMode": "Alpha", "type": "CIMPolygonGraphic"}}, {"visible": true, "name": "PageNumber", "anchor": "BottomMidPoint", "rotationCenter": {"x": 4.248093544798721, "y": 0.5000000000000009}, "lockedAspectRatio": true, "type": "CIMGraphicElement", "graphic": {"placement": "Unspecified", "symbol": {"symbol": {"extrapolateBaselines": true, "verticalGlyphOrientation": "Right", "symbol": {"symbolLayers": [{"color": {"values": [130, 130, 130, 100], "type": "CIMRGBColor"}, "enable": true, "type": "CIMSolidFill"}], "type": "CIMPolygonSymbol"}, "type": "CIMTextSymbol", "hinting": "Default", "lineGapType": "ExtraLeading", "textCase": "Normal", "fontEncoding": "Unicode", "fontType": "Unspecified", "depth3D": 1, "horizontalAlignment": "Center", "fontFamilyName": "Tahoma", "fontStyleName": "Regular", "haloSize": 1, "ligatures": true, "height": 8, "billboardMode3D": "FaceNearPlane", "fontEffects": "Normal", "symbol3DProperties": {"scaleZ": 0, "dominantSizeAxis3D": "Z", "scaleY": 0, "rotationOrder3D": "XYZ", "type": "CIM3DSymbolProperties"}, "kerning": true, "letterWidth": 100, "textDirection": "LTR", "wordSpacing": 100, "verticalAlignment": "Bottom", "blockProgression": "TTB"}, "type": "CIMSymbolReference"}, "shape": {"x": 4.248093544798721, "z": 0, "y": 0.5000000000000009}, "blendingMode": "Alpha", "type": "CIMTextGraphic", "text": "Page 1 of #"}}, {"visible": true, "name": "FooterLine", "anchor": "BottomLeftCorner", "rotationCenter": {"x": 0.500000000000002, "y": 0.68}, "type": "CIMGraphicElement", "graphic": {"placement": "Unspecified", "symbol": {"symbol": {"symbolLayers": [{"miterLimit": 10, "type": "CIMSolidStroke", "width": 0.5, "color": {"values": [11, 121.99610900878906, 192, 100], "type": "CIMRGBColor"}, "lineStyle3D": "Strip", "enable": true, "joinStyle": "Round", "capStyle": "Round"}], "type": "CIMLineSymbol"}, "type": "CIMSymbolReference"}, "blendingMode": "Alpha", "line": {"hasZ": true, "paths": [[[0.500000000000002, 0.6800027946275057, null], [8.000000000000004, 0.68, null]]]}, "type": "CIMLineGraphic"}}, {"visible": true, "name": "ReportTitleFooter", "anchor": "BottomRightCorner", "rotationCenter": {"x": 8.000000000000005, "y": 0.5000000000000001}, "lockedAspectRatio": true, "type": "CIMGraphicElement", "graphic": {"placement": "Unspecified", "symbol": {"symbol": {"extrapolateBaselines": true, "verticalGlyphOrientation": "Right", "symbol": {"symbolLayers": [{"color": {"values": [130, 130, 130, 100], "type": "CIMRGBColor"}, "enable": true, "type": "CIMSolidFill"}], "type": "CIMPolygonSymbol"}, "type": "CIMTextSymbol", "hinting": "Default", "lineGapType": "ExtraLeading", "textCase": "Normal", "fontEncoding": "Unicode", "fontType": "Unspecified", "depth3D": 1, "horizontalAlignment": "Right", "fontFamilyName": "Tahoma", "fontStyleName": "Regular", "haloSize": 1, "ligatures": true, "height": 8, "billboardMode3D": "FaceNearPlane", "fontEffects": "Normal", "symbol3DProperties": {"scaleZ": 0, "dominantSizeAxis3D": "Z", "scaleY": 0, "rotationOrder3D": "XYZ", "type": "CIM3DSymbolProperties"}, "kerning": true, "letterWidth": 100, "textDirection": "LTR", "wordSpacing": 100, "verticalAlignment": "Bottom", "blockProgression": "TTB"}, "type": "CIMSymbolReference"}, "shape": {"x": 8.000000000000007, "y": 0.5000000000000001}, "blendingMode": "Alpha", "type": "CIMTextGraphic", "text": "Report Title"}}, {"visible": true, "name": "Logo", "anchor": "BottomLeftCorner", "rotationCenter": {"x": 0.5688520734581086, "y": 9.816655152432439}, "lockedAspectRatio": true, "type": "CIMGraphicElement", "graphic": {"pictureURL": "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAJAAAACQCAYAAADnRuK4AAAACXBIWXMAAAsTAAALEwEAmpwYAAAAIGNIUk0AAHolAACAgwAA+f8AAIDoAABSCAABFVgAADqXAAAXb9daH5AAAAVnSURBVHja7J3bkuMwCESDSv//y+zjzlbNOpYtQTc0z5OJgEODFF/M3T8y2VObTfyMqBITQAJm93caSZysO0BOsC4Dj591A2gXNBa0HjSYfAdEsxE0p5NmC2tmUabvTpPswhwQmBO+GNsa0AHyItCs+mfJMb39/VPg5HaAi3mEwi80BfIG0KzGwJDjOwUOvCrtUqMjSoGgQC5wjiuSv4QZUoEEznNFsgB4oIdoFzivQLrb1o63l5EAjuDZA9I3QDyi0EcwPBCyWwQiu4hVWPwGIDyraiU1+jdWHpm7AQrPrs93gPCbGh21CQzOzu+lOdlls0ECj9/8G9/wN4zmWf9/EDnjF0A4WMArwXPdPzefRN/doqMm0ATOWnxGAjzIiXLBkzMD+Q1wjKTaXfDkD9HsLcNB1+RocRkBwbZgIG3Td7qAPq9AaPB8a5uMiYNuqSMZnohbc1gPDynOrEYiPJFzlH20RT+ytpmUUJSh3QF9oTrknJsdjL5S7u7NfHbxeUesbBYbBZ30xXnCBE8MQLvkPnM4dEDYGeHxXdv4J/AgAqFZ57ACvZ17TlYadQWz2ywSrLfXW/92VZ8JnD0APVUf1mAZcSFFq6hNBehvMNSy9ivQHfXRnRP44Bw77xovFiRrqjp3AZKyCJ4jCiT14YLHMgBifF6PfpJI8HcSJulqTZaU1LbtfmwKhgFX/c7LXAXPpl0Yq+0GyQn8PWoTaTEvE2nBvmiX+osCKSh14LEMgLR1r6k8ng2QTPYVoLeUdlApFvWxDIB2LCbr9hkTPLW28Wz3YVWCxzIAcmKQTPDUUyDt4pqo5yBPvNQnOV/M23ipGoBFvSvDD8Bz+jXbUp+bAFXYYRg48GV3jzOAaGetLqlPjRnIqlWttvEalBliFFIIkxgewdVAgdjagXZeRVqY1IWkEHQ9kNQHGiApSXxcHQEgkxrIUFoYgwpVud7HKwKkVlY4jjPY+ehXM3YC16sD9DOhUU9rrfSyXUgfRmIwvl3u6hsrz8nnH9sQgyPrGuSV5Yf+Vq0LXIGy5JkRIuj2O4ACvfpO1Q5nRfAxYP0pw5M+K3iSd2Focq331zdUIC8MDtP6jREgzT16b7zgqQIPE0CCRzOQ7CU8jrje8Um+ql9GBw+lApngwbRJEGRvDI9mIMHzCh5HX/sArobuMxiF/0NJ1NyjbbzmnrT1CyDNPdsUSOdBgmfZBymQhmbNQJp7cABSG1PrWvJjKLGCRy1MllYIY+EDamNSHymQ4Nnry+rzgaRCMhoFYtz+tlKfbwBJhdSG6WcgKxj0UgX49H1hHpwYAwe00mmz7QQIzTH2RHkleO4ChDYLGVgiWz//cZBWFGPSyqnPCkBWxWGBvNePseFLXAmkVp9XMRwNAnQqgHry6wOA9MDwptv1nQqkE+q14HtVeHa2sKxg6WeV5NY7Diyi+qN0mdVn+wgykBYTrD7d1OtIvsahRTk4PGpbQDNQNEQs505eHZ6dQ3QERH44KRWPKI77NAIW6+DgVG1ZIQURdTmHB3/uRMUaETxhNgIX72Dw2MFEZL2sJhzyeTAx/p/AGniFVr6YnwKgn075RXVahwBX92sGOHj1PlMDDqgDAwiznhnk7CmIul0dAOcvwnvjXfBgDshIAN2BaOWd7yZw+rSwuzu0uy0NHR4/WGitFeiJGnW5cN4+pPe9TYDE+EIlt70DVAA926WxtqE2M9sEq8ZKP5q2GPQHYNBXd2uCQAr0aj5SkqVAjxUpQ5VkBAp01UoYVMkEED5MvrCLUpsTQF+rfAWo01C1AdjcS44PDgi4FIhYnU6D1bZFzmb+rjxJVXNTgW181/YngGRqYV1NrWvB/gwAupFXbGGGEi4AAAAASUVORK5CYII=", "placement": "BottomLeftCorner", "blendingMode": "Alpha", "type": "CIMPictureGraphic", "frame": {"backgroundSymbol": {"symbol": {"symbolLayers": [{"color": {"values": [255, 255, 255, 0], "type": "CIMRGBColor"}, "enable": true, "type": "CIMSolidFill"}], "effects": [{"method": "Square", "offset": 0, "option": "Fast", "count": 1, "type": "CIMGeometricEffectOffset"}], "type": "CIMPolygonSymbol"}, "type": "CIMSymbolReference"}, "borderSymbol": {"symbol": {"symbolLayers": [{"miterLimit": 10, "type": "CIMSolidStroke", "width": 1, "color": {"values": [255, 255, 255, 0], "type": "CIMRGBColor"}, "lineStyle3D": "Strip", "enable": true, "joinStyle": "Round", "capStyle": "Round"}], "effects": [{"method": "Square", "offset": 0, "option": "Fast", "count": 1, "type": "CIMGeometricEffectOffset"}], "type": "CIMLineSymbol"}, "type": "CIMSymbolReference"}, "type": "CIMGraphicFrame"}, "sourceURL": "C:\\\\data\\\\State\\\\env impact\\\\logo_white.png", "box": {"zmin": 0, "ymin": 9.816655152432439, "xmin": 0.5688520734581086, "zmax": 0, "ymax": 10.461954490810816, "xmax": 1.4152942010810814}}}, {"lineSymbol": {"symbolName": "Symbol_1221", "symbol": {"symbolLayers": [{"miterLimit": 10, "type": "CIMSolidStroke", "width": 1, "color": {"values": [0, 0, 0, 100], "type": "CIMRGBColor"}, "lineStyle3D": "Strip", "enable": true, "joinStyle": "Round", "capStyle": "Round"}], "type": "CIMLineSymbol"}, "type": "CIMSymbolReference"}, "divisionMarkHeight": 7, "anchor": "BottomLeftCorner", "labelSymbol": {"symbolName": "Symbol_1222", "symbol": {"extrapolateBaselines": true, "verticalGlyphOrientation": "Right", "symbol": {"symbolLayers": [{"color": {"values": [0, 0, 0, 100], "type": "CIMRGBColor"}, "enable": true, "type": "CIMSolidFill"}], "type": "CIMPolygonSymbol"}, "fontEffects": "Normal", "hinting": "Default", "lineGapType": "ExtraLeading", "textCase": "Normal", "fontEncoding": "Unicode", "compatibilityMode": true, "fontType": "Unspecified", "depth3D": 1, "horizontalAlignment": "Center", "fontFamilyName": "Arial", "fontStyleName": "Regular", "haloSize": 1, "ligatures": false, "shadowColor": {"values": [0, 0, 0, 100], "type": "CIMRGBColor"}, "height": 10, "billboardMode3D": "FaceNearPlane", "type": "CIMTextSymbol", "kerning": true, "letterWidth": 100, "textDirection": "LTR", "drawSoftHyphen": true, "verticalAlignment": "Baseline", "blockProgression": "TTB", "wordSpacing": 100}, "type": "CIMSymbolReference"}, "type": "CIMScaleLine", "frame": {"rings": [[[0.5184697855750491, 2.7088648048181616], [0.5184697855750491, 3.0029186242626063], [5.486523985239859, 3.0029186242626063], [5.486523985239859, 2.7088648048181616], [0.5184697855750491, 2.7088648048181616]]]}, "divisionMarkSymbol": {"symbolName": "Symbol_1224", "symbol": {"symbolLayers": [{"miterLimit": 10, "type": "CIMSolidStroke", "width": 1, "color": {"values": [0, 0, 0, 100], "type": "CIMRGBColor"}, "lineStyle3D": "Strip", "enable": true, "joinStyle": "Round", "capStyle": "Round"}], "type": "CIMLineSymbol"}, "type": "CIMSymbolReference"}, "unitLabel": "Kilometers", "unitLabelPosition": "AfterLabels", "name": "ScaleBarKM", "unitLabelSymbol": {"symbolName": "Symbol_1223", "symbol": {"extrapolateBaselines": true, "verticalGlyphOrientation": "Right", "symbol": {"symbolLayers": [{"color": {"values": [0, 0, 0, 100], "type": "CIMRGBColor"}, "enable": true, "type": "CIMSolidFill"}], "type": "CIMPolygonSymbol"}, "fontEffects": "Normal", "hinting": "Default", "lineGapType": "ExtraLeading", "textCase": "Normal", "fontEncoding": "Unicode", "compatibilityMode": true, "fontType": "Unspecified", "depth3D": 1, "horizontalAlignment": "Left", "fontFamilyName": "Arial", "fontStyleName": "Regular", "haloSize": 1, "ligatures": false, "shadowColor": {"values": [0, 0, 0, 100], "type": "CIMRGBColor"}, "height": 10, "billboardMode3D": "FaceNearPlane", "type": "CIMTextSymbol", "kerning": true, "letterWidth": 100, "textDirection": "LTR", "drawSoftHyphen": true, "verticalAlignment": "Center", "blockProgression": "TTB", "wordSpacing": 100}, "type": "CIMSymbolReference"}, "mapFrame": "MapFrame", "markPosition": "Above", "units": {"uwkid": 9036}, "labelGap": 3, "labelPosition": "Above", "labelFrequency": "DivisionsAndFirstMidpoint", "divisions": 2, "divisionsBeforeZero": 0, "numberFormat": {"useSeparator": true, "alignmentWidth": 12, "roundingOption": "esriRoundNumberOfDecimals", "type": "CIMNumericFormat", "roundingValue": 2, "alignmentOption": "esriAlignLeft"}, "division": 10, "visible": true, "subdivisions": 4, "graphicFrame": {"backgroundSymbol": {"symbol": {"symbolLayers": [{"color": {"values": [255, 255, 255, 0], "type": "CIMRGBColor"}, "enable": true, "type": "CIMSolidFill"}], "effects": [{"method": "Square", "offset": 0, "option": "Fast", "count": 1, "type": "CIMGeometricEffectOffset"}], "type": "CIMPolygonSymbol"}, "type": "CIMSymbolReference"}, "borderSymbol": {"symbol": {"symbolLayers": [{"miterLimit": 10, "type": "CIMSolidStroke", "width": 1, "color": {"values": [255, 255, 255, 0], "type": "CIMRGBColor"}, "lineStyle3D": "Strip", "enable": true, "joinStyle": "Round", "capStyle": "Round"}], "effects": [{"method": "Square", "offset": 0, "option": "Fast", "count": 1, "type": "CIMGeometricEffectOffset"}], "type": "CIMLineSymbol"}, "type": "CIMSymbolReference"}, "type": "CIMGraphicFrame"}, "subdivisionMarkSymbol": {"symbolName": "Symbol_1225", "symbol": {"symbolLayers": [{"miterLimit": 10, "type": "CIMSolidStroke", "width": 1, "color": {"values": [0, 0, 0, 100], "type": "CIMRGBColor"}, "lineStyle3D": "Strip", "enable": true, "joinStyle": "Round", "capStyle": "Round"}], "type": "CIMLineSymbol"}, "type": "CIMSymbolReference"}, "rotationCenter": {"x": 0.5184697855750491, "y": 2.7088648048181616}, "fittingStrategy": "AdjustDivision", "subdivisionMarkHeight": 5, "unitLabelGap": 3, "markFrequency": "DivisionsAndSubdivisions"}, {"visible": true, "name": "CurrentTime", "anchor": "BottomRightCorner", "rotationCenter": {"x": 7.9067778775140924, "y": 9.91867874084698}, "lockedAspectRatio": true, "type": "CIMGraphicElement", "graphic": {"placement": "Unspecified", "symbol": {"symbol": {"extrapolateBaselines": true, "verticalGlyphOrientation": "Right", "symbol": {"symbolLayers": [{"color": {"values": [255, 255, 255, 100], "type": "CIMRGBColor"}, "enable": true, "type": "CIMSolidFill"}], "type": "CIMPolygonSymbol"}, "type": "CIMTextSymbol", "hinting": "Default", "lineGapType": "ExtraLeading", "textCase": "Normal", "fontEncoding": "Unicode", "fontType": "Unspecified", "depth3D": 1, "horizontalAlignment": "Right", "fontFamilyName": "Tahoma", "fontStyleName": "Regular", "haloSize": 1, "ligatures": true, "height": 10, "billboardMode3D": "FaceNearPlane", "fontEffects": "Normal", "symbol3DProperties": {"scaleZ": 0, "dominantSizeAxis3D": "Z", "scaleY": 0, "rotationOrder3D": "XYZ", "type": "CIM3DSymbolProperties"}, "kerning": true, "letterWidth": 100, "textDirection": "LTR", "wordSpacing": 100, "verticalAlignment": "Bottom", "blockProgression": "TTB"}, "type": "CIMSymbolReference"}, "shape": {"x": 7.906777877514093, "y": 9.91867874084698}, "blendingMode": "Alpha", "type": "CIMTextGraphic", "text": " <dyn type=\\"date\\" format=\\"\\"/> <dyn type=\\"time\\" format=\\"\\"/>"}}, {"visible": true, "name": "Credits", "anchor": "BottomLeftCorner", "rotationCenter": {"x": 0.5184697855750491, "y": 3.0832236241919833}, "lockedAspectRatio": true, "type": "CIMGraphicElement", "graphic": {"placement": "Unspecified", "symbol": {"symbol": {"extrapolateBaselines": true, "verticalGlyphOrientation": "Right", "symbol": {"symbolLayers": [{"color": {"values": [0, 0, 0, 100], "type": "CIMRGBColor"}, "enable": true, "type": "CIMSolidFill"}], "type": "CIMPolygonSymbol"}, "type": "CIMTextSymbol", "hinting": "Default", "lineGapType": "ExtraLeading", "textCase": "Normal", "fontEncoding": "Unicode", "fontType": "Unspecified", "depth3D": 1, "horizontalAlignment": "Left", "fontFamilyName": "Tahoma", "fontStyleName": "Regular", "haloSize": 1, "ligatures": true, "height": 8, "billboardMode3D": "FaceNearPlane", "fontEffects": "Normal", "symbol3DProperties": {"scaleZ": 0, "dominantSizeAxis3D": "Z", "scaleY": 0, "rotationOrder3D": "XYZ", "type": "CIM3DSymbolProperties"}, "kerning": true, "letterWidth": 100, "textDirection": "LTR", "wordSpacing": 100, "verticalAlignment": "Bottom", "blockProgression": "TTB"}, "type": "CIMSymbolReference"}, "shape": {"x": 0.5184697855750491, "y": 3.0832236241919833}, "blendingMode": "Alpha", "type": "CIMTextGraphic", "text": "Credits: <dyn type=\\"mapFrame\\" name=\\"MapFrame\\" property=\\"credits\\"/>"}}, {"visible": true, "name": "NorthArrow", "graphicFrame": {"backgroundSymbol": {"symbol": {"symbolLayers": [{"color": {"values": [255, 255, 255, 0], "type": "CIMRGBColor"}, "enable": true, "type": "CIMSolidFill"}], "effects": [{"method": "Square", "offset": 0, "option": "Fast", "count": 1, "type": "CIMGeometricEffectOffset"}], "type": "CIMPolygonSymbol"}, "type": "CIMSymbolReference"}, "borderSymbol": {"symbol": {"symbolLayers": [{"miterLimit": 10, "type": "CIMSolidStroke", "width": 1, "color": {"values": [255, 255, 255, 0], "type": "CIMRGBColor"}, "lineStyle3D": "Strip", "enable": true, "joinStyle": "Round", "capStyle": "Round"}], "effects": [{"method": "Square", "offset": 0, "option": "Fast", "count": 1, "type": "CIMGeometricEffectOffset"}], "type": "CIMLineSymbol"}, "type": "CIMSymbolReference"}, "type": "CIMGraphicFrame"}, "anchor": "CenterPoint", "rotationCenter": {"x": 5.747237281504665, "y": 2.92294900620705}, "lockedAspectRatio": true, "type": "CIMMarkerNorthArrow", "northType": "TrueNorth", "pointSymbol": {"symbolName": "Symbol_1128", "symbol": {"angleAlignment": "Display", "scaleX": 1, "haloSize": 1, "symbolLayers": [{"fontFamilyName": "ESRI North", "fontStyleName": "Regular", "dominantSizeAxis3D": "Y", "size": 61.315343576137735, "anchorPoint": {"x": 0, "y": 0}, "type": "CIMCharacterMarker", "billboardMode3D": "FaceNearPlane", "respectFrame": true, "symbol": {"symbolLayers": [{"color": {"values": [0, 0, 0, 100], "type": "CIMRGBColor"}, "enable": true, "type": "CIMSolidFill"}], "type": "CIMPolygonSymbol"}, "characterIndex": 175, "anchorPointUnits": "Absolute", "scaleX": 1, "fontType": "Unspecified", "enable": true, "scaleSymbolsProportionally": true}], "type": "CIMPointSymbol"}, "type": "CIMSymbolReference"}, "frame": {"rings": [[[5.605715670750311, 2.6285598048887837], [5.605715670750311, 3.2170248048887835], [5.888432337416978, 3.2170248048887835], [5.888432337416978, 2.6285598048887837], [5.605715670750311, 2.6285598048887837]]]}, "mapFrame": "MapFrame"}, {"visible": true, "name": "Scale", "anchor": "BottomLeftCorner", "rotationCenter": {"x": 0.5184697855750491, "y": 2.494445221555451}, "lockedAspectRatio": true, "type": "CIMGraphicElement", "graphic": {"placement": "Unspecified", "symbol": {"symbol": {"extrapolateBaselines": true, "verticalGlyphOrientation": "Right", "symbol": {"symbolLayers": [{"color": {"values": [0, 0, 0, 100], "type": "CIMRGBColor"}, "enable": true, "type": "CIMSolidFill"}], "type": "CIMPolygonSymbol"}, "type": "CIMTextSymbol", "hinting": "Default", "lineGapType": "ExtraLeading", "textCase": "Normal", "fontEncoding": "Unicode", "fontType": "TTOpenType", "depth3D": 1, "horizontalAlignment": "Left", "fontFamilyName": "Tahoma", "fontStyleName": "Regular", "haloSize": 1, "ligatures": true, "height": 8, "billboardMode3D": "FaceNearPlane", "fontEffects": "Normal", "symbol3DProperties": {"scaleZ": 0, "dominantSizeAxis3D": "Z", "scaleY": 0, "rotationOrder3D": "XYZ", "type": "CIM3DSymbolProperties"}, "kerning": true, "letterWidth": 100, "textDirection": "LTR", "wordSpacing": 100, "verticalAlignment": "Bottom", "blockProgression": "TTB"}, "type": "CIMSymbolReference"}, "shape": {"x": 0.5184697855750491, "y": 2.494445221555451}, "blendingMode": "Alpha", "type": "CIMTextGraphic", "text": "Scale <dyn type=\\"mapFrame\\" name=\\"MapFrame\\" property=\\"scale\\" preStr=\\"1:\\"/>"}}, {"visible": true, "name": "EvenRowBackground", "anchor": "TopLeftCorner", "rotationCenter": {"x": 5.804481300528421, "y": 1.7511750121663399}, "type": "CIMGraphicElement", "graphic": {"placement": "Unspecified", "symbol": {"symbol": {"symbolLayers": [{"miterLimit": 10, "type": "CIMSolidStroke", "width": 1, "color": {"values": [0, 0, 0, 0], "type": "CIMRGBColor"}, "lineStyle3D": "Strip", "enable": true, "joinStyle": "Round", "capStyle": "Round"}, {"color": {"values": [243, 100], "type": "CIMGrayColor"}, "enable": true, "type": "CIMSolidFill"}], "type": "CIMPolygonSymbol"}, "type": "CIMSymbolReference"}, "polygon": {"hasZ": true, "rings": [[[5.804481300528421, 1.7511750121663399, null], [7.794808681480801, 1.7511750121663399, null], [7.794808681480801, 1.4872333472284511, null], [5.804481300528421, 1.4872333472284511, null], [5.804481300528421, 1.7511750121663399, null]]]}, "blendingMode": "Alpha", "type": "CIMPolygonGraphic"}}, {"visible": true, "name": "vertLine", "anchor": "TopLeftCorner", "rotationCenter": {"x": 6.940678889009295, "y": 2.247897992555129}, "type": "CIMGraphicElement", "graphic": {"placement": "Unspecified", "symbol": {"symbol": {"symbolLayers": [{"miterLimit": 10, "type": "CIMSolidStroke", "width": 0.5, "color": {"values": [78, 78, 78, 100], "type": "CIMRGBColor"}, "lineStyle3D": "Strip", "enable": true, "joinStyle": "Round", "capStyle": "Round"}], "type": "CIMLineSymbol"}, "type": "CIMSymbolReference"}, "blendingMode": "Alpha", "line": {"hasZ": true, "paths": [[[6.940678889009295, 2.247897992555129, null], [6.943303746474558, 0.2478979925551288, null]]]}, "type": "CIMLineGraphic"}}, {"visible": true, "name": "horzLine", "anchor": "TopLeftCorner", "rotationCenter": {"x": 5.951093886584002, "y": 2.2678411889275507}, "type": "CIMGraphicElement", "graphic": {"placement": "Unspecified", "symbol": {"symbol": {"symbolLayers": [{"miterLimit": 10, "type": "CIMSolidStroke", "width": 0.5, "color": {"values": [78, 78, 78, 100], "type": "CIMRGBColor"}, "lineStyle3D": "Strip", "enable": true, "joinStyle": "Round", "capStyle": "Round"}], "type": "CIMLineSymbol"}, "type": "CIMSymbolReference"}, "blendingMode": "Alpha", "line": {"hasZ": true, "paths": [[[5.951093886584002, 2.2678411889275507, null], [7.951093886584002, 2.2678411889275507, null]]]}, "type": "CIMLineGraphic"}}, {"visible": true, "name": "FieldName", "anchor": "TopLeftCorner", "rotationCenter": {"x": 5.875887115614395, "y": 2.202186638974349}, "lockedAspectRatio": true, "type": "CIMGraphicElement", "graphic": {"placement": "Unspecified", "symbol": {"symbol": {"extrapolateBaselines": true, "verticalGlyphOrientation": "Right", "symbol": {"symbolLayers": [{"color": {"values": [0, 0, 0, 100], "type": "CIMRGBColor"}, "enable": true, "type": "CIMSolidFill"}], "type": "CIMPolygonSymbol"}, "type": "CIMTextSymbol", "hinting": "Default", "lineGapType": "ExtraLeading", "textCase": "Normal", "fontEncoding": "Unicode", "fontType": "TTOpenType", "depth3D": 1, "horizontalAlignment": "Left", "fontFamilyName": "Arial", "fontStyleName": "Bold", "haloSize": 1, "ligatures": true, "height": 10, "billboardMode3D": "FaceNearPlane", "fontEffects": "Normal", "symbol3DProperties": {"scaleZ": 0, "dominantSizeAxis3D": "Z", "scaleY": 0, "rotationOrder3D": "XYZ", "type": "CIM3DSymbolProperties"}, "kerning": true, "letterWidth": 100, "textDirection": "LTR", "wordSpacing": 100, "verticalAlignment": "Top", "blockProgression": "TTB"}, "type": "CIMSymbolReference"}, "shape": {"x": 5.875887115614395, "y": 2.202186638974351}, "blendingMode": "Alpha", "type": "CIMTextGraphic", "text": "Field Name"}}, {"visible": true, "name": "FieldValue", "anchor": "TopLeftCorner", "rotationCenter": {"x": 5.875887115614395, "y": 1.9924426844935565}, "lockedAspectRatio": true, "type": "CIMGraphicElement", "graphic": {"placement": "Unspecified", "symbol": {"symbol": {"extrapolateBaselines": true, "verticalGlyphOrientation": "Right", "symbol": {"symbolLayers": [{"color": {"values": [0, 0, 0, 100], "type": "CIMRGBColor"}, "enable": true, "type": "CIMSolidFill"}], "type": "CIMPolygonSymbol"}, "type": "CIMTextSymbol", "hinting": "Default", "lineGapType": "ExtraLeading", "textCase": "Normal", "fontEncoding": "Unicode", "fontType": "TTOpenType", "depth3D": 1, "horizontalAlignment": "Left", "fontFamilyName": "Arial", "fontStyleName": "Regular", "haloSize": 1, "ligatures": true, "height": 10, "billboardMode3D": "FaceNearPlane", "fontEffects": "Normal", "symbol3DProperties": {"scaleZ": 0, "dominantSizeAxis3D": "Z", "scaleY": 0, "rotationOrder3D": "XYZ", "type": "CIM3DSymbolProperties"}, "kerning": true, "letterWidth": 100, "textDirection": "LTR", "wordSpacing": 100, "verticalAlignment": "Top", "blockProgression": "TTB"}, "type": "CIMSymbolReference"}, "shape": {"x": 5.875887115614395, "y": 1.9924426844935565}, "blendingMode": "Alpha", "type": "CIMTextGraphic", "text": "Field Value"}}], "page": {"guides": [{"position": 0.5, "orientation": "Horizontal", "type": "CIMGuide"}, {"position": 10.5, "orientation": "Horizontal", "type": "CIMGuide"}, {"position": 0.5, "orientation": "Vertical", "type": "CIMGuide"}, {"position": 8, "orientation": "Vertical", "type": "CIMGuide"}, {"position": 0.8125, "orientation": "Horizontal", "type": "CIMGuide"}, {"position": 4.25, "orientation": "Vertical", "type": "CIMGuide"}], "width": 8.5, "showGuides": true, "units": {"uwkid": 109008}, "type": "CIMPage", "height": 11, "smallestRulerDivision": 0, "showRulers": true}, "dateExported": {"start": 1462280613173, "type": "TimeInstant"}, "metadataURI": "CIMPATH=Metadata/4e33b9fcbaa0bd47ee4b7102894fbc51.xml", "datePrinted": {"start": 1461228038678, "type": "TimeInstant"}}, "version": "1.2.0", "build": 5023, "type": "CIMLayoutDocument"}'
            pagx = self.write_pagx(map_layout_json, self.map_layout_name)
        else:
            base_name = os.path.splitext(os.path.basename(template))[0]
            self.map_layout_name = base_name + str(self.time_stamp)
            pagx = self.update_template(template, self.map_layout_name, self.map_element_names)
        self.map_pagx = pagx
        self.temp_files.append(self.map_pagx)
        self.aprx.importDocument(self.map_pagx)

    def init_overflow_layout(self, template):
        pagx = None
        if template in ['', '#', ' ', None]:
            self.overflow_layout_name = 'zzOverflowLayout{0}'.format(self.time_stamp)
            overflow_layout_json = '{"build": 5023, "type": "CIMLayoutDocument", "layoutDefinition": {"uRI": "CIMPATH=layout/layout1.xml", "page": {"width": 8.5, "guides": [{"type": "CIMGuide", "orientation": "Horizontal", "position": 0.5}, {"type": "CIMGuide", "orientation": "Horizontal", "position": 10.5}, {"type": "CIMGuide", "orientation": "Vertical", "position": 0.5}, {"type": "CIMGuide", "orientation": "Vertical", "position": 8}, {"type": "CIMGuide", "orientation": "Horizontal", "position": 0.75}, {"type": "CIMGuide", "orientation": "Horizontal", "position": 10.0625}, {"type": "CIMGuide", "orientation": "Horizontal", "position": 9.75}, {"type": "CIMGuide", "orientation": "Vertical", "position": 4.25}], "smallestRulerDivision": 0, "showRulers": true, "units": {"uwkid": 109008}, "type": "CIMPage", "showGuides": true, "height": 11}, "metadataURI": "CIMPATH=Metadata/00d23df2f0543ef21d2e170bbcdaddc2.xml", "dateExported": {"type": "TimeInstant", "start": 1461228142077}, "name": "' + self.overflow_layout_name + '", "type": "CIMLayout", "elements": [{"visible": true, "graphic": {"blendingMode": "Alpha", "type": "CIMLineGraphic", "line": {"hasZ": true, "paths": [[[0.728072303850187, 9.409174941587972, null], [2.728072303850187, 9.409174941587972, null]]]}, "symbol": {"type": "CIMSymbolReference", "symbol": {"type": "CIMLineSymbol", "symbolLayers": [{"enable": true, "lineStyle3D": "Strip", "capStyle": "Round", "color": {"type": "CIMRGBColor", "values": [78, 78, 78, 100]}, "width": 0.5, "type": "CIMSolidStroke", "joinStyle": "Round", "miterLimit": 10}]}}, "placement": "Unspecified"}, "name": "horzLine", "type": "CIMGraphicElement", "anchor": "TopLeftCorner", "rotationCenter": {"x": 0.728072303850187, "y": 9.409174941587972}}, {"visible": true, "graphic": {"blendingMode": "Alpha", "type": "CIMLineGraphic", "line": {"hasZ": true, "paths": [[[1.7176573062754796, 9.38923174521555, null], [1.720282163740743, 7.38923174521555, null]]]}, "symbol": {"type": "CIMSymbolReference", "symbol": {"type": "CIMLineSymbol", "symbolLayers": [{"enable": true, "lineStyle3D": "Strip", "capStyle": "Round", "color": {"type": "CIMRGBColor", "values": [78, 78, 78, 100]}, "width": 0.5, "type": "CIMSolidStroke", "joinStyle": "Round", "miterLimit": 10}]}}, "placement": "Unspecified"}, "name": "vertLine", "type": "CIMGraphicElement", "anchor": "TopLeftCorner", "rotationCenter": {"x": 1.7176573062754796, "y": 9.38923174521555}}, {"visible": true, "graphic": {"blendingMode": "Alpha", "type": "CIMPolygonGraphic", "symbol": {"type": "CIMSymbolReference", "symbol": {"type": "CIMPolygonSymbol", "symbolLayers": [{"enable": true, "lineStyle3D": "Strip", "capStyle": "Round", "color": {"type": "CIMRGBColor", "values": [0, 0, 0, 100]}, "width": 1, "type": "CIMSolidStroke", "joinStyle": "Round", "miterLimit": 10}]}}, "placement": "Unspecified", "polygon": {"hasZ": true, "rings": [[[0.5039586141498891, 9.749999999999998, null], [8.000000000000004, 9.749999999999998, null], [8.000000000000004, 0.7500000000000018, null], [0.5039586141498891, 0.7500000000000018, null], [0.5039586141498891, 9.749999999999998, null]]]}}, "name": "ContentDisplayArea", "type": "CIMGraphicElement", "anchor": "TopLeftCorner", "rotationCenter": {"x": 0.5039586141498891, "y": 9.749999999999998}}, {"visible": true, "graphic": {"blendingMode": "Alpha", "type": "CIMLineGraphic", "line": {"hasZ": true, "paths": [[[0.500000000000002, 0.6800027946275057, null], [8.000000000000004, 0.68, null]]]}, "symbol": {"type": "CIMSymbolReference", "symbol": {"type": "CIMLineSymbol", "symbolLayers": [{"enable": true, "lineStyle3D": "Strip", "capStyle": "Round", "color": {"type": "CIMRGBColor", "values": [11, 121.99610900878906, 192, 100]}, "width": 0.5, "type": "CIMSolidStroke", "joinStyle": "Round", "miterLimit": 10}]}}, "placement": "Unspecified"}, "name": "FooterLine", "type": "CIMGraphicElement", "anchor": "BottomLeftCorner", "rotationCenter": {"x": 0.500000000000002, "y": 0.68}}, {"visible": true, "lockedAspectRatio": true, "name": "PageNumber", "type": "CIMGraphicElement", "graphic": {"text": "Page 1 of #", "shape": {"z": 0, "x": 4.251979307074945, "y": 0.5}, "symbol": {"type": "CIMSymbolReference", "symbol": {"blockProgression": "TTB", "fontEncoding": "Unicode", "wordSpacing": 100, "kerning": true, "extrapolateBaselines": true, "symbol3DProperties": {"rotationOrder3D": "XYZ", "type": "CIM3DSymbolProperties", "scaleZ": 0, "scaleY": 0, "dominantSizeAxis3D": "Z"}, "fontStyleName": "Regular", "verticalGlyphOrientation": "Right", "symbol": {"type": "CIMPolygonSymbol", "symbolLayers": [{"enable": true, "type": "CIMSolidFill", "color": {"type": "CIMRGBColor", "values": [130, 130, 130, 100]}}]}, "fontEffects": "Normal", "horizontalAlignment": "Center", "haloSize": 1, "lineGapType": "ExtraLeading", "hinting": "Default", "fontType": "Unspecified", "textCase": "Normal", "verticalAlignment": "Bottom", "fontFamilyName": "Tahoma", "textDirection": "LTR", "billboardMode3D": "FaceNearPlane", "type": "CIMTextSymbol", "height": 8, "letterWidth": 100, "depth3D": 1, "ligatures": true}}, "blendingMode": "Alpha", "type": "CIMTextGraphic", "placement": "Unspecified"}, "anchor": "BottomMidPoint", "rotationCenter": {"x": 4.251979307074945, "y": 0.5}}, {"visible": true, "lockedAspectRatio": true, "name": "ReportTitleFooter", "type": "CIMGraphicElement", "graphic": {"text": "Report Title", "shape": {"x": 8, "y": 0.5019678941441441}, "symbol": {"type": "CIMSymbolReference", "symbol": {"blockProgression": "TTB", "fontEncoding": "Unicode", "wordSpacing": 100, "kerning": true, "extrapolateBaselines": true, "symbol3DProperties": {"rotationOrder3D": "XYZ", "type": "CIM3DSymbolProperties", "scaleZ": 0, "scaleY": 0, "dominantSizeAxis3D": "Z"}, "fontStyleName": "Regular", "verticalGlyphOrientation": "Right", "symbol": {"type": "CIMPolygonSymbol", "symbolLayers": [{"enable": true, "type": "CIMSolidFill", "color": {"type": "CIMRGBColor", "values": [130, 130, 130, 100]}}]}, "fontEffects": "Normal", "horizontalAlignment": "Right", "haloSize": 1, "lineGapType": "ExtraLeading", "hinting": "Default", "fontType": "Unspecified", "textCase": "Normal", "verticalAlignment": "Bottom", "fontFamilyName": "Tahoma", "textDirection": "LTR", "billboardMode3D": "FaceNearPlane", "type": "CIMTextSymbol", "height": 8, "letterWidth": 100, "depth3D": 1, "ligatures": true}}, "blendingMode": "Alpha", "type": "CIMTextGraphic", "placement": "Unspecified"}, "anchor": "BottomRightCorner", "rotationCenter": {"x": 7.999999999999999, "y": 0.5019678941441441}}, {"visible": true, "graphic": {"blendingMode": "Alpha", "type": "CIMPolygonGraphic", "symbol": {"type": "CIMSymbolReference", "symbol": {"type": "CIMPolygonSymbol", "symbolLayers": [{"enable": true, "lineStyle3D": "Strip", "capStyle": "Round", "color": {"type": "CIMRGBColor", "values": [0, 0, 0, 0]}, "width": 0, "type": "CIMSolidStroke", "joinStyle": "Round", "miterLimit": 10}, {"enable": true, "type": "CIMSolidFill", "color": {"type": "CIMGrayColor", "values": [243, 100]}}]}}, "placement": "Unspecified", "polygon": {"hasZ": true, "rings": [[[0.5814597177946053, 8.892508764826761, null], [2.5717870987469857, 8.892508764826761, null], [2.5717870987469857, 8.628567099888873, null], [0.5814597177946053, 8.628567099888873, null], [0.5814597177946053, 8.892508764826761, null]]]}}, "name": "EvenRowBackground", "type": "CIMGraphicElement", "anchor": "TopLeftCorner", "rotationCenter": {"x": 0.5814597177946053, "y": 8.892508764826761}}, {"visible": true, "lockedAspectRatio": true, "name": "TableTitle", "type": "CIMGraphicElement", "graphic": {"text": "Table Title", "shape": {"x": 0.6528655328805799, "y": 9.707910500287841}, "symbol": {"type": "CIMSymbolReference", "symbol": {"blockProgression": "TTB", "fontEncoding": "Unicode", "wordSpacing": 100, "kerning": true, "extrapolateBaselines": true, "symbol3DProperties": {"rotationOrder3D": "XYZ", "type": "CIM3DSymbolProperties", "scaleZ": 0, "scaleY": 0, "dominantSizeAxis3D": "Z"}, "fontStyleName": "Regular", "verticalGlyphOrientation": "Right", "symbol": {"type": "CIMPolygonSymbol", "symbolLayers": [{"enable": true, "type": "CIMSolidFill", "color": {"type": "CIMRGBColor", "values": [11, 121.99610900878906, 192, 100]}}]}, "fontEffects": "Normal", "horizontalAlignment": "Left", "haloSize": 1, "lineGapType": "ExtraLeading", "hinting": "Default", "fontType": "TTOpenType", "textCase": "Normal", "verticalAlignment": "Top", "fontFamilyName": "Arial", "textDirection": "LTR", "billboardMode3D": "FaceNearPlane", "type": "CIMTextSymbol", "height": 14, "letterWidth": 100, "depth3D": 1, "ligatures": true}}, "blendingMode": "Alpha", "type": "CIMTextGraphic", "placement": "Unspecified"}, "anchor": "TopLeftCorner", "rotationCenter": {"x": 0.6528655328805799, "y": 9.70791050028784}}, {"visible": true, "graphic": {"blendingMode": "Alpha", "type": "CIMPolygonGraphic", "symbol": {"type": "CIMSymbolReference", "symbol": {"type": "CIMPolygonSymbol", "symbolLayers": [{"enable": true, "lineStyle3D": "Strip", "capStyle": "Round", "color": {"type": "CIMRGBColor", "values": [0, 0, 0, 0]}, "width": 1, "type": "CIMSolidStroke", "joinStyle": "Round", "miterLimit": 10}, {"enable": true, "type": "CIMSolidFill", "color": {"type": "CIMRGBColor", "values": [11, 121.99610900878906, 192, 0]}}]}}, "placement": "Unspecified", "polygon": {"hasZ": true, "rings": [[[0.5555555555555562, 9.742846119500802, null], [7.9444444444444455, 9.742846119500802, null], [7.9444444444444455, 9.265937926356994, null], [0.5555555555555562, 9.265937926356994, null], [0.5555555555555562, 9.742846119500802, null]]]}}, "name": "TableHeaderBackground", "type": "CIMGraphicElement", "anchor": "TopLeftCorner", "rotationCenter": {"x": 0.5555555555555562, "y": 9.742846119500802}}, {"visible": true, "lockedAspectRatio": true, "name": "FieldValue", "type": "CIMGraphicElement", "graphic": {"text": "Field Value", "shape": {"x": 0.6528655328805799, "y": 9.133776437153978}, "symbol": {"type": "CIMSymbolReference", "symbol": {"blockProgression": "TTB", "fontEncoding": "Unicode", "wordSpacing": 100, "kerning": true, "extrapolateBaselines": true, "symbol3DProperties": {"rotationOrder3D": "XYZ", "type": "CIM3DSymbolProperties", "scaleZ": 0, "scaleY": 0, "dominantSizeAxis3D": "Z"}, "fontStyleName": "Regular", "verticalGlyphOrientation": "Right", "symbol": {"type": "CIMPolygonSymbol", "symbolLayers": [{"enable": true, "type": "CIMSolidFill", "color": {"type": "CIMRGBColor", "values": [0, 0, 0, 100]}}]}, "fontEffects": "Normal", "horizontalAlignment": "Left", "haloSize": 1, "lineGapType": "ExtraLeading", "hinting": "Default", "fontType": "TTOpenType", "textCase": "Normal", "verticalAlignment": "Top", "fontFamilyName": "Arial", "textDirection": "LTR", "billboardMode3D": "FaceNearPlane", "type": "CIMTextSymbol", "height": 10, "letterWidth": 100, "depth3D": 1, "ligatures": true}}, "blendingMode": "Alpha", "type": "CIMTextGraphic", "placement": "Unspecified"}, "anchor": "TopLeftCorner", "rotationCenter": {"x": 0.6528655328805799, "y": 9.133776437153976}}, {"visible": true, "lockedAspectRatio": true, "name": "FieldName", "type": "CIMGraphicElement", "graphic": {"text": "Field Name", "shape": {"x": 0.6528655328805799, "y": 9.343520391634774}, "symbol": {"type": "CIMSymbolReference", "symbol": {"blockProgression": "TTB", "fontEncoding": "Unicode", "wordSpacing": 100, "kerning": true, "extrapolateBaselines": true, "symbol3DProperties": {"rotationOrder3D": "XYZ", "type": "CIM3DSymbolProperties", "scaleZ": 0, "scaleY": 0, "dominantSizeAxis3D": "Z"}, "fontStyleName": "Bold", "verticalGlyphOrientation": "Right", "symbol": {"type": "CIMPolygonSymbol", "symbolLayers": [{"enable": true, "type": "CIMSolidFill", "color": {"type": "CIMRGBColor", "values": [0, 0, 0, 100]}}]}, "fontEffects": "Normal", "horizontalAlignment": "Left", "haloSize": 1, "lineGapType": "ExtraLeading", "hinting": "Default", "fontType": "TTOpenType", "textCase": "Normal", "verticalAlignment": "Top", "fontFamilyName": "Arial", "textDirection": "LTR", "billboardMode3D": "FaceNearPlane", "type": "CIMTextSymbol", "height": 10, "letterWidth": 100, "depth3D": 1, "ligatures": true}}, "blendingMode": "Alpha", "type": "CIMTextGraphic", "placement": "Unspecified"}, "anchor": "TopLeftCorner", "rotationCenter": {"x": 0.6528655328805799, "y": 9.343520391634772}}, {"visible": true, "graphic": {"blendingMode": "Alpha", "type": "CIMPolygonGraphic", "symbol": {"type": "CIMSymbolReference", "symbol": {"type": "CIMPolygonSymbol", "symbolLayers": [{"enable": true, "lineStyle3D": "Strip", "capStyle": "Round", "color": {"type": "CIMRGBColor", "values": [0, 0, 0, 0]}, "width": 1, "type": "CIMSolidStroke", "joinStyle": "Round", "miterLimit": 10}, {"enable": true, "type": "CIMSolidFill", "color": {"type": "CIMRGBColor", "values": [11, 121.99610900878906, 192, 100]}}]}}, "placement": "Unspecified", "polygon": {"hasZ": true, "rings": [[[0.5000000000000062, 10.5, null], [7.99232008788524, 10.5, null], [7.99232008788524, 10.061432466681081, null], [0.5000000000000062, 10.061432466681081, null], [0.5000000000000062, 10.5, null]]]}}, "name": "ReportHeaderBackground", "type": "CIMGraphicElement", "anchor": "BottomLeftCorner", "rotationCenter": {"x": 0.5000000000000062, "y": 10.061432466681081}}, {"visible": true, "lockedAspectRatio": true, "name": "Logo", "type": "CIMGraphicElement", "graphic": {"frame": {"backgroundSymbol": {"type": "CIMSymbolReference", "symbol": {"effects": [{"method": "Square", "type": "CIMGeometricEffectOffset", "count": 1, "offset": 0, "option": "Fast"}], "type": "CIMPolygonSymbol", "symbolLayers": [{"enable": true, "type": "CIMSolidFill", "color": {"type": "CIMRGBColor", "values": [255, 255, 255, 0]}}]}}, "type": "CIMGraphicFrame", "borderSymbol": {"type": "CIMSymbolReference", "symbol": {"effects": [{"method": "Square", "type": "CIMGeometricEffectOffset", "count": 1, "offset": 0, "option": "Fast"}], "type": "CIMLineSymbol", "symbolLayers": [{"enable": true, "lineStyle3D": "Strip", "capStyle": "Round", "color": {"type": "CIMRGBColor", "values": [255, 255, 255, 0]}, "width": 1, "type": "CIMSolidStroke", "joinStyle": "Round", "miterLimit": 10}]}}}, "sourceURL": "C:\\\\data\\\\State\\\\env impact\\\\logo_white.png", "pictureURL": "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAJAAAACQCAYAAADnRuK4AAAACXBIWXMAAAsTAAALEwEAmpwYAAAAIGNIUk0AAHolAACAgwAA+f8AAIDoAABSCAABFVgAADqXAAAXb9daH5AAAAVnSURBVHja7J3bkuMwCESDSv//y+zjzlbNOpYtQTc0z5OJgEODFF/M3T8y2VObTfyMqBITQAJm93caSZysO0BOsC4Dj591A2gXNBa0HjSYfAdEsxE0p5NmC2tmUabvTpPswhwQmBO+GNsa0AHyItCs+mfJMb39/VPg5HaAi3mEwi80BfIG0KzGwJDjOwUOvCrtUqMjSoGgQC5wjiuSv4QZUoEEznNFsgB4oIdoFzivQLrb1o63l5EAjuDZA9I3QDyi0EcwPBCyWwQiu4hVWPwGIDyraiU1+jdWHpm7AQrPrs93gPCbGh21CQzOzu+lOdlls0ECj9/8G9/wN4zmWf9/EDnjF0A4WMArwXPdPzefRN/doqMm0ATOWnxGAjzIiXLBkzMD+Q1wjKTaXfDkD9HsLcNB1+RocRkBwbZgIG3Td7qAPq9AaPB8a5uMiYNuqSMZnohbc1gPDynOrEYiPJFzlH20RT+ytpmUUJSh3QF9oTrknJsdjL5S7u7NfHbxeUesbBYbBZ30xXnCBE8MQLvkPnM4dEDYGeHxXdv4J/AgAqFZ57ACvZ17TlYadQWz2ywSrLfXW/92VZ8JnD0APVUf1mAZcSFFq6hNBehvMNSy9ivQHfXRnRP44Bw77xovFiRrqjp3AZKyCJ4jCiT14YLHMgBifF6PfpJI8HcSJulqTZaU1LbtfmwKhgFX/c7LXAXPpl0Yq+0GyQn8PWoTaTEvE2nBvmiX+osCKSh14LEMgLR1r6k8ng2QTPYVoLeUdlApFvWxDIB2LCbr9hkTPLW28Wz3YVWCxzIAcmKQTPDUUyDt4pqo5yBPvNQnOV/M23ipGoBFvSvDD8Bz+jXbUp+bAFXYYRg48GV3jzOAaGetLqlPjRnIqlWttvEalBliFFIIkxgewdVAgdjagXZeRVqY1IWkEHQ9kNQHGiApSXxcHQEgkxrIUFoYgwpVud7HKwKkVlY4jjPY+ehXM3YC16sD9DOhUU9rrfSyXUgfRmIwvl3u6hsrz8nnH9sQgyPrGuSV5Yf+Vq0LXIGy5JkRIuj2O4ACvfpO1Q5nRfAxYP0pw5M+K3iSd2Focq331zdUIC8MDtP6jREgzT16b7zgqQIPE0CCRzOQ7CU8jrje8Um+ql9GBw+lApngwbRJEGRvDI9mIMHzCh5HX/sArobuMxiF/0NJ1NyjbbzmnrT1CyDNPdsUSOdBgmfZBymQhmbNQJp7cABSG1PrWvJjKLGCRy1MllYIY+EDamNSHymQ4Nnry+rzgaRCMhoFYtz+tlKfbwBJhdSG6WcgKxj0UgX49H1hHpwYAwe00mmz7QQIzTH2RHkleO4ChDYLGVgiWz//cZBWFGPSyqnPCkBWxWGBvNePseFLXAmkVp9XMRwNAnQqgHry6wOA9MDwptv1nQqkE+q14HtVeHa2sKxg6WeV5NY7Diyi+qN0mdVn+wgykBYTrD7d1OtIvsahRTk4PGpbQDNQNEQs505eHZ6dQ3QERH44KRWPKI77NAIW6+DgVG1ZIQURdTmHB3/uRMUaETxhNgIX72Dw2MFEZL2sJhzyeTAx/p/AGniFVr6YnwKgn075RXVahwBX92sGOHj1PlMDDqgDAwiznhnk7CmIul0dAOcvwnvjXfBgDshIAN2BaOWd7yZw+rSwuzu0uy0NHR4/WGitFeiJGnW5cN4+pPe9TYDE+EIlt70DVAA926WxtqE2M9sEq8ZKP5q2GPQHYNBXd2uCQAr0aj5SkqVAjxUpQ5VkBAp01UoYVMkEED5MvrCLUpsTQF+rfAWo01C1AdjcS44PDgi4FIhYnU6D1bZFzmb+rjxJVXNTgW181/YngGRqYV1NrWvB/gwAupFXbGGGEi4AAAAASUVORK5CYII=", "blendingMode": "Alpha", "type": "CIMPictureGraphic", "box": {"xmin": 0.5237212021621606, "zmax": 0, "ymin": 10.078202852637217, "ymax": 10.483229614043866, "zmin": 0, "xmax": 0.9751151733600724}, "placement": "BottomLeftCorner"}, "anchor": "BottomLeftCorner", "rotationCenter": {"x": 0.5237212021621606, "y": 10.078202852637217}}, {"visible": true, "lockedAspectRatio": true, "name": "ReportTitle", "type": "CIMGraphicElement", "graphic": {"text": "Report Title", "shape": {"x": 1.084796287761, "y": 10.431595139590543}, "symbol": {"type": "CIMSymbolReference", "symbol": {"blockProgression": "TTB", "fontEncoding": "Unicode", "wordSpacing": 100, "kerning": true, "extrapolateBaselines": true, "symbol3DProperties": {"rotationOrder3D": "XYZ", "type": "CIM3DSymbolProperties", "scaleZ": 0, "scaleY": 0, "dominantSizeAxis3D": "Z"}, "fontStyleName": "Bold", "verticalGlyphOrientation": "Right", "symbol": {"type": "CIMPolygonSymbol", "symbolLayers": [{"enable": true, "type": "CIMSolidFill", "color": {"type": "CIMRGBColor", "values": [255, 255, 255, 255]}}]}, "fontEffects": "Normal", "horizontalAlignment": "Left", "haloSize": 1, "lineGapType": "ExtraLeading", "hinting": "Default", "fontType": "Unspecified", "textCase": "Normal", "verticalAlignment": "Top", "fontFamilyName": "Tahoma", "textDirection": "LTR", "billboardMode3D": "FaceNearPlane", "type": "CIMTextSymbol", "height": 18, "letterWidth": 100, "depth3D": 1, "ligatures": true}}, "blendingMode": "Alpha", "type": "CIMTextGraphic", "placement": "Unspecified"}, "anchor": "TopLeftCorner", "rotationCenter": {"x": 1.084796287761, "y": 10.431595139590542}}], "datePrinted": {"type": "TimeInstant"}, "sourceModifiedTime": {"type": "TimeInstant"}}, "version": "1.2.0"}' 
            pagx = self.write_pagx(overflow_layout_json, self.overflow_layout_name)
        else:
            base_name = os.path.splitext(os.path.basename(template))[0]
            self.overflow_layout_name = base_name + str(self.time_stamp)
            pagx = self.update_template(template, self.overflow_layout_name, self.overflow_element_names)
        self.overflow_pagx = pagx
        self.temp_files.append(self.overflow_pagx)
        self.aprx.importDocument(self.overflow_pagx)

    def update_template(self, template, name, elements):
        pagx = None
        with open(template) as data_file:    
            data = json.load(data_file)
            data['layoutDefinition']['name'] = name
            template_elements = data['layoutDefinition']['elements']
            template_element_names = [elm['name'] for elm in template_elements]
            missing_elements = []
            for elm in elements:
                if not elm in template_element_names:
                    if not elm in self.optional_element_names:
                        missing_elements.append(elm)
                else:
                    i = template_element_names.index(elm)
                    template_element = template_elements[i]
                    if template_element['anchor'] != 'TopLeftCorner':
                        data['layoutDefinition']['elements'][i]['anchor'] = 'TopLeftCorner'
            if len(missing_elements) > 0:
                arcpy.AddError("Missing required elements in: " + template)
                for elm in missing_elements:
                    arcpy.AddError("Cannot locate element: " + elm)
                sys.exit()
            s = json.dumps(data)
            pagx = self.temp_dir + os.sep + name + '.pagx'
            with open(pagx, 'w') as write_file:
                write_file.writelines(s)
        return pagx

    def find_elements(self, type):
        layout_name = self.map_layout_name if type == 'map' else self.overflow_layout_name
        element_names = self.map_element_names if type == 'map' else self.overflow_element_names
        layouts = self.aprx.listLayouts(layout_name)
        if len(layouts) > 0:
            out_elements = {}
            for elm_name in element_names:
                elms = layouts[0].listElements(wildcard = elm_name)
                if len(elms) > 0:
                    elm = elms[0]
                    out_elements[elm.name] = elm
            return out_elements
        else:
            arcpy.AddError("Cannot locate Layout: " + layout_name)
            sys.exit()

    def update_layouts(self):
        self.set_layout('map', self.map_pagx)
        x = 0
        t=None
        first_overflow = True
        for table in self.tables:
            if table.is_overflow:
                self.delete_elements()
                if first_overflow:
                    first_overflow = False
                    os.remove(self.map_pagx)
                else:
                    self.init_layouts(True)
                self.set_layout('overflow', self.overflow_pagx)
                self.temp_files.remove(self.overflow_pagx)
                os.remove(self.overflow_pagx)
                if table.full_overflow:
                    table.is_overflow = False
            overflow = table.init_table(self.cur_elements, self.remaining_height, self.layout_type)
            x += 1
            if overflow:
                overflow_table = Table(table.title, table.overflow_rows, table.fields)
                overflow_table.is_overflow = True
                overflow_table.full_overflow = table.full_overflow
                overflow_table.field_widths = table.field_widths
                overflow_table.auto_adjust = table.auto_adjust
                overflow_table.row_width = table.row_width
                overflow_table.max_vals = table.max_vals
                overflow_table.total_row_index = table.total_row_index
                overflow_table.has_buffer_rows = table.has_buffer_rows
                overflow_table.buffer_rows = table.buffer_rows
                overflow_table.adjust_header_columns = table.adjust_header_columns
                overflow_table.field_name_lengths = table.field_name_lengths
                if hasattr(table, 'p_fields'):
                    if len(table.p_fields) > 0:
                        overflow_table.p_fields = table.p_fields
                table.total_row_index = None
                self.tables.insert(x, overflow_table)
            else:
                if table.has_buffer_rows:
                    fields = table.fields
                    if hasattr(table, 'p_fields'):
                        if len(table.p_fields) > 0:
                            fields = table.p_fields
                    buffer_rows_table = Table(table.title + BUFFER_TITLE, table.buffer_rows, fields)
                    buffer_rows_table.adjust_header_columns = table.adjust_header_columns
                    self.tables.insert(x, buffer_rows_table)

            if table.row_count > 0:
                table_header_background = table.table_header_background
                if self.base_y == None:
                    if table.is_overflow:
                        height = self.place_holder.elementHeight
                    else:
                        height = self.place_holder.elementHeight - table_header_background.elementHeight
                else:
                    height = self.remaining_height - table_header_background.elementHeight

                if not table.is_overflow:
                    table_header_background = table.table_header_background.clone('table_header_background_clone')
                    table_title = table.table_title.clone('table_title_clone')
                    table_title.text = table.title
                    f = (table_header_background.elementHeight - table_title.elementHeight) / 2
                    if self.base_y == None:
                        table_header_background.elementPositionX = self.base_x
                        table_header_background.elementPositionY = self.cur_y
      
                        table_title.elementPositionX = self.base_x
                        table_title.elementPositionY = self.cur_y - f
                    else:                  
                        table_header_background.elementPositionX = self.base_x
                        table_header_background.elementPositionY = self.base_y

                        table_title.elementPositionX = self.base_x
                        table_title.elementPositionY = self.base_y - f
                        self.cur_y = self.base_y
                        self.cur_x = self.base_x

                    self.cur_y -= (table_header_background.elementHeight + Y_MARGIN)
                start_y = self.cur_y 

                arcpy.AddMessage("Generating Table: " + table.title)  
                if table.title == ANALYSIS_SUMMARY:  
                    self.add_row_backgrounds(table)
                    self.add_table_lines('vertLineA', True, table)
                    self.add_table_lines('horzLineA', False, table)
                else:
                    self.add_row_backgrounds(table)
                    self.add_table_lines('vertLine', True, table)
                    self.add_table_lines('horzLine', False, table)

                self.base_y = self.cur_y - Y_MARGIN
                eh = self.place_holder.elementHeight
                esy = self.place_holder.elementPositionY
                self.remaining_height = eh - (esy - self.base_y)

                #first reset the x/y
                self.cur_x = self.place_holder.elementPositionX
                self.cur_y = start_y
                self.add_values(table)
        self.delete_elements()

    def drop_add_layers(self, map_frame):
        print('No Legend')
        #l = arcpy.mp.LayerFile(r'C:\Users\john4818\Documents\ArcGIS\Projects\EnvImpact1.2\ArtificialReefSite.lyrx')

        #ll = self.aprx.listLayouts(self.map_layout_name)

        #le = ll[0].listElements('LEGEND_ELEMENT')
        #le[0].mapFrame.map.addLayer(l)
        #le[0].title = "AAAAA"
        #self.aprx.save()

        ##TODO this is not working yet
        #legend = self.cur_elements['Legend']
        #legend.title = "AAAA"
        #layers = legend.mapFrame.map.listLayers()
        #ly=[]
        #for l in layers:
        #    if l.isFeatureLayer:
        #        l_path = l.saveACopy(self.temp_dir + os.sep + l.name)
        #        ly.append(l_path)
        #        self.temp_files.append(l_path)
        #        legend.mapFrame.map.removeLayer(l)
        #for l in ly:
        #    lf = arcpy.mp.LayerFile(str(l))
        #    legend.mapFrame.map.addLayer(lf)

    def set_layout(self, layout_type, pagx):
        self.page_num += 1
        self.layout_type = layout_type
        layout_name = self.map_layout_name if layout_type == 'map' else self.overflow_layout_name
        layouts = self.aprx.listLayouts(layout_name)
        if len(layouts) > 0:
            self.cur_elements = self.elements[layout_name]
            self.place_holder = self.cur_elements['ContentDisplayArea']
            self.cur_x = self.place_holder.elementPositionX
            self.cur_y = self.place_holder.elementPositionY
            self.base_x = self.cur_x
            self.base_y = None
            self.remaining_height = None
            self.layout = layouts[0]
            arcpy.AddMessage("Layout Set: " + layout_name)
            self.add_pdf()
            self.set_element_props()
        else:
            arcpy.AddError("Cannot find layout:" + layout_name)
            sys.exit()
            self.remaining_height = None

    def set_element_props(self):
        #set map specific element props
        if self.layout_type == 'map':
            self.cur_elements['ReportSubTitle'].text = self.sub_title
            self.cur_elements['ReportType'].text = self.report_type
            if not self.map in ['', ' ', None]:
                maps = self.aprx.listMaps(self.map)
                map_frame = self.cur_elements['MapFrame']
                if len(maps) > 0:
                    user_map = maps[0]
                    ext = user_map.defaultCamera.getExtent()
                    map_frame.map = user_map
                    self.drop_add_layers(map_frame)
                    map_frame.camera.setExtent(ext)            
            if self.scale_unit == 'Metric Units':
                self.cur_elements['ScaleBarM'].visible = False
                self.cur_elements['ScaleBarKM'].visible = True
            else:
                self.cur_elements['ScaleBarKM'].visible = False
                self.cur_elements['ScaleBarM'].visible = True

        #set common element props   
        self.cur_elements['ReportTitle'].text = self.report_title
        if not self.logo in ['', ' ', None]:
            if 'Logo' in self.cur_elements:
                self.cur_elements['Logo'].sourceImage = self.logo
        self.cur_elements['PageNumber'].text = self.page_num
        self.cur_elements['ReportTitleFooter'].text = self.report_title

    def add_row_backgrounds(self, table):
        x = 0
        cur_x = self.cur_x
        cur_y = self.cur_y
        for row_height in table.row_heights:
            if x % 2 == 0:
                temp_row_background = table.row_background.clone("row_background_clone")
                temp_row_background.elementWidth = table.row_width
                temp_row_background.elementHeight = row_height
                temp_row_background.elementPositionX = cur_x
                temp_row_background.elementPositionY = cur_y             
            cur_y -= row_height
            x += 1 

    def add_table_lines(self, element_name, vert, table): 
        line = self.cur_elements[element_name].clone(element_name + "_clone")
        if vert:
            full_vert = []
            line.elementHeight = table.table_height
            collection = table.field_widths
            if not table.total_row_index == None:
                f = table.total_row_index
                while f < len(collection):
                    f += 1
                    full_vert.append(f)                
        else:
            line.elementWidth = table.row_width
            collection = table.row_heights
        line.elementPositionX = self.cur_x
        line.elementPositionY = self.cur_y
        cur_pos = line.elementPositionX if vert else line.elementPositionY
        x = 0      
        for adjust_value in collection:
            line_clone = line.clone(element_name + "_clone")   
            if vert:
                if len(full_vert) > 0 and not x + 1 in full_vert:
                    total_row_height = table.row_heights[len(table.row_heights) -1]
                    line_clone.elementHeight -= total_row_height
                cur_pos += adjust_value
                line_clone.elementPositionX = cur_pos
            else:
                cur_pos -= adjust_value
                line_clone.elementPositionY = cur_pos
                self.cur_y = cur_pos
            x += 1

    def add_values(self, table):
        field_value = table.field_value
        field_name = table.field_name
        base_x = self.cur_x
        x = 0
        date_fields = []
        percent_fields = []               
        for f in table.fields:
            if f.type == 'Date':
                date_fields.append(x)
            if f.name == PERCENT_FIELD:
                percent_fields.append(x)
            if not table.is_overflow:
                elm = field_name.clone("header_clone")
                v = f.aliasName
                if x in table.adjust_header_columns:
                    v = table.adjust_header_columns[x]
                elm.text = v
                elm.elementPositionX = self.cur_x + Y_MARGIN
                elm.elementPositionY = self.cur_y - Y_MARGIN
                self.cur_x += table.field_widths[x]
            x += 1
        first_row = True
        new_row = True
        x = 0
        xx = 0
        for row in table.rows:
            for v in row:
                if new_row:
                    x = 0
                    self.cur_x = base_x
                    if not table.is_overflow:
                        self.cur_y -= float(table.row_heights[xx])
                    new_row = False
                elm = field_value.clone("cell_clone")
                if x in date_fields:
                    elm.text = v
                elif x in percent_fields:
                    elm.text = str(v) + '%'
                else:
                    elm.text = v if not is_float(v) else NUM_DIGITS.format(float(v)) 
                new_x = self.cur_x + Y_MARGIN
                if not table.total_row_index == None:
                    if xx == len(table.rows) -1:
                        if x == table.total_row_index:
                            w = elm.elementWidth + X_MARGIN
                            new_x = (self.cur_x + table.field_widths[x]) - w
                elm.elementPositionX = new_x
                elm.elementPositionY = self.cur_y - Y_MARGIN
                self.cur_x += table.field_widths[x]
                x += 1
            new_row = True
            if table.is_overflow:
                self.cur_y -= float(table.row_heights[xx])
            xx += 1

    def delete_elements(self):
        for elm in self.cur_elements:
            if elm not in ['ReportTitle', 'ReportSubTitle', 'ReportType', 'Logo', 'PageNumber', 'ReportTitleFooter', 'Legend']:
                if elm in self.cur_elements:
                    if hasattr(self.cur_elements[elm], 'delete'):
                        self.cur_elements[elm].delete()

    def add_pdf(self):
        self.pdfs.append({'title': self.report_title, 'layout': self.layout})

    def export_pdf(self):
        arcpy.AddMessage("Exporting Report...")
        x = 0
        for pdf in self.pdfs:
            unique = str(x) if len(self.pdfs) > 1 else ''
            unique += self.time_stamp
            pdf_path = os.path.join(self.path, pdf['title'] + unique + ".pdf")
            pdf['layout'].exportToPDF(pdf_path)
            self.pdf_paths.append(pdf_path)
            x += 1     
        self.append_pages()
        arcpy.AddMessage("Report exported sucessfully.")

    def append_pages(self):
        if os.path.isfile(self.pdf_path) and os.access(self.pdf_path, os.R_OK):
            try:
                os.remove(self.pdf_path)
            except OSError:
                arcpy.AddWarning("Unable to export report: " + self.pdf_path)
                os.path.basename(self.pdf_path)
                base_name = os.path.splitext(os.path.basename(self.pdf_path))[0] + self.time_stamp
                base_path = os.path.dirname(self.pdf_path)
                self.pdf_path = "{0}\\{1}.pdf".format(base_path, base_name)
        try:
            arcpy.AddMessage(self.pdf_path)
            pdf_doc = arcpy.mp.PDFDocumentCreate(self.pdf_path)
        except OSError:
            arcpy.AddMessage("Unable to export report: " + self.pdf_path)
            arcpy.AddMessage("Please ensure you have write access to: " + base_path)
            if len(self.pdf_paths) > 0:
                for pdf in self.pdf_paths:
                    os.remove(pdf)
                sys.exit()
        for pdf in self.pdf_paths:
            pdf_doc.appendPages(pdf)
            os.remove(pdf)
        pdf_doc.saveAndClose()

    def generate_report(self, folder):
        if folder in ['', ' ', None] or not os.path.isdir(folder):
            self.path = self.aprx.homeFolder
        else:
            self.path = folder
        self.pdf_path = os.path.join(self.path, self.report_title + ".pdf")
        self.update_layouts()
        self.export_pdf()
        return self.pdf_path

    def write_pagx(self, json, name):
        base_path = self.temp_dir
        pagx = base_path + os.sep + 'temp_' + name + '.pagx'
        with open(pagx, 'w') as write_file:
            write_file.writelines(json)
        return pagx

def is_float(value):
  try:
    if str(value).count('.') == 0:
        return False
    float(value)
    return True
  except ValueError:
    return False
    
def main():
    arcpy.env.overwriteOutput = True

    report_title = arcpy.GetParameterAsText(0)             #required parameter for report title
    sub_title = arcpy.GetParameterAsText(1)                #optional parameter for sub-title
    logo = arcpy.GetParameterAsText(2)                     #optional report logo image
    tables = arcpy.GetParameterAsText(3)                   #required multivalue parameter for input tables
    map = arcpy.GetParameterAsText(4)                      #required parameter for the map 
    scale_unit = arcpy.GetParameterAsText(5)               #required scale unit with default set
    report_type = arcpy.GetParameterAsText(6)              #optional report type eg. ...
    map_report_template = arcpy.GetParameterAsText(7)      #optional parameter for path to new pagX files
    overflow_report_template = arcpy.GetParameterAsText(8) #optional parameter for path to new pagX files
    out_folder = arcpy.GetParameterAsText(9)               #folder that will contain the final output report

    report = None
    try:
        tables = [t.strip("'") for t in tables.split(';')]

        report = Report(report_title, sub_title, logo, map, scale_unit, report_type, map_report_template, overflow_report_template)
   
        for table in tables:
            table_title = os.path.splitext(os.path.basename(table))[0]       
            desc = arcpy.Describe(table)
            if desc.dataType == "FeatureLayer" or desc.dataType == "TableView":
                fi = desc.fieldInfo
                fields = [f for f in desc.fields if f.type not in ['Geometry', 'OID'] and
                               fi.getVisible(fi.findFieldByName(f.name) == 'VISIBLE')]
            else:
                fields = [f for f in desc.fields if f.type not in ['Geometry', 'OID']]
            cur = arcpy.da.SearchCursor(table, [f.name for f in fields])
            test_rows = [[str(v).replace('\n','') for v in r] for r in cur]
            report.add_table(table_title, test_rows, fields)

        pdf = report.generate_report(out_folder)
        os.startfile(pdf)
    except Exception as ex:
        arcpy.AddError(ex.args)
    finally:
        if report != None:
            for pdf in report.pdf_paths:
                if os.path.exists(pdf):
                    os.remove(pdf)
            for file in report.temp_files:
                if os.path.exists(file):
                    os.remove(file)

if __name__ == '__main__':
    main()





