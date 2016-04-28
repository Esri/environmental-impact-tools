import arcpy
import os, sys, json, time, textwrap

#avoid showing urls and or url fields...hold off on this for now...

#special handling for RESULTTYPE 
#AOI vs BUFFER

#issue with table that overflows to more than 2 pages...empty row at t he bottom

#look at the wrapping issue with the cameo data

#should look at setting the line symbol cap to something other than round 
# to see if we can avoid the occasional dangle I am seeing

#TOTAL calc

#get legend to come to life when we set the map

SPLIT_FIELD = 'RESULTTYPE'
SPLIT_VAL_BUFFER = 'BUFFER'
SPLIT_VAL_AOI = 'AOI'

BUFFER_TITLE = ' (within buffer)'

MARGIN = .025

PROJECT_PATH = r"C:\Users\john4818\Documents\ArcGIS\Projects\EnvImpact1.2\EnvImpact1.aprx"
#PROJECT_PATH = "CURRENT"

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

    def calc_widths(self):
        self.auto_adjust = []
        field_name_lengths = []
        
        for f in self.fields:
            self.field_name.text = f
            w = self.field_name.elementWidth + (MARGIN * 2)
            field_name_lengths.append(w)
            self.field_widths.append(w)

        self.max_vals = self.get_max_vals(self.rows)

        x = 0
        for v in self.max_vals:
            length = field_name_lengths[x]
            self.field_value.text = v
            potential_length = self.field_value.elementWidth + (MARGIN * 2)
            if potential_length > length:
                self.field_widths[x] = potential_length
            x += 1

        self.row_width = sum(self.field_widths)

        if self.row_width > self.content_display.elementWidth:
            self.auto_adjust = self.adjust_row_widths(field_name_lengths)        

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

    def adjust_row_widths(self, field_name_lengths):
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
        for w in large_field_widths:
            if float(w)/float(sum_widths) > 0:
                self.field_widths[idx] = display_width * (1/len(pos_widths))
                auto_adjust.append(idx)
            idx += 1

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
        self.header_height = h_height + (MARGIN * 2)

        c_height = self.field_value.elementHeight 
        self.row_height = c_height + (MARGIN * 2)

        row_heights = []
        if len(self.auto_adjust) > 0:
            row_heights = [self.row_height] * len(self.rows)
            for column_index in self.auto_adjust:
                col_width = self.field_widths[column_index]
                fit_width = col_width - (MARGIN * 2)
                long_val = self.max_vals[column_index]
                max_chars = self.calc_num_chars(fit_width, long_val)
                x = 0
                for row in self.rows:
                    v = str(row[column_index])
                    if len(v) > max_chars:
                        wrapped_val = textwrap.wrap(v, max_chars)
                        wrapped_height = (len(wrapped_val) * (self.row_height - MARGIN))
                        if wrapped_height > row_heights[x]:
                            row_heights[x] = wrapped_height
                        row[column_index] = '\n'.join(wrapped_val)
                    x += 1
        else:
            table_height = self.row_count * self.row_height
            row_heights = [self.row_height] * len(self.rows)

        if not self.is_overflow:
            row_heights.insert(0, self.header_height)  
        table_height = sum(row_heights)
        self.row_heights = row_heights

        if self.remaining_height == None:
            self.remaining_height = self.content_display.elementHeight
        if not self.is_overflow:
            self.remaining_height -= (self.table_header_background.elementHeight + MARGIN)
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
        self.field_name = elements['FieldName']
        self.field_value = elements['FieldValue']
        self.table_header_background = elements['TableHeaderBackground']
        self.table_title = elements['TableTitle']
        self.row_background = elements['EvenRowBackground']
        self.content_display = elements['ContentDisplayArea']

        #Map only elements
        if layout_type == 'map':
            self.report_title = elements['ReportTitle']
    
    def check_result_type(self):
        if SPLIT_FIELD in self.fields:
            split_field_idx = self.fields.index(SPLIT_FIELD)
            buffer_rows = []
            aoi_rows = []
            for r in self.rows:
                v = r[split_field_idx]
                if v == SPLIT_VAL_BUFFER:
                    r.pop(split_field_idx)
                    buffer_rows.append(r)
                else:
                    #TODO do we want to default to this or handle differently
                    r.pop(split_field_idx)
                    aoi_rows.append(r)
            if len(buffer_rows) > 0:
                self.has_buffer_rows = True
                self.buffer_rows = buffer_rows
            self.rows = aoi_rows
            self.fields.remove(SPLIT_FIELD)

    def init_table(self, elements, remaining_height, layout_type):
        #locate placeholder elements
        self.init_elements(elements, layout_type)

        self.remaining_height = remaining_height

        self.check_result_type()

        #Calculate the column/row widths and the row/table heights
        if len(self.field_widths) == 0:
            self.calc_widths()
        overflow = self.calc_heights()
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
    
        #self.aprx = arcpy.mp.ArcGISProject(PROJECT_PATH)
        self.aprx = arcpy.mp.ArcGISProject('CURRENT')

        self.map_element_names =  ['horzLine', 'vertLine', 'FieldValue', 
                                    'FieldName', 'ReportType', 'TableHeaderBackground', 
                                    'TableTitle', 'EvenRowBackground', 'ContentDisplayArea',
                                    'ReportSubTitle', 'ReportTitle', 'Scale Bar',
                                    'Logo', 'PageNumber', 'ReportTitleFooter', 'Map Map Frame']

        self.overflow_element_names =  ['horzLine', 'vertLine', 'FieldValue', 
                                        'FieldName', 'TableHeaderBackground', 'TableTitle', 
                                        'EvenRowBackground', 'ContentDisplayArea', 'ReportTitle',
                                        'Logo', 'PageNumber', 'ReportTitleFooter']

        self.init_layouts(False)

    def update_time_stamp(self):
        self.time_stamp = time.strftime("%m%d%Y-%H%M%S")

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
            self.map_layout_name = 'MapLayout{0}'.format(self.time_stamp)
            map_layout_json = '{"version": "1.2.0", "binaryReferences": [{"type": "CIMBinaryReference", "data": "<?xml version=\\"1.0\\"?>\\r\\n<metadata xml:lang=\\"en\\"><Esri><CreaDate>20160415</CreaDate><CreaTime>10502300</CreaTime><ArcGISFormat>1.0</ArcGISFormat><SyncOnce>TRUE</SyncOnce></Esri><dataIdInfo><idCitation><resTitle>Layout</resTitle></idCitation></dataIdInfo></metadata>\\r\\n", "uRI": "CIMPATH=Metadata/2edbf119be763c7b7b29b3cc4aeaf083.xml"}, {"type": "CIMBinaryReference", "data": "<?xml version=\\"1.0\\"?>\\r\\n<metadata xml:lang=\\"en\\"><Esri><CreaDate>20160415</CreaDate><CreaTime>10501700</CreaTime><ArcGISFormat>1.0</ArcGISFormat><SyncOnce>TRUE</SyncOnce></Esri><dataIdInfo><idCitation><resTitle>Map</resTitle></idCitation></dataIdInfo><Binary><Thumbnail><Data EsriPropertyType=\\"PictureX\\">/9j/4AAQSkZJRgABAQEAAAAAAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0a\\r\\nHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/2wBDAQkJCQwLDBgNDRgyIRwhMjIyMjIy\\r\\nMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjL/wAARCACFAMgDAREA\\r\\nAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQA\\r\\nAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3\\r\\nODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWm\\r\\np6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEA\\r\\nAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSEx\\r\\nBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElK\\r\\nU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3\\r\\nuLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD3PSf+\\r\\nQNY/9e8f/oIoAoXSnz3x/eP867Kb91Hm1V7zK5BB5GK1TMxKYD41yfepbGi2qhFyaxbNoxGPiTeC\\r\\nSrbcoQR2z3/L9Kzk7M2gk0+4kEnlRpJu3ySDbGP8/wBfSocbuyNVLlV2XLOFFVnxlieT2PfIFEuw\\r\\nU9dS7UGghGRQBnpkF1wAisQAO/ua1iupzTfQrpiOUjsTXQ9Ucy0kKJcSYVD1wWYcVk5G8YdSdWyO\\r\\nVwDQhOxE6Ss3yvhSMYpiEt7cwIqlycDGO3/66QMsdV70dR9CAuQX3gBFPUnscdvrU3syrXiQSxFW\\r\\n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\\r\\n87kurf8AIGvv+veT/wBBNYnWf//Z</Data></Thumbnail></Binary></metadata>\\r\\n", "uRI": "CIMPATH=Metadata/ccaf3d769eb8676927a18c0b5bed382a.xml"}], "layoutDefinition": {"dateExported": {"type": "TimeInstant"}, "datePrinted": {"type": "TimeInstant", "start": 1461164297575}, "name": "' + self.map_layout_name + '", "elements": [{"graphic": {"placement": "Unspecified", "type": "CIMPolygonGraphic", "blendingMode": "Alpha", "symbol": {"type": "CIMSymbolReference", "symbol": {"type": "CIMPolygonSymbol", "symbolLayers": [{"color": {"type": "CIMRGBColor", "values": [0, 0, 0, 100]}, "capStyle": "Round", "miterLimit": 10, "type": "CIMSolidStroke", "width": 1, "lineStyle3D": "Strip", "joinStyle": "Round", "enable": true}]}}, "polygon": {"hasZ": true, "rings": [[[0.5000000000000018, 5.023148148148151, null], [8.000000000000005, 5.023148148148151, null], [8.000000000000005, 0.749999999999992, null], [0.5000000000000018, 0.749999999999992, null], [0.5000000000000018, 5.023148148148151, null]]]}}, "name": "ContentDisplayArea", "anchor": "TopLeftCorner", "type": "CIMGraphicElement", "rotationCenter": {"x": 0.5000000000000018, "y": 5.023148148148151}, "visible": true}, {"graphic": {"placement": "Unspecified", "type": "CIMLineGraphic", "blendingMode": "Alpha", "symbol": {"type": "CIMSymbolReference", "symbol": {"type": "CIMLineSymbol", "symbolLayers": [{"color": {"type": "CIMRGBColor", "values": [0, 0, 0, 100]}, "capStyle": "Round", "miterLimit": 10, "type": "CIMSolidStroke", "width": 1, "lineStyle3D": "Strip", "joinStyle": "Round", "enable": true}]}}, "line": {"paths": [[[0.5124186327888687, 4.67256647754833, null], [2.5124186327888687, 4.67256647754833, null]]], "hasZ": true}}, "name": "horzLine", "anchor": "TopLeftCorner", "type": "CIMGraphicElement", "rotationCenter": {"x": 0.5124186327888687, "y": 4.67256647754833}, "visible": true}, {"graphic": {"placement": "Unspecified", "type": "CIMLineGraphic", "blendingMode": "Alpha", "symbol": {"type": "CIMSymbolReference", "symbol": {"type": "CIMLineSymbol", "symbolLayers": [{"color": {"type": "CIMRGBColor", "values": [0, 0, 0, 100]}, "capStyle": "Round", "miterLimit": 10, "type": "CIMSolidStroke", "width": 1, "lineStyle3D": "Strip", "joinStyle": "Round", "enable": true}]}}, "line": {"paths": [[[2.51238689284104, 4.6725664775483295, null], [2.5124186327888687, 3.673055109399892, null]]], "hasZ": true}}, "name": "vertLine", "anchor": "TopLeftCorner", "type": "CIMGraphicElement", "rotationCenter": {"x": 2.51238689284104, "y": 4.6725664775483295}, "visible": true}, {"graphic": {"placement": "Unspecified", "symbol": {"type": "CIMSymbolReference", "symbol": {"fontFamilyName": "Arial", "haloSize": 1, "symbol": {"type": "CIMPolygonSymbol", "symbolLayers": [{"type": "CIMSolidFill", "enable": true, "color": {"type": "CIMRGBColor", "values": [0, 0, 0, 100]}}]}, "textCase": "Normal", "height": 10, "verticalAlignment": "Top", "symbol3DProperties": {"scaleZ": 0, "type": "CIM3DSymbolProperties", "rotationOrder3D": "XYZ", "dominantSizeAxis3D": "Z", "scaleY": 0}, "ligatures": true, "fontType": "TTOpenType", "fontEncoding": "Unicode", "fontStyleName": "Bold", "lineGapType": "ExtraLeading", "letterWidth": 100, "verticalGlyphOrientation": "Right", "type": "CIMTextSymbol", "horizontalAlignment": "Left", "fontEffects": "Normal", "extrapolateBaselines": true, "blockProgression": "TTB", "textDirection": "LTR", "billboardMode3D": "FaceNearPlane", "depth3D": 1, "hinting": "Default", "wordSpacing": 100, "kerning": true}}, "shape": {"x": 0.5934970669224624, "y": 4.6238224202821225}, "text": "Field Name", "type": "CIMTextGraphic", "blendingMode": "Alpha"}, "lockedAspectRatio": true, "name": "FieldName", "anchor": "TopLeftCorner", "type": "CIMGraphicElement", "rotationCenter": {"x": 0.5934970669224624, "y": 4.6238224202821225}, "visible": true}, {"graphic": {"placement": "Unspecified", "symbol": {"type": "CIMSymbolReference", "symbol": {"fontFamilyName": "Arial", "haloSize": 1, "symbol": {"type": "CIMPolygonSymbol", "symbolLayers": [{"type": "CIMSolidFill", "enable": true, "color": {"type": "CIMRGBColor", "values": [0, 0, 0, 100]}}]}, "textCase": "Normal", "height": 10, "verticalAlignment": "Top", "symbol3DProperties": {"scaleZ": 0, "type": "CIM3DSymbolProperties", "rotationOrder3D": "XYZ", "dominantSizeAxis3D": "Z", "scaleY": 0}, "ligatures": true, "fontType": "TTOpenType", "fontEncoding": "Unicode", "fontStyleName": "Regular", "lineGapType": "ExtraLeading", "letterWidth": 100, "verticalGlyphOrientation": "Right", "type": "CIMTextSymbol", "horizontalAlignment": "Left", "fontEffects": "Normal", "extrapolateBaselines": true, "blockProgression": "TTB", "textDirection": "LTR", "billboardMode3D": "FaceNearPlane", "depth3D": 1, "hinting": "Default", "wordSpacing": 100, "kerning": true}}, "shape": {"x": 0.5934970669224624, "y": 4.414078465801327}, "text": "Field Value", "type": "CIMTextGraphic", "blendingMode": "Alpha"}, "lockedAspectRatio": true, "name": "FieldValue", "anchor": "TopLeftCorner", "type": "CIMGraphicElement", "rotationCenter": {"x": 0.5934970669224624, "y": 4.414078465801326}, "visible": true}, {"frame": {"rings": [[[0.5184697855750491, 5.222222222222216], [0.5184697855750491, 9.281250000000004], [5.537037037037043, 9.281250000000004], [5.537037037037043, 5.222222222222216], [0.5184697855750491, 5.222222222222216]]]}, "graphicFrame": {"type": "CIMGraphicFrame", "backgroundSymbol": {"type": "CIMSymbolReference", "symbol": {"effects": [{"offset": 0, "type": "CIMGeometricEffectOffset", "count": 1, "option": "Fast", "method": "Square"}], "type": "CIMPolygonSymbol", "symbolLayers": [{"type": "CIMSolidFill", "enable": true, "color": {"type": "CIMRGBColor", "values": [255, 255, 255, 0]}}]}}, "borderSymbol": {"type": "CIMSymbolReference", "symbol": {"effects": [{"offset": 0, "type": "CIMGeometricEffectOffset", "count": 1, "option": "Fast", "method": "Square"}], "type": "CIMLineSymbol", "symbolLayers": [{"color": {"type": "CIMRGBColor", "values": [0, 0, 0, 100]}, "capStyle": "Round", "miterLimit": 10, "type": "CIMSolidStroke", "width": 1, "lineStyle3D": "Strip", "joinStyle": "Round", "enable": true}]}}}, "uRI": "CIMPATH=map/map3.xml", "name": "Map Map Frame", "autoCamera": {"autoCameraType": "Extent", "marginType": "Percent", "type": "CIMAutoCamera", "source": "None"}, "anchor": "BottomLeftCorner", "type": "CIMMapFrame", "rotationCenter": {"x": 0.5184697855750491, "y": 5.222222222222216}, "visible": true, "view": {"timeDisplay": {"defaultTimeIntervalUnits": "esriTimeUnitsUnknown", "type": "CIMMapTimeDisplay", "timeRelation": "esriTimeRelationOverlaps"}, "camera": {"x": -8546445.768560957, "z": "NaN", "y": 4744531.008165626, "type": "CIMViewCamera", "scale": 636915.4936921074, "pitch": -90}, "verticalExaggerationScaleFactor": 1, "viewableObjectPath": "CIMPATH=map/map.xml", "type": "CIMMapView", "viewingMode": "Map"}}, {"graphic": {"placement": "Unspecified", "type": "CIMPolygonGraphic", "blendingMode": "Alpha", "symbol": {"type": "CIMSymbolReference", "symbol": {"type": "CIMPolygonSymbol", "symbolLayers": [{"color": {"type": "CIMRGBColor", "values": [0, 0, 0, 0]}, "capStyle": "Round", "miterLimit": 10, "type": "CIMSolidStroke", "width": 1, "lineStyle3D": "Strip", "joinStyle": "Round", "enable": true}, {"type": "CIMSolidFill", "enable": true, "color": {"type": "CIMRGBColor", "values": [11, 121.99610900878906, 192, 100]}}]}}, "polygon": {"hasZ": true, "rings": [[[1.2693865500071997, 10.5, null], [8.000000000000002, 10.5, null], [8.000000000000002, 9.792000000000002, null], [1.2693865500071997, 9.792000000000002, null], [1.2693865500071997, 10.5, null]]]}}, "name": "ReportHeaderBackground", "anchor": "BottomLeftCorner", "type": "CIMGraphicElement", "rotationCenter": {"x": 1.2693865500071997, "y": 9.792000000000002}, "visible": true}, {"graphic": {"placement": "Unspecified", "symbol": {"type": "CIMSymbolReference", "symbol": {"fontFamilyName": "Tahoma", "haloSize": 1, "symbol": {"type": "CIMPolygonSymbol", "symbolLayers": [{"type": "CIMSolidFill", "enable": true, "color": {"type": "CIMRGBColor", "values": [255, 255, 255, 255]}}]}, "textCase": "Normal", "height": 14, "verticalAlignment": "Bottom", "symbol3DProperties": {"scaleZ": 0, "type": "CIM3DSymbolProperties", "rotationOrder3D": "XYZ", "dominantSizeAxis3D": "Z", "scaleY": 0}, "ligatures": true, "fontType": "Unspecified", "fontEncoding": "Unicode", "fontStyleName": "Regular", "lineGapType": "ExtraLeading", "letterWidth": 100, "verticalGlyphOrientation": "Right", "type": "CIMTextSymbol", "horizontalAlignment": "Left", "fontEffects": "Normal", "extrapolateBaselines": true, "blockProgression": "TTB", "textDirection": "LTR", "billboardMode3D": "FaceNearPlane", "depth3D": 1, "hinting": "Default", "wordSpacing": 100, "kerning": true}}, "shape": {"x": 1.405532031711079, "y": 10.193474522079494}, "text": "Report Type", "type": "CIMTextGraphic", "blendingMode": "Alpha"}, "lockedAspectRatio": true, "name": "ReportType", "anchor": "BottomLeftCorner", "type": "CIMGraphicElement", "rotationCenter": {"x": 1.405532031711079, "y": 10.193474522079494}, "visible": true}, {"graphic": {"placement": "Unspecified", "symbol": {"type": "CIMSymbolReference", "symbol": {"fontFamilyName": "Tahoma", "haloSize": 1, "symbol": {"type": "CIMPolygonSymbol", "symbolLayers": [{"type": "CIMSolidFill", "enable": true, "color": {"type": "CIMRGBColor", "values": [255, 255, 255, 255]}}]}, "textCase": "Normal", "height": 18, "verticalAlignment": "Top", "symbol3DProperties": {"scaleZ": 0, "type": "CIM3DSymbolProperties", "rotationOrder3D": "XYZ", "dominantSizeAxis3D": "Z", "scaleY": 0}, "ligatures": true, "fontType": "Unspecified", "fontEncoding": "Unicode", "fontStyleName": "Bold", "lineGapType": "ExtraLeading", "letterWidth": 100, "verticalGlyphOrientation": "Right", "type": "CIMTextSymbol", "horizontalAlignment": "Left", "fontEffects": "Normal", "extrapolateBaselines": true, "blockProgression": "TTB", "textDirection": "LTR", "billboardMode3D": "FaceNearPlane", "depth3D": 1, "hinting": "Default", "wordSpacing": 100, "kerning": true}}, "shape": {"x": 1.405532031711079, "y": 10.165805251246155}, "text": "Report Title", "type": "CIMTextGraphic", "blendingMode": "Alpha"}, "lockedAspectRatio": true, "name": "ReportTitle", "anchor": "TopLeftCorner", "type": "CIMGraphicElement", "rotationCenter": {"x": 1.405532031711079, "y": 10.165805251246155}, "visible": true}, {"graphic": {"placement": "Unspecified", "symbol": {"type": "CIMSymbolReference", "symbol": {"fontFamilyName": "Tahoma", "haloSize": 1, "symbol": {"type": "CIMPolygonSymbol", "symbolLayers": [{"type": "CIMSolidFill", "enable": true, "color": {"type": "CIMRGBColor", "values": [255, 255, 255, 100]}}]}, "textCase": "Normal", "height": 12, "verticalAlignment": "Bottom", "symbol3DProperties": {"scaleZ": 0, "type": "CIM3DSymbolProperties", "rotationOrder3D": "XYZ", "dominantSizeAxis3D": "Z", "scaleY": 0}, "ligatures": true, "fontType": "Unspecified", "fontEncoding": "Unicode", "fontStyleName": "Regular", "lineGapType": "ExtraLeading", "letterWidth": 100, "verticalGlyphOrientation": "Right", "type": "CIMTextSymbol", "horizontalAlignment": "Left", "fontEffects": "Normal", "extrapolateBaselines": true, "blockProgression": "TTB", "textDirection": "LTR", "billboardMode3D": "FaceNearPlane", "depth3D": 1, "hinting": "Default", "wordSpacing": 100, "kerning": true}}, "shape": {"x": 6.591677299782655, "y": 9.914340407496155}, "text": "Analyst: <dyn type=\\"user\\"/>", "type": "CIMTextGraphic", "blendingMode": "Alpha"}, "lockedAspectRatio": true, "name": "CurrentUser", "anchor": "BottomLeftCorner", "type": "CIMGraphicElement", "rotationCenter": {"x": 6.590843966449322, "y": 9.914340407496155}, "visible": true}, {"graphic": {"placement": "Unspecified", "symbol": {"type": "CIMSymbolReference", "symbol": {"fontFamilyName": "Tahoma", "haloSize": 1, "symbol": {"type": "CIMPolygonSymbol", "symbolLayers": [{"type": "CIMSolidFill", "enable": true, "color": {"type": "CIMRGBColor", "values": [0, 0, 0, 255]}}]}, "textCase": "Normal", "height": 18, "verticalAlignment": "Bottom", "symbol3DProperties": {"scaleZ": 0, "type": "CIM3DSymbolProperties", "rotationOrder3D": "XYZ", "dominantSizeAxis3D": "Z", "scaleY": 0}, "ligatures": true, "fontType": "Unspecified", "fontEncoding": "Unicode", "fontStyleName": "Regular", "lineGapType": "ExtraLeading", "letterWidth": 100, "verticalGlyphOrientation": "Right", "type": "CIMTextSymbol", "horizontalAlignment": "Left", "fontEffects": "Normal", "extrapolateBaselines": true, "blockProgression": "TTB", "textDirection": "LTR", "billboardMode3D": "FaceNearPlane", "depth3D": 1, "hinting": "Default", "wordSpacing": 100, "kerning": true}}, "shape": {"x": 0.5934970669224628, "y": 9.407407407407405}, "text": "Report Sub-Title", "type": "CIMTextGraphic", "blendingMode": "Alpha"}, "lockedAspectRatio": true, "name": "ReportSubTitle", "anchor": "BottomLeftCorner", "type": "CIMGraphicElement", "rotationCenter": {"x": 0.5934970669224628, "y": 9.407407407407405}, "visible": true}, {"frame": {"rings": [[[5.697592627914723, 5.677083333333352], [5.697592627914723, 9.281250000000004], [7.8642592945813865, 9.281250000000004], [7.8642592945813865, 5.677083333333352], [5.697592627914723, 5.677083333333352]]]}, "defaultPatchHeight": 12, "mapFrame": "Map Map Frame", "autoAdd": false, "verticalPatchGap": 5, "textGap": 5, "autoFonts": true, "anchor": "BottomLeftCorner", "fittingStrategy": "AdjustColumnsAndSize", "scaleSymbols": true, "name": "Legend", "autoReorder": false, "showTitle": false, "horizontalPatchGap": 5, "visible": true, "graphicFrame": {"type": "CIMGraphicFrame", "backgroundSymbol": {"type": "CIMSymbolReference", "symbol": {"effects": [{"offset": 0, "type": "CIMGeometricEffectOffset", "count": 1, "option": "Fast", "method": "Square"}], "type": "CIMPolygonSymbol", "symbolLayers": [{"type": "CIMSolidFill", "enable": true, "color": {"type": "CIMRGBColor", "values": [255, 255, 255, 0]}}]}}, "borderSymbol": {"type": "CIMSymbolReference", "symbol": {"effects": [{"offset": 0, "type": "CIMGeometricEffectOffset", "count": 1, "option": "Fast", "method": "Square"}], "type": "CIMLineSymbol", "symbolLayers": [{"color": {"type": "CIMRGBColor", "values": [255, 255, 255, 0]}, "capStyle": "Round", "miterLimit": 10, "type": "CIMSolidStroke", "width": 1, "lineStyle3D": "Strip", "joinStyle": "Round", "enable": true}]}}}, "headingGap": 5, "title": "Legend", "horizontalItemGap": 5, "type": "CIMLegend", "layerNameGap": 5, "rotationCenter": {"x": 5.697592627914723, "y": 5.677083333333352}, "verticalItemGap": 5, "defaultPatchWidth": 24, "groupGap": 5, "minFontSize": 4, "titleGap": 5, "autoVisibility": true}, {"frame": {"rings": [[[5.697592627914726, 5.222222222222215], [5.697592627914726, 5.51627604166666], [7.744876554983919, 5.51627604166666], [7.744876554983919, 5.222222222222215], [5.697592627914726, 5.222222222222215]]]}, "mapFrame": "Map Map Frame", "subdivisionMarkSymbol": {"type": "CIMSymbolReference", "symbol": {"type": "CIMLineSymbol", "symbolLayers": [{"color": {"type": "CIMRGBColor", "values": [0, 0, 0, 100]}, "capStyle": "Round", "miterLimit": 10, "type": "CIMSolidStroke", "width": 1, "lineStyle3D": "Strip", "joinStyle": "Round", "enable": true}]}, "symbolName": "Symbol_1152"}, "anchor": "BottomLeftCorner", "divisionMarkHeight": 7, "fittingStrategy": "AdjustDivision", "subdivisionMarkHeight": 5, "units": {"uwkid": 9093}, "name": "Scale Bar", "markFrequency": "DivisionsAndSubdivisions", "visible": true, "unitLabelSymbol": {"type": "CIMSymbolReference", "symbol": {"fontFamilyName": "Arial", "haloSize": 1, "symbol": {"type": "CIMPolygonSymbol", "symbolLayers": [{"type": "CIMSolidFill", "enable": true, "color": {"type": "CIMRGBColor", "values": [0, 0, 0, 100]}}]}, "textCase": "Normal", "height": 10, "fontEncoding": "Unicode", "ligatures": false, "fontType": "Unspecified", "compatibilityMode": true, "fontStyleName": "Regular", "drawSoftHyphen": true, "lineGapType": "ExtraLeading", "letterWidth": 100, "verticalGlyphOrientation": "Right", "type": "CIMTextSymbol", "horizontalAlignment": "Left", "fontEffects": "Normal", "extrapolateBaselines": true, "blockProgression": "TTB", "textDirection": "LTR", "billboardMode3D": "FaceNearPlane", "depth3D": 1, "hinting": "Default", "wordSpacing": 100, "kerning": true, "shadowColor": {"type": "CIMRGBColor", "values": [0, 0, 0, 100]}, "verticalAlignment": "Center"}, "symbolName": "Symbol_1150"}, "labelGap": 3, "graphicFrame": {"type": "CIMGraphicFrame", "backgroundSymbol": {"type": "CIMSymbolReference", "symbol": {"effects": [{"offset": 0, "type": "CIMGeometricEffectOffset", "count": 1, "option": "Fast", "method": "Square"}], "type": "CIMPolygonSymbol", "symbolLayers": [{"type": "CIMSolidFill", "enable": true, "color": {"type": "CIMRGBColor", "values": [255, 255, 255, 0]}}]}}, "borderSymbol": {"type": "CIMSymbolReference", "symbol": {"effects": [{"offset": 0, "type": "CIMGeometricEffectOffset", "count": 1, "option": "Fast", "method": "Square"}], "type": "CIMLineSymbol", "symbolLayers": [{"color": {"type": "CIMRGBColor", "values": [255, 255, 255, 0]}, "capStyle": "Round", "miterLimit": 10, "type": "CIMSolidStroke", "width": 1, "lineStyle3D": "Strip", "joinStyle": "Round", "enable": true}]}}}, "labelFrequency": "DivisionsAndFirstMidpoint", "labelSymbol": {"type": "CIMSymbolReference", "symbol": {"fontFamilyName": "Arial", "haloSize": 1, "symbol": {"type": "CIMPolygonSymbol", "symbolLayers": [{"type": "CIMSolidFill", "enable": true, "color": {"type": "CIMRGBColor", "values": [0, 0, 0, 100]}}]}, "textCase": "Normal", "height": 10, "fontEncoding": "Unicode", "ligatures": false, "fontType": "Unspecified", "compatibilityMode": true, "fontStyleName": "Regular", "drawSoftHyphen": true, "lineGapType": "ExtraLeading", "letterWidth": 100, "verticalGlyphOrientation": "Right", "type": "CIMTextSymbol", "horizontalAlignment": "Center", "fontEffects": "Normal", "extrapolateBaselines": true, "blockProgression": "TTB", "textDirection": "LTR", "billboardMode3D": "FaceNearPlane", "depth3D": 1, "hinting": "Default", "wordSpacing": 100, "kerning": true, "shadowColor": {"type": "CIMRGBColor", "values": [0, 0, 0, 100]}, "verticalAlignment": "Baseline"}, "symbolName": "Symbol_1149"}, "labelPosition": "Above", "unitLabelGap": 3, "type": "CIMScaleLine", "unitLabelPosition": "AfterLabels", "rotationCenter": {"x": 5.697592627914726, "y": 5.222222222222215}, "divisionMarkSymbol": {"type": "CIMSymbolReference", "symbol": {"type": "CIMLineSymbol", "symbolLayers": [{"color": {"type": "CIMRGBColor", "values": [0, 0, 0, 100]}, "capStyle": "Round", "miterLimit": 10, "type": "CIMSolidStroke", "width": 1, "lineStyle3D": "Strip", "joinStyle": "Round", "enable": true}]}, "symbolName": "Symbol_1151"}, "divisionsBeforeZero": 0, "numberFormat": {"alignmentWidth": 12, "useSeparator": true, "roundingOption": "esriRoundNumberOfDecimals", "roundingValue": 2, "alignmentOption": "esriAlignLeft", "type": "CIMNumericFormat"}, "division": 7.5, "lineSymbol": {"type": "CIMSymbolReference", "symbol": {"type": "CIMLineSymbol", "symbolLayers": [{"color": {"type": "CIMRGBColor", "values": [0, 0, 0, 100]}, "capStyle": "Round", "miterLimit": 10, "type": "CIMSolidStroke", "width": 1, "lineStyle3D": "Strip", "joinStyle": "Round", "enable": true}]}, "symbolName": "Symbol_1148"}, "subdivisions": 4, "unitLabel": "Miles", "divisions": 2, "markPosition": "Above"}, {"graphic": {"placement": "Unspecified", "type": "CIMPolygonGraphic", "blendingMode": "Alpha", "symbol": {"type": "CIMSymbolReference", "symbol": {"type": "CIMPolygonSymbol", "symbolLayers": [{"color": {"type": "CIMRGBColor", "values": [0, 0, 0, 0]}, "capStyle": "Round", "miterLimit": 10, "type": "CIMSolidStroke", "width": 1, "lineStyle3D": "Strip", "joinStyle": "Round", "enable": true}, {"type": "CIMSolidFill", "enable": true, "color": {"type": "CIMRGBColor", "values": [11, 121.99610900878906, 192, 0]}}]}}, "polygon": {"hasZ": true, "rings": [[[0.49618708959743874, 5.0231481481481515, null], [7.885075978486328, 5.0231481481481515, null], [7.885075978486328, 4.623822420282124, null], [0.49618708959743874, 4.623822420282124, null], [0.49618708959743874, 5.0231481481481515, null]]]}}, "name": "TableHeaderBackground", "anchor": "TopLeftCorner", "type": "CIMGraphicElement", "rotationCenter": {"x": 0.49618708959743874, "y": 5.0231481481481515}, "visible": true}, {"graphic": {"placement": "Unspecified", "symbol": {"type": "CIMSymbolReference", "symbol": {"fontFamilyName": "Tahoma", "haloSize": 1, "symbol": {"type": "CIMPolygonSymbol", "symbolLayers": [{"type": "CIMSolidFill", "enable": true, "color": {"type": "CIMRGBColor", "values": [11, 121.99610900878906, 192, 100]}}]}, "textCase": "Normal", "height": 14, "verticalAlignment": "Top", "symbol3DProperties": {"scaleZ": 0, "type": "CIM3DSymbolProperties", "rotationOrder3D": "XYZ", "dominantSizeAxis3D": "Z", "scaleY": 0}, "ligatures": true, "fontType": "Unspecified", "fontEncoding": "Unicode", "fontStyleName": "Regular", "lineGapType": "ExtraLeading", "letterWidth": 100, "verticalGlyphOrientation": "Right", "type": "CIMTextSymbol", "horizontalAlignment": "Left", "fontEffects": "Normal", "extrapolateBaselines": true, "blockProgression": "TTB", "textDirection": "LTR", "billboardMode3D": "FaceNearPlane", "depth3D": 1, "hinting": "Default", "wordSpacing": 100, "kerning": true}}, "shape": {"x": 0.5934970669224624, "y": 4.988212528935189}, "text": "Table Title", "type": "CIMTextGraphic", "blendingMode": "Alpha"}, "lockedAspectRatio": true, "name": "TableTitle", "anchor": "TopLeftCorner", "type": "CIMGraphicElement", "rotationCenter": {"x": 0.5934970669224624, "y": 4.988212528935188}, "visible": true}, {"graphic": {"placement": "Unspecified", "type": "CIMPolygonGraphic", "blendingMode": "Alpha", "symbol": {"type": "CIMSymbolReference", "symbol": {"type": "CIMPolygonSymbol", "symbolLayers": [{"color": {"type": "CIMRGBColor", "values": [0, 0, 0, 0]}, "capStyle": "Round", "miterLimit": 10, "type": "CIMSolidStroke", "width": 1, "lineStyle3D": "Strip", "joinStyle": "Round", "enable": true}, {"type": "CIMSolidFill", "enable": true, "color": {"type": "CIMGrayColor", "values": [228.435791015625, 100]}}]}}, "polygon": {"hasZ": true, "rings": [[[0.5220912518364877, 4.1728107934741105, null], [2.5124186327888682, 4.1728107934741105, null], [2.5124186327888682, 3.908869128536222, null], [0.5220912518364877, 3.908869128536222, null], [0.5220912518364877, 4.1728107934741105, null]]]}}, "name": "EvenRowBackground", "anchor": "TopLeftCorner", "type": "CIMGraphicElement", "rotationCenter": {"x": 0.5220912518364877, "y": 4.1728107934741105}, "visible": true}, {"graphic": {"placement": "Unspecified", "symbol": {"type": "CIMSymbolReference", "symbol": {"fontFamilyName": "Tahoma", "haloSize": 1, "symbol": {"type": "CIMPolygonSymbol", "symbolLayers": [{"type": "CIMSolidFill", "enable": true, "color": {"type": "CIMRGBColor", "values": [255, 255, 255, 100]}}]}, "textCase": "Normal", "height": 12, "verticalAlignment": "Bottom", "symbol3DProperties": {"scaleZ": 0, "type": "CIM3DSymbolProperties", "rotationOrder3D": "XYZ", "dominantSizeAxis3D": "Z", "scaleY": 0}, "ligatures": true, "fontType": "Unspecified", "fontEncoding": "Unicode", "fontStyleName": "Regular", "lineGapType": "ExtraLeading", "letterWidth": 100, "verticalGlyphOrientation": "Right", "type": "CIMTextSymbol", "horizontalAlignment": "Left", "fontEffects": "Normal", "extrapolateBaselines": true, "blockProgression": "TTB", "textDirection": "LTR", "billboardMode3D": "FaceNearPlane", "depth3D": 1, "hinting": "Default", "wordSpacing": 100, "kerning": true}}, "shape": {"x": 6.820579246238428, "y": 10.210238844996162}, "text": "Date Exported <dyn type=\\"layout\\" name=\\"MapLayout\\" property=\\"dateExported\\" format=\\"short|short\\"/>", "type": "CIMTextGraphic", "blendingMode": "Alpha"}, "lockedAspectRatio": true, "name": "PrintDate", "anchor": "BottomLeftCorner", "type": "CIMGraphicElement", "rotationCenter": {"x": 6.820579246238428, "y": 10.210238844996162}, "visible": true}, {"graphic": {"frame": {"type": "CIMGraphicFrame", "backgroundSymbol": {"type": "CIMSymbolReference", "symbol": {"effects": [{"offset": 0, "type": "CIMGeometricEffectOffset", "count": 1, "option": "Fast", "method": "Square"}], "type": "CIMPolygonSymbol", "symbolLayers": [{"type": "CIMSolidFill", "enable": true, "color": {"type": "CIMRGBColor", "values": [255, 255, 255, 0]}}]}}, "borderSymbol": {"type": "CIMSymbolReference", "symbol": {"effects": [{"offset": 0, "type": "CIMGeometricEffectOffset", "count": 1, "option": "Fast", "method": "Square"}], "type": "CIMLineSymbol", "symbolLayers": [{"color": {"type": "CIMRGBColor", "values": [255, 255, 255, 0]}, "capStyle": "Round", "miterLimit": 10, "type": "CIMSolidStroke", "width": 1, "lineStyle3D": "Strip", "joinStyle": "Round", "enable": true}]}}}, "placement": "BottomLeftCorner", "box": {"zmin": 0, "zmax": 0, "ymin": 9.80082243616118, "ymax": 10.491177563838818, "xmax": 1.2693865500071992, "xmin": 0.5000000000000001}, "pictureURL": "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAJAAAACQCAYAAADnRuK4AAAAGXRFWHRTb2Z0d2FyZQBBZG9iZSBJbWFnZVJlYWR5ccllPAAAAyJpVFh0WE1MOmNvbS5hZG9iZS54bXAAAAAAADw/eHBhY2tldCBiZWdpbj0i77u/IiBpZD0iVzVNME1wQ2VoaUh6cmVTek5UY3prYzlkIj8+IDx4OnhtcG1ldGEgeG1sbnM6eD0iYWRvYmU6bnM6bWV0YS8iIHg6eG1wdGs9IkFkb2JlIFhNUCBDb3JlIDUuMy1jMDExIDY2LjE0NTY2MSwgMjAxMi8wMi8wNi0xNDo1NjoyNyAgICAgICAgIj4gPHJkZjpSREYgeG1sbnM6cmRmPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5LzAyLzIyLXJkZi1zeW50YXgtbnMjIj4gPHJkZjpEZXNjcmlwdGlvbiByZGY6YWJvdXQ9IiIgeG1sbnM6eG1wPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvIiB4bWxuczp4bXBNTT0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL21tLyIgeG1sbnM6c3RSZWY9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9zVHlwZS9SZXNvdXJjZVJlZiMiIHhtcDpDcmVhdG9yVG9vbD0iQWRvYmUgUGhvdG9zaG9wIENTNiAoV2luZG93cykiIHhtcE1NOkluc3RhbmNlSUQ9InhtcC5paWQ6OUExMDlGRjNGMEFCMTFFMzkzNTBDRUI5OURCRjEwQUEiIHhtcE1NOkRvY3VtZW50SUQ9InhtcC5kaWQ6OUExMDlGRjRGMEFCMTFFMzkzNTBDRUI5OURCRjEwQUEiPiA8eG1wTU06RGVyaXZlZEZyb20gc3RSZWY6aW5zdGFuY2VJRD0ieG1wLmlpZDo5QTEwOUZGMUYwQUIxMUUzOTM1MENFQjk5REJGMTBBQSIgc3RSZWY6ZG9jdW1lbnRJRD0ieG1wLmRpZDo5QTEwOUZGMkYwQUIxMUUzOTM1MENFQjk5REJGMTBBQSIvPiA8L3JkZjpEZXNjcmlwdGlvbj4gPC9yZGY6UkRGPiA8L3g6eG1wbWV0YT4gPD94cGFja2V0IGVuZD0iciI/PmCxY3sAABjGSURBVHja7F0LvJfj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", "sourceURL": "C:\\\\data\\\\State\\\\env impact\\\\201602 Release\\\\DownloadStaging\\\\Application\\\\js\\\\library\\\\themes\\\\images\\\\logo.png", "type": "CIMPictureGraphic", "blendingMode": "Alpha"}, "lockedAspectRatio": true, "name": "Logo", "anchor": "BottomLeftCorner", "type": "CIMGraphicElement", "rotationCenter": {"x": 0.5000000000000001, "y": 9.80082243616118}, "visible": true}, {"graphic": {"placement": "Unspecified", "symbol": {"type": "CIMSymbolReference", "symbol": {"fontFamilyName": "Tahoma", "haloSize": 1, "symbol": {"type": "CIMPolygonSymbol", "symbolLayers": [{"type": "CIMSolidFill", "enable": true, "color": {"type": "CIMRGBColor", "values": [130, 130, 130, 100]}}]}, "textCase": "Normal", "height": 8, "verticalAlignment": "Bottom", "symbol3DProperties": {"scaleZ": 0, "type": "CIM3DSymbolProperties", "rotationOrder3D": "XYZ", "dominantSizeAxis3D": "Z", "scaleY": 0}, "ligatures": true, "fontType": "Unspecified", "fontEncoding": "Unicode", "fontStyleName": "Regular", "lineGapType": "ExtraLeading", "letterWidth": 100, "verticalGlyphOrientation": "Right", "type": "CIMTextSymbol", "horizontalAlignment": "Left", "fontEffects": "Normal", "extrapolateBaselines": true, "blockProgression": "TTB", "textDirection": "LTR", "billboardMode3D": "FaceNearPlane", "depth3D": 1, "hinting": "Default", "wordSpacing": 100, "kerning": true}}, "shape": {"x": 0.5039586141498883, "z": 0, "y": 0.5}, "text": "Page 1 of #", "type": "CIMTextGraphic", "blendingMode": "Alpha"}, "lockedAspectRatio": true, "name": "PageNumber", "anchor": "BottomLeftCorner", "type": "CIMGraphicElement", "rotationCenter": {"x": 0.5039586141498883, "y": 0.5}, "visible": true}, {"graphic": {"placement": "Unspecified", "type": "CIMLineGraphic", "blendingMode": "Alpha", "symbol": {"type": "CIMSymbolReference", "symbol": {"type": "CIMLineSymbol", "symbolLayers": [{"color": {"type": "CIMRGBColor", "values": [11, 121.99610900878906, 192, 100]}, "capStyle": "Round", "miterLimit": 10, "type": "CIMSolidStroke", "width": 0.5, "lineStyle3D": "Strip", "joinStyle": "Round", "enable": true}]}}, "line": {"paths": [[[0.500000000000002, 0.6800027946275057, null], [8.000000000000004, 0.68, null]]], "hasZ": true}}, "name": "FooterLine", "anchor": "BottomLeftCorner", "type": "CIMGraphicElement", "rotationCenter": {"x": 0.500000000000002, "y": 0.68}, "visible": true}, {"graphic": {"placement": "Unspecified", "symbol": {"type": "CIMSymbolReference", "symbol": {"fontFamilyName": "Tahoma", "haloSize": 1, "symbol": {"type": "CIMPolygonSymbol", "symbolLayers": [{"type": "CIMSolidFill", "enable": true, "color": {"type": "CIMRGBColor", "values": [130, 130, 130, 100]}}]}, "textCase": "Normal", "height": 8, "verticalAlignment": "Bottom", "symbol3DProperties": {"scaleZ": 0, "type": "CIM3DSymbolProperties", "rotationOrder3D": "XYZ", "dominantSizeAxis3D": "Z", "scaleY": 0}, "ligatures": true, "fontType": "Unspecified", "fontEncoding": "Unicode", "fontStyleName": "Regular", "lineGapType": "ExtraLeading", "letterWidth": 100, "verticalGlyphOrientation": "Right", "type": "CIMTextSymbol", "horizontalAlignment": "Right", "fontEffects": "Normal", "extrapolateBaselines": true, "blockProgression": "TTB", "textDirection": "LTR", "billboardMode3D": "FaceNearPlane", "depth3D": 1, "hinting": "Default", "wordSpacing": 100, "kerning": true}}, "shape": {"x": 8, "y": 0.5019678941441441}, "text": "Report Title", "type": "CIMTextGraphic", "blendingMode": "Alpha"}, "lockedAspectRatio": true, "name": "ReportTitleFooter", "anchor": "BottomRightCorner", "type": "CIMGraphicElement", "rotationCenter": {"x": 7.999999999999999, "y": 0.5019678941441441}, "visible": true}], "metadataURI": "CIMPATH=Metadata/2edbf119be763c7b7b29b3cc4aeaf083.xml", "type": "CIMLayout", "page": {"height": 11, "units": {"uwkid": 109008}, "showRulers": true, "showGuides": true, "smallestRulerDivision": 0, "type": "CIMPage", "width": 8.5, "guides": [{"orientation": "Horizontal", "type": "CIMGuide", "position": 0.5}, {"orientation": "Horizontal", "type": "CIMGuide", "position": 10.5}, {"orientation": "Vertical", "type": "CIMGuide", "position": 0.5}, {"orientation": "Vertical", "type": "CIMGuide", "position": 8}, {"orientation": "Horizontal", "type": "CIMGuide", "position": 0.75}]}, "uRI": "CIMPATH=layout/maplayout.xml", "sourceModifiedTime": {"type": "TimeInstant"}}, "build": 5023, "layerDefinitions": [{"backgroundColor": {"type": "CIMRGBColor", "values": [254, 254, 254, 100]}, "sourceModifiedTime": {"type": "TimeInstant"}, "showLegends": true, "serviceLayerID": -1, "displayCacheType": "Permanent", "type": "CIMTiledServiceLayer", "transparentColor": {"type": "CIMRGBColor", "values": [254, 254, 254, 100]}, "visibility": true, "maxDisplayCacheAge": 5, "description": "World_Topo_Map", "layer3DProperties": {"maxDistance": -1, "preloadTextureCutoffHigh": 0, "textureCutoffHigh": 0.25, "verticalUnit": {"uwkid": 9001}, "useCompressedTextures": true, "type": "CIM3DLayerProperties", "textureCutoffLow": 1, "verticalExaggeration": 1, "lighting": "OneSideDataNormal", "isLayerLit": true, "preloadTextureCutoffLow": 0.25, "castShadows": true, "minDistance": -1, "layerFaceCulling": "None"}, "name": "Topographic", "serviceConnection": {"objectType": "MapServer", "type": "CIMAGSServiceConnection", "serverConnection": {"url": "http://services.arcgisonline.com/ArcGIS/services", "type": "CIMInternetServerConnection", "hideUserProperty": true, "anonymous": true}, "objectName": "World_Topo_Map", "url": "http://services.arcgisonline.com/ArcGIS/rest/services/World_Topo_Map/MapServer"}, "showPopups": false, "maxScale": 70.5310735, "layerType": "BasemapBackground", "uRI": "CIMPATH=cc292b32a452498da8f8360b2a14faa3.xml", "minScale": 591657527.591555}], "type": "CIMLayoutDocument", "mapDefinitions": [{"mapType": "Map", "metadataURI": "CIMPATH=Metadata/ccaf3d769eb8676927a18c0b5bed382a.xml", "type": "CIMMap", "illumination": {"ambientLight": 75, "sunPositionZ": 0.5, "sunAzimuth": 315, "sunPositionX": -0.61237243569579, "type": "CIMIlluminationProperties", "illuminationSource": "AbsoluteSunPosition", "sunAltitude": 30, "sunPositionY": 0.61237243569579}, "layers": ["CIMPATH=cc292b32a452498da8f8360b2a14faa3.xml"], "spatialReference": {"wkid": 102100, "latestWkid": 3857}, "defaultViewingMode": "Map", "generalPlacementProperties": {"type": "CIMMaplexGeneralPlacementProperties", "placementQuality": "High", "unplacedLabelColor": {"type": "CIMRGBColor", "values": [255, 0, 0, 100]}, "invertedLabelTolerance": 2, "keyNumberGroups": [{"minimumNumberOfLines": 2, "horizontalAlignment": "Left", "name": "Default", "maximumNumberOfLines": 20, "type": "CIMMaplexKeyNumberGroup", "delimiterCharacter": ".", "numberResetType": "None"}]}, "defaultExtent": {"ymin": -11999999.999983612, "spatialReference": {"wkid": 102100, "latestWkid": 3857}, "ymax": 15999999.999965921, "xmax": 17999999.99998911, "xmin": -17999999.99998911}, "snappingProperties": {"snapRequestType": "SnapRequestType_GeometricAndVisualSnapping", "xYToleranceUnit": "SnapXYToleranceUnitPixel", "type": "CIMSnappingProperties", "snapToSketchEnabled": true, "xYTolerance": 10}, "sourceModifiedTime": {"type": "TimeInstant"}, "uRI": "CIMPATH=map/map3.xml", "name": "Map"}]}'
            pagx = self.write_pagx(map_layout_json, self.map_layout_name)
        else:
            base_name = os.path.splitext(os.path.basename(template))[0]
            self.map_layout_name = base_name + str(self.time_stamp)
            pagx = self.update_template(template, self.map_layout_name, self.map_element_names)
        self.map_pagx = pagx
        self.aprx.importDocument(self.map_pagx)

    def init_overflow_layout(self, template):
        pagx = None
        if template in ['', '#', ' ', None]:
            self.overflow_layout_name = 'OverflowLayout{0}'.format(self.time_stamp)
            overflow_layout_json = '{"version": "1.2.0", "build": 5023, "layoutDefinition": {"name": "' + self.overflow_layout_name + '", "datePrinted": {"type": "TimeInstant"}, "sourceModifiedTime": {"type": "TimeInstant"}, "type": "CIMLayout", "uRI": "CIMPATH=layout/layout1.xml", "dateExported": {"type": "TimeInstant"}, "elements": [{"name": "horzLine", "graphic": {"symbol": {"symbol": {"symbolLayers": [{"enable": true, "miterLimit": 10, "capStyle": "Round", "width": 1, "color": {"values": [0, 0, 0, 100], "type": "CIMRGBColor"}, "joinStyle": "Round", "lineStyle3D": "Strip", "type": "CIMSolidStroke"}], "type": "CIMLineSymbol"}, "type": "CIMSymbolReference"}, "placement": "Unspecified", "line": {"paths": [[[0.728072303850187, 9.59488237564073, null], [2.728072303850187, 9.59488237564073, null]]], "hasZ": true}, "blendingMode": "Alpha", "type": "CIMLineGraphic"}, "visible": true, "anchor": "TopLeftCorner", "type": "CIMGraphicElement", "rotationCenter": {"y": 9.59488237564073, "x": 0.728072303850187}}, {"name": "vertLine", "graphic": {"symbol": {"symbol": {"symbolLayers": [{"enable": true, "miterLimit": 10, "capStyle": "Round", "width": 1, "color": {"values": [0, 0, 0, 100], "type": "CIMRGBColor"}, "joinStyle": "Round", "lineStyle3D": "Strip", "type": "CIMSolidStroke"}], "type": "CIMLineSymbol"}, "type": "CIMSymbolReference"}, "placement": "Unspecified", "line": {"paths": [[[1.7176573062754796, 9.574939179268307, null], [1.720282163740743, 7.574939179268307, null]]], "hasZ": true}, "blendingMode": "Alpha", "type": "CIMLineGraphic"}, "visible": true, "anchor": "TopLeftCorner", "type": "CIMGraphicElement", "rotationCenter": {"y": 9.574939179268307, "x": 1.7176573062754796}}, {"name": "ContentDisplayArea", "graphic": {"symbol": {"symbol": {"symbolLayers": [{"enable": true, "miterLimit": 10, "capStyle": "Round", "width": 1, "color": {"values": [0, 0, 0, 100], "type": "CIMRGBColor"}, "joinStyle": "Round", "lineStyle3D": "Strip", "type": "CIMSolidStroke"}], "type": "CIMPolygonSymbol"}, "type": "CIMSymbolReference"}, "placement": "Unspecified", "polygon": {"rings": [[[0.5039586141498891, 9.999999999999948, null], [8.000000000000004, 9.999999999999948, null], [8.000000000000004, 0.7500000000000018, null], [0.5039586141498891, 0.7500000000000018, null], [0.5039586141498891, 9.999999999999948, null]]], "hasZ": true}, "blendingMode": "Alpha", "type": "CIMPolygonGraphic"}, "visible": true, "anchor": "TopLeftCorner", "type": "CIMGraphicElement", "rotationCenter": {"y": 9.999999999999948, "x": 0.5039586141498891}}, {"name": "FooterLine", "graphic": {"symbol": {"symbol": {"symbolLayers": [{"enable": true, "miterLimit": 10, "capStyle": "Round", "width": 0.5, "color": {"values": [11, 121.99610900878906, 192, 100], "type": "CIMRGBColor"}, "joinStyle": "Round", "lineStyle3D": "Strip", "type": "CIMSolidStroke"}], "type": "CIMLineSymbol"}, "type": "CIMSymbolReference"}, "placement": "Unspecified", "line": {"paths": [[[0.500000000000002, 0.6800027946275057, null], [8.000000000000004, 0.68, null]]], "hasZ": true}, "blendingMode": "Alpha", "type": "CIMLineGraphic"}, "visible": true, "anchor": "BottomLeftCorner", "type": "CIMGraphicElement", "rotationCenter": {"y": 0.68, "x": 0.500000000000002}}, {"name": "PageNumber", "lockedAspectRatio": true, "graphic": {"placement": "Unspecified", "text": "Page 1 of #", "type": "CIMTextGraphic", "shape": {"z": 0, "y": 0.5, "x": 0.5039586141498883}, "symbol": {"symbol": {"fontEncoding": "Unicode", "haloSize": 1, "textCase": "Normal", "kerning": true, "type": "CIMTextSymbol", "verticalAlignment": "Bottom", "textDirection": "LTR", "symbol3DProperties": {"scaleZ": 0, "rotationOrder3D": "XYZ", "scaleY": 0, "dominantSizeAxis3D": "Z", "type": "CIM3DSymbolProperties"}, "height": 8, "fontEffects": "Normal", "lineGapType": "ExtraLeading", "blockProgression": "TTB", "hinting": "Default", "wordSpacing": 100, "billboardMode3D": "FaceNearPlane", "verticalGlyphOrientation": "Right", "fontType": "Unspecified", "horizontalAlignment": "Left", "fontFamilyName": "Tahoma", "depth3D": 1, "fontStyleName": "Regular", "symbol": {"symbolLayers": [{"enable": true, "color": {"values": [130, 130, 130, 100], "type": "CIMRGBColor"}, "type": "CIMSolidFill"}], "type": "CIMPolygonSymbol"}, "extrapolateBaselines": true, "ligatures": true, "letterWidth": 100}, "type": "CIMSymbolReference"}, "blendingMode": "Alpha"}, "visible": true, "anchor": "BottomLeftCorner", "type": "CIMGraphicElement", "rotationCenter": {"y": 0.5, "x": 0.5039586141498883}}, {"name": "ReportTitleFooter", "lockedAspectRatio": true, "graphic": {"placement": "Unspecified", "text": "Report Title", "type": "CIMTextGraphic", "shape": {"y": 0.5019678941441441, "x": 8}, "symbol": {"symbol": {"fontEncoding": "Unicode", "haloSize": 1, "textCase": "Normal", "kerning": true, "type": "CIMTextSymbol", "verticalAlignment": "Bottom", "textDirection": "LTR", "symbol3DProperties": {"scaleZ": 0, "rotationOrder3D": "XYZ", "scaleY": 0, "dominantSizeAxis3D": "Z", "type": "CIM3DSymbolProperties"}, "height": 8, "fontEffects": "Normal", "lineGapType": "ExtraLeading", "blockProgression": "TTB", "hinting": "Default", "wordSpacing": 100, "billboardMode3D": "FaceNearPlane", "verticalGlyphOrientation": "Right", "fontType": "Unspecified", "horizontalAlignment": "Right", "fontFamilyName": "Tahoma", "depth3D": 1, "fontStyleName": "Regular", "symbol": {"symbolLayers": [{"enable": true, "color": {"values": [130, 130, 130, 100], "type": "CIMRGBColor"}, "type": "CIMSolidFill"}], "type": "CIMPolygonSymbol"}, "extrapolateBaselines": true, "ligatures": true, "letterWidth": 100}, "type": "CIMSymbolReference"}, "blendingMode": "Alpha"}, "visible": true, "anchor": "BottomRightCorner", "type": "CIMGraphicElement", "rotationCenter": {"y": 0.5019678941441441, "x": 7.999999999999999}}, {"name": "EvenRowBackground", "graphic": {"symbol": {"symbol": {"symbolLayers": [{"enable": true, "miterLimit": 10, "capStyle": "Round", "width": 1, "color": {"values": [0, 0, 0, 0], "type": "CIMRGBColor"}, "joinStyle": "Round", "lineStyle3D": "Strip", "type": "CIMSolidStroke"}, {"enable": true, "color": {"values": [228.435791015625, 100], "type": "CIMGrayColor"}, "type": "CIMSolidFill"}], "type": "CIMPolygonSymbol"}, "type": "CIMSymbolReference"}, "placement": "Unspecified", "polygon": {"rings": [[[0.5814597177946053, 9.078216198879518, null], [2.5717870987469857, 9.078216198879518, null], [2.5717870987469857, 8.81427453394163, null], [0.5814597177946053, 8.81427453394163, null], [0.5814597177946053, 9.078216198879518, null]]], "hasZ": true}, "blendingMode": "Alpha", "type": "CIMPolygonGraphic"}, "visible": true, "anchor": "TopLeftCorner", "type": "CIMGraphicElement", "rotationCenter": {"y": 9.078216198879518, "x": 0.5814597177946053}}, {"name": "TableTitle", "lockedAspectRatio": true, "graphic": {"placement": "Unspecified", "text": "Table Title", "type": "CIMTextGraphic", "shape": {"y": 9.893617934340599, "x": 0.6528655328805799}, "symbol": {"symbol": {"fontEncoding": "Unicode", "haloSize": 1, "textCase": "Normal", "kerning": true, "type": "CIMTextSymbol", "verticalAlignment": "Top", "textDirection": "LTR", "symbol3DProperties": {"scaleZ": 0, "rotationOrder3D": "XYZ", "scaleY": 0, "dominantSizeAxis3D": "Z", "type": "CIM3DSymbolProperties"}, "height": 14, "fontEffects": "Normal", "lineGapType": "ExtraLeading", "blockProgression": "TTB", "hinting": "Default", "wordSpacing": 100, "billboardMode3D": "FaceNearPlane", "verticalGlyphOrientation": "Right", "fontType": "Unspecified", "horizontalAlignment": "Left", "fontFamilyName": "Tahoma", "depth3D": 1, "fontStyleName": "Regular", "symbol": {"symbolLayers": [{"enable": true, "color": {"values": [11, 121.99610900878906, 192, 100], "type": "CIMRGBColor"}, "type": "CIMSolidFill"}], "type": "CIMPolygonSymbol"}, "extrapolateBaselines": true, "ligatures": true, "letterWidth": 100}, "type": "CIMSymbolReference"}, "blendingMode": "Alpha"}, "visible": true, "anchor": "TopLeftCorner", "type": "CIMGraphicElement", "rotationCenter": {"y": 9.893617934340597, "x": 0.6528655328805799}}, {"name": "TableHeaderBackground", "graphic": {"symbol": {"symbol": {"symbolLayers": [{"enable": true, "miterLimit": 10, "capStyle": "Round", "width": 1, "color": {"values": [0, 0, 0, 0], "type": "CIMRGBColor"}, "joinStyle": "Round", "lineStyle3D": "Strip", "type": "CIMSolidStroke"}, {"enable": true, "color": {"values": [11, 121.99610900878906, 192, 0], "type": "CIMRGBColor"}, "type": "CIMSolidFill"}], "type": "CIMPolygonSymbol"}, "type": "CIMSymbolReference"}, "placement": "Unspecified", "polygon": {"rings": [[[0.5555555555555562, 9.92855355355356, null], [7.9444444444444455, 9.92855355355356, null], [7.9444444444444455, 9.529227825687531, null], [0.5555555555555562, 9.529227825687531, null], [0.5555555555555562, 9.92855355355356, null]]], "hasZ": true}, "blendingMode": "Alpha", "type": "CIMPolygonGraphic"}, "visible": true, "anchor": "TopLeftCorner", "type": "CIMGraphicElement", "rotationCenter": {"y": 9.92855355355356, "x": 0.5555555555555562}}, {"name": "FieldValue", "lockedAspectRatio": true, "graphic": {"placement": "Unspecified", "text": "Field Value", "type": "CIMTextGraphic", "shape": {"y": 9.319483871206735, "x": 0.6528655328805799}, "symbol": {"symbol": {"fontEncoding": "Unicode", "haloSize": 1, "textCase": "Normal", "kerning": true, "type": "CIMTextSymbol", "verticalAlignment": "Top", "textDirection": "LTR", "symbol3DProperties": {"scaleZ": 0, "rotationOrder3D": "XYZ", "scaleY": 0, "dominantSizeAxis3D": "Z", "type": "CIM3DSymbolProperties"}, "height": 10, "fontEffects": "Normal", "lineGapType": "ExtraLeading", "blockProgression": "TTB", "hinting": "Default", "wordSpacing": 100, "billboardMode3D": "FaceNearPlane", "verticalGlyphOrientation": "Right", "fontType": "TTOpenType", "horizontalAlignment": "Left", "fontFamilyName": "Arial", "depth3D": 1, "fontStyleName": "Regular", "symbol": {"symbolLayers": [{"enable": true, "color": {"values": [0, 0, 0, 100], "type": "CIMRGBColor"}, "type": "CIMSolidFill"}], "type": "CIMPolygonSymbol"}, "extrapolateBaselines": true, "ligatures": true, "letterWidth": 100}, "type": "CIMSymbolReference"}, "blendingMode": "Alpha"}, "visible": true, "anchor": "TopLeftCorner", "type": "CIMGraphicElement", "rotationCenter": {"y": 9.319483871206735, "x": 0.6528655328805799}}, {"name": "FieldName", "lockedAspectRatio": true, "graphic": {"placement": "Unspecified", "text": "Field Name", "type": "CIMTextGraphic", "shape": {"y": 9.529227825687531, "x": 0.6528655328805799}, "symbol": {"symbol": {"fontEncoding": "Unicode", "haloSize": 1, "textCase": "Normal", "kerning": true, "type": "CIMTextSymbol", "verticalAlignment": "Top", "textDirection": "LTR", "symbol3DProperties": {"scaleZ": 0, "rotationOrder3D": "XYZ", "scaleY": 0, "dominantSizeAxis3D": "Z", "type": "CIM3DSymbolProperties"}, "height": 10, "fontEffects": "Normal", "lineGapType": "ExtraLeading", "blockProgression": "TTB", "hinting": "Default", "wordSpacing": 100, "billboardMode3D": "FaceNearPlane", "verticalGlyphOrientation": "Right", "fontType": "TTOpenType", "horizontalAlignment": "Left", "fontFamilyName": "Arial", "depth3D": 1, "fontStyleName": "Bold", "symbol": {"symbolLayers": [{"enable": true, "color": {"values": [0, 0, 0, 100], "type": "CIMRGBColor"}, "type": "CIMSolidFill"}], "type": "CIMPolygonSymbol"}, "extrapolateBaselines": true, "ligatures": true, "letterWidth": 100}, "type": "CIMSymbolReference"}, "blendingMode": "Alpha"}, "visible": true, "anchor": "TopLeftCorner", "type": "CIMGraphicElement", "rotationCenter": {"y": 9.529227825687531, "x": 0.6528655328805799}}, {"name": "Logo", "lockedAspectRatio": true, "graphic": {"placement": "BottomLeftCorner", "box": {"ymax": 10.5, "zmax": 0, "xmax": 0.987584725792505, "zmin": 0, "xmin": 0.5000000000000001, "ymin": 10.062499999999998}, "type": "CIMPictureGraphic", "frame": {"borderSymbol": {"symbol": {"effects": [{"offset": 0, "option": "Fast", "method": "Square", "type": "CIMGeometricEffectOffset", "count": 1}], "symbolLayers": [{"enable": true, "miterLimit": 10, "capStyle": "Round", "width": 1, "color": {"values": [255, 255, 255, 0], "type": "CIMRGBColor"}, "joinStyle": "Round", "lineStyle3D": "Strip", "type": "CIMSolidStroke"}], "type": "CIMLineSymbol"}, "type": "CIMSymbolReference"}, "backgroundSymbol": {"symbol": {"effects": [{"offset": 0, "option": "Fast", "method": "Square", "type": "CIMGeometricEffectOffset", "count": 1}], "symbolLayers": [{"enable": true, "color": {"values": [255, 255, 255, 0], "type": "CIMRGBColor"}, "type": "CIMSolidFill"}], "type": "CIMPolygonSymbol"}, "type": "CIMSymbolReference"}, "type": "CIMGraphicFrame"}, "sourceURL": "C:\\\\data\\\\State\\\\env impact\\\\201602 Release\\\\DownloadStaging\\\\Application\\\\js\\\\library\\\\themes\\\\images\\\\logo.png", "pictureURL": "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAJAAAACQCAYAAADnRuK4AAAAGXRFWHRTb2Z0d2FyZQBBZG9iZSBJbWFnZVJlYWR5ccllPAAAAyJpVFh0WE1MOmNvbS5hZG9iZS54bXAAAAAAADw/eHBhY2tldCBiZWdpbj0i77u/IiBpZD0iVzVNME1wQ2VoaUh6cmVTek5UY3prYzlkIj8+IDx4OnhtcG1ldGEgeG1sbnM6eD0iYWRvYmU6bnM6bWV0YS8iIHg6eG1wdGs9IkFkb2JlIFhNUCBDb3JlIDUuMy1jMDExIDY2LjE0NTY2MSwgMjAxMi8wMi8wNi0xNDo1NjoyNyAgICAgICAgIj4gPHJkZjpSREYgeG1sbnM6cmRmPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5LzAyLzIyLXJkZi1zeW50YXgtbnMjIj4gPHJkZjpEZXNjcmlwdGlvbiByZGY6YWJvdXQ9IiIgeG1sbnM6eG1wPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvIiB4bWxuczp4bXBNTT0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL21tLyIgeG1sbnM6c3RSZWY9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9zVHlwZS9SZXNvdXJjZVJlZiMiIHhtcDpDcmVhdG9yVG9vbD0iQWRvYmUgUGhvdG9zaG9wIENTNiAoV2luZG93cykiIHhtcE1NOkluc3RhbmNlSUQ9InhtcC5paWQ6OUExMDlGRjNGMEFCMTFFMzkzNTBDRUI5OURCRjEwQUEiIHhtcE1NOkRvY3VtZW50SUQ9InhtcC5kaWQ6OUExMDlGRjRGMEFCMTFFMzkzNTBDRUI5OURCRjEwQUEiPiA8eG1wTU06RGVyaXZlZEZyb20gc3RSZWY6aW5zdGFuY2VJRD0ieG1wLmlpZDo5QTEwOUZGMUYwQUIxMUUzOTM1MENFQjk5REJGMTBBQSIgc3RSZWY6ZG9jdW1lbnRJRD0ieG1wLmRpZDo5QTEwOUZGMkYwQUIxMUUzOTM1MENFQjk5REJGMTBBQSIvPiA8L3JkZjpEZXNjcmlwdGlvbj4gPC9yZGY6UkRGPiA8L3g6eG1wbWV0YT4gPD94cGFja2V0IGVuZD0iciI/PmCxY3sAABjGSURBVHja7F0LvJfj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", "blendingMode": "Alpha"}, "visible": true, "anchor": "BottomLeftCorner", "type": "CIMGraphicElement", "rotationCenter": {"y": 10.062499999999998, "x": 0.5000000000000001}}, {"name": "ReportHeaderBackground", "graphic": {"symbol": {"symbol": {"symbolLayers": [{"enable": true, "miterLimit": 10, "capStyle": "Round", "width": 1, "color": {"values": [0, 0, 0, 0], "type": "CIMRGBColor"}, "joinStyle": "Round", "lineStyle3D": "Strip", "type": "CIMSolidStroke"}, {"enable": true, "color": {"values": [11, 121.99610900878906, 192, 100], "type": "CIMRGBColor"}, "type": "CIMSolidFill"}], "type": "CIMPolygonSymbol"}, "type": "CIMSymbolReference"}, "placement": "Unspecified", "polygon": {"rings": [[[0.9874572136477675, 10.5, null], [7.99232008788524, 10.5, null], [7.99232008788524, 10.062500000000002, null], [0.9874572136477675, 10.062500000000002, null], [0.9874572136477675, 10.5, null]]], "hasZ": true}, "blendingMode": "Alpha", "type": "CIMPolygonGraphic"}, "visible": true, "anchor": "BottomLeftCorner", "type": "CIMGraphicElement", "rotationCenter": {"y": 10.062500000000002, "x": 0.9874572136477675}}, {"name": "ReportTitle", "lockedAspectRatio": true, "graphic": {"placement": "Unspecified", "text": "Report Title", "type": "CIMTextGraphic", "shape": {"y": 10.41994855796273, "x": 1.084796287761}, "symbol": {"symbol": {"fontEncoding": "Unicode", "haloSize": 1, "textCase": "Normal", "kerning": true, "type": "CIMTextSymbol", "verticalAlignment": "Top", "textDirection": "LTR", "symbol3DProperties": {"scaleZ": 0, "rotationOrder3D": "XYZ", "scaleY": 0, "dominantSizeAxis3D": "Z", "type": "CIM3DSymbolProperties"}, "height": 18, "fontEffects": "Normal", "lineGapType": "ExtraLeading", "blockProgression": "TTB", "hinting": "Default", "wordSpacing": 100, "billboardMode3D": "FaceNearPlane", "verticalGlyphOrientation": "Right", "fontType": "Unspecified", "horizontalAlignment": "Left", "fontFamilyName": "Tahoma", "depth3D": 1, "fontStyleName": "Bold", "symbol": {"symbolLayers": [{"enable": true, "color": {"values": [255, 255, 255, 255], "type": "CIMRGBColor"}, "type": "CIMSolidFill"}], "type": "CIMPolygonSymbol"}, "extrapolateBaselines": true, "ligatures": true, "letterWidth": 100}, "type": "CIMSymbolReference"}, "blendingMode": "Alpha"}, "visible": true, "anchor": "TopLeftCorner", "type": "CIMGraphicElement", "rotationCenter": {"y": 10.41994855796273, "x": 1.084796287761}}], "metadataURI": "CIMPATH=Metadata/127c326aab3d1f0f715730e9018dc138.xml", "page": {"smallestRulerDivision": 0, "showGuides": true, "type": "CIMPage", "guides": [{"position": 0.5, "orientation": "Horizontal", "type": "CIMGuide"}, {"position": 10.5, "orientation": "Horizontal", "type": "CIMGuide"}, {"position": 0.5, "orientation": "Vertical", "type": "CIMGuide"}, {"position": 8, "orientation": "Vertical", "type": "CIMGuide"}, {"position": 0.75, "orientation": "Horizontal", "type": "CIMGuide"}, {"position": 10, "orientation": "Horizontal", "type": "CIMGuide"}, {"position": 10.0625, "orientation": "Horizontal", "type": "CIMGuide"}], "height": 11, "showRulers": true, "units": {"uwkid": 109008}, "width": 8.5}}, "type": "CIMLayoutDocument"}'
            pagx = self.write_pagx(overflow_layout_json, self.overflow_layout_name)
        else:
            base_name = os.path.splitext(os.path.basename(template))[0]
            self.overflow_layout_name = base_name + str(self.time_stamp)
            pagx = self.update_template(template, self.overflow_layout_name, self.overflow_element_names)
        self.overflow_pagx = pagx
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
            pagx = os.path.dirname(os.path.abspath(__file__)) + os.sep + name + '.pagx'
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
                os.remove(self.overflow_pagx)
                if table.full_overflow:
                    table.is_overflow = False
                table.rows = [[v.replace('\n','') for v in r] for r in table.rows]
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
                self.tables.insert(x, overflow_table)
            if table.has_buffer_rows:
                if overflow:
                    x += 1
                buffer_rows_table = Table(table.title + BUFFER_TITLE, table.buffer_rows, table.fields)
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

                    self.cur_y -= (table_header_background.elementHeight + MARGIN)
                start_y = self.cur_y 

                arcpy.AddMessage("Generating Table: " + table.title)     
                self.add_row_backgrounds(table)
                self.add_table_lines('vertLine', True, table)
                self.add_table_lines('horzLine', False, table)

                self.base_y = self.cur_y - MARGIN
                eh = self.place_holder.elementHeight
                esy = self.place_holder.elementPositionY
                self.remaining_height = eh - (esy - self.base_y)

                #first reset the x/y
                self.cur_x = self.place_holder.elementPositionX
                self.cur_y = start_y
                self.add_values(table)
        self.delete_elements()

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
        if self.layout_type == 'map':
            self.cur_elements['ReportTitle'].text = self.report_title
            self.cur_elements['Logo'].sourceImage = self.logo
            self.cur_elements['PageNumber'].text = self.page_num
            self.cur_elements['ReportTitleFooter'].text = self.report_title
            self.cur_elements['ReportSubTitle'].text = self.sub_title
            self.cur_elements['ReportType'].text = self.report_type
            #TODO doesn't seem like we can set the unit on the element
            if not self.map in ['', ' ', None]:
                maps = self.aprx.listMaps(self.map)
                map_frame = self.cur_elements['Map Map Frame']
                if len(maps) > 0:
                    user_map = maps[0]
                    ext = user_map.defaultCamera.getExtent()
                    map_frame.map = user_map
                    map_frame.camera.setExtent(ext)
                    
        else:
            self.cur_elements['ReportTitle'].text = self.report_title
            self.cur_elements['Logo'].sourceImage = self.logo
            self.cur_elements['PageNumber'].text = self.page_num
            self.cur_elements['ReportTitleFooter'].text = self.report_title

    def add_row_backgrounds(self, table):
        x = 0
        cur_x = self.cur_x
        cur_y = self.cur_y
        for adjust_value in table.row_heights:
            if x % 2 == 0:
                temp_row_background = table.row_background.clone("row_background_clone")
                temp_row_background.elementWidth = table.row_width
                if x == 0:                 
                    temp_row_background.elementHeight = table.row_heights[0]
                else:
                    temp_row_background.elementHeight = table.row_heights[x]        
                temp_row_background.elementPositionX = cur_x
                temp_row_background.elementPositionY = cur_y
                cur_y -= temp_row_background.elementHeight
            else:
                cur_y -= table.row_heights[x]
            x += 1  

    def add_table_lines(self, element_name, vert, table): 
        line = self.cur_elements[element_name].clone(element_name + "_clone")
        if vert:
            line.elementHeight = table.table_height
            collection = table.field_widths
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
        if not table.is_overflow:
            for f in table.fields:
                elm = field_name.clone("header_clone")
                elm.text = f
                elm.elementPositionX = self.cur_x + MARGIN
                elm.elementPositionY = self.cur_y - MARGIN
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
                elm.text = v
                elm.elementPositionX = self.cur_x + MARGIN
                elm.elementPositionY = self.cur_y - MARGIN
                self.cur_x += table.field_widths[x]
                x += 1
            new_row = True
            if table.is_overflow:
                self.cur_y -= float(table.row_heights[xx])
            xx += 1

    def delete_elements(self):
        for elm in self.cur_elements:
            if elm not in ['ReportTitle', 'ReportSubTitle', 'ReportType', 'Logo', 'PageNumber', 'ReportTitleFooter']:
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
            arcpy.AddMessage("Exporting report with modified name: " + self.pdf_path)
            pdf_doc = arcpy.mp.PDFDocumentCreate(self.pdf_path)
        except OSError:
            arcpy.AddMessage("Unable to export report: " + self.pdf_path)
            arcpy.AddMessage("Please ensure you have write access to: " + base_path)
            if len(self.pdf_paths) > 0:
                for pdf in self.pdf_paths:
                    os.remove(pdf)
                    self.pdf_paths.remove(pdf)
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
        self.pdf_path = os.path.join(folder, self.report_title + ".pdf")
        self.update_layouts()
        self.export_pdf()
        return self.pdf_path

    def write_pagx(self, json, name):
        #base_path = os.path.dirname(os.path.realpath(__file__))
        base_path = self.aprx.homeFolder
        pagx = base_path + os.sep + 'temp_' + name + '.pagx'
        with open(pagx, 'w') as write_file:
            write_file.writelines(json)
        return pagx
    
def main():
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

    tables = [t.strip("'") for t in tables.split(';')]

    report = Report(report_title, sub_title, logo, map, scale_unit, report_type, map_report_template, overflow_report_template)
   
    for table in tables:
        table_title = os.path.splitext(os.path.basename(table))[0]       
        desc = arcpy.Describe(table)
        if desc.dataType == "FeatureLayer" or desc.dataType == "TableView":
            fi = desc.fieldInfo
            test_fields = [f.name for f in desc.fields if f.type not in ['Geometry', 'OID'] and
                           fi.getVisible(fi.findFieldByName(f.name) == 'VISIBLE')]
        else:
            test_fields = [f.name for f in desc.fields if f.type not in ['Geometry', 'OID']]
        cur = arcpy.da.SearchCursor(table, test_fields)
        test_rows = [[str(v).replace('\n','') for v in r] for r in cur]
        report.add_table(table_title, test_rows, test_fields)

    pdf = report.generate_report(out_folder)
    os.startfile(pdf)

def test():
    report_title = "Generic Resource Assignment Analysis12345"
    map_report_template = None
    #map_report_template = r'C:\Solutions\EnvironmentalImpact\MapLayout.pagx'
    overflow_report_template = None
    #overflow_report_template = r'C:\Solutions\EnvironmentalImpact\OverflowLayout.pagx'
    report_title = "report title"   
    sub_title = "sub title"              
    logo = r"C:\Solutions\EnvironmentalImpact\matplotlibTests\EnvImpact\EnvImpact\Eagle Nesting Locations within Buffer.png"                     #optional report logo image
    tables = r"C:\Temp\New File Geodatabase.gdb\LongFieldNames".split(';')       
    map = None                      
    scale_unit = None               
    report_type = "report type"             

    out_folder = r"C:\temp"  

    table_title = "Eagle Nesting Locations within Buffer"
    test_rows2 = [["HL054", "Hillsborough dfsdfsdf sdf sdf sdfgdfg sdg fdsg sdfg sgdfgs", "256.3553791", 'BUFFER'],
                 ["HL018", "Hillsborough", "878.4109422", 'AOI'],
                 ["HL018", "Hillsborough", "878.4109423", 'AOI'],
                 ["HL018", "Hillsborough", "878.4109424", 'AOI'],
                 ["HL018", "Hillsborough", "878.4109425", 'AOI'],
                 ["HL018", "Hillsborough", "878.4109426", 'AOI'],
                 ["HL018", "Hillsborough dfs sdffsfsdddfsdddddddddddddddddddddddddd dffffsdffffff", "878.4109427", 'AOI'],
                 ["HL018", "Hillsborough aaaaaaaaaaaaaaaaaaaaddsdsd", "878.4109428", 'AOI'],
                 ["HL018", "Hillsborough", "878.4109429", 'AOI']]
    test_fields = ["NestID", "C", "D", 'RESULTTYPE']

    table_title2 = "Eagle Nesting Locations within Buffer2"

    #create the ReportData and add the table
    
    report = Report(report_title, sub_title, logo, map, "Meter", "report_type", map_report_template, overflow_report_template)
    #report.add_table(table_title, test_rows, test_fields)
    report.add_table(table_title2, test_rows2, test_fields)

    pdf = report.generate_report(out_folder)
    os.startfile(pdf)

def test2():
    report_title = "sd"             #required parameter for report title
    sub_title = "asd"                #optional parameter for sub-title
    logo = r"C:\Solutions\EnvironmentalImpact\matplotlibTests\EnvImpact\EnvImpact\Eagle Nesting Locations within Buffer.png"                     #optional report logo image
    tables = r"C:\Solutions\Cameo\data\SampleCAMEO_data.gdb\Facilities" #required multivalue parameter for input tables
    map = None                           #required parameter for the map 

    scale_unit = None               #required scale unit with default set
    report_type = "asd"             #optional report type eg. ...
    map_report_template = r"C:\Users\john4818\Documents\ArcGIS\Projects\EnvImpact1.2\ModifiedMapThingy.pagx"      #optional parameter for path to new pagX files
    overflow_report_template = r"C:\Users\john4818\Documents\ArcGIS\Projects\EnvImpact1.2\OverflowThingy.pagx" #optional parameter for path to new pagX files
    out_folder = r"C:\Solutions\EnvironmentalImpact\New folder"              #folder that will contain the final output report

    tables = [t.strip("'") for t in tables.split(';')]

    report = Report(report_title, sub_title, logo, map, scale_unit, report_type, map_report_template, overflow_report_template)
   
    for table in tables:
        table_title = os.path.splitext(os.path.basename(table))[0]       
        desc = arcpy.Describe(table)
        if desc.dataType == "FeatureLayer" or desc.dataType == "TableView":
            fi = desc.fieldInfo
            test_fields = [f.name for f in desc.fields if f.type not in ['Geometry', 'OID'] and
                           fi.getVisible(fi.findFieldByName(f.name) == 'VISIBLE')]
        else:
            test_fields = [f.name for f in desc.fields if f.type not in ['Geometry', 'OID']]
        cur = arcpy.da.SearchCursor(table, test_fields)
        test_rows = [[str(v) for v in r] for r in cur]
        report.add_table(table_title, test_rows, test_fields)

    pdf = report.generate_report(out_folder)
    os.startfile(pdf)

if __name__ == '__main__':
    main()
    #test()
    #test2()


