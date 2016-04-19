import arcpy
import os, sys

#Elements should be anchored at upper left and have the appropriate names

#filter

#TODO logo, title, subtitle among others should be exposed

# when text will overflow need to detrmine where to split and calculate the row_height accordingly
 

#list of core element names that should be in the project
# these elements will be cloned and re-used for the generated tables
#If these names are changed in the mxd update this list as well
# as any place they are hardcoded in the following
ELEMENTS = ['horzLine', 'vertLine', 'cellText', 
            'headerText', 'analysisType', 'tableAreaHeader', 
            'tableAreaHeaderText', 'evenRowBackground', 'contentDisplayArea']

# need to look at export/import of Overflow Layout so we only need 1 but could support 1-n
MAP_LAYOUT_NAME = 'MapLayout'
OVERFLOW_LAYOUT_NAME = 'OverflowLayout'

MARGIN = .025

class Table:
    def __init__(self, title, rows, fields):
        self.rows = rows
        self.fields = fields
        self.title = title
        self.row_count = len(self.rows)

        self.column_widths = []
        self.row_heights = []
        self.header_element = None
        self.cell_elelemt = None
        self.header_bar = None
        self.header_bar_text = None
        self.row_background = None
        self.content_display = None
        self.overflow_rows = None
        self.is_overflow = False

    def calc_column_widths(self):
        field_name_lengths = []
        for f in self.fields:
            self.header_element.text = f
            field_name_lengths.append(self.header_element.elementWidth + (MARGIN * 2))      
     
        update = False
        for row in self.rows:
            x = 0
            for v in row:
                val = field_name_lengths[x]
                if len(self.column_widths) == len(field_name_lengths):
                    val = self.column_widths[x]
                    update = True
                self.cell_element.text = v
                potential_val = self.cell_element.elementWidth + (MARGIN * 2)
                if potential_val > val:
                    val = potential_val
                if update:
                    self.column_widths[x] = val
                else:
                    self.column_widths.append(val)
                x+=1
        self.row_width = 0
        for w in self.column_widths:
            self.row_width += w

    def calc_row_heights(self):
        row_heights = [self.row_height] * len(self.rows)
        if not self.is_overflow:
            row_heights.insert(0, self.header_height)
        self.row_heights = row_heights

    def calc_table_height(self, remaining_height):
        h_height = self.header_element.elementHeight
        self.header_height = h_height + (h_height * (MARGIN * 2))

        c_height = self.cell_element.elementHeight
        self.row_height = c_height + (c_height * (MARGIN * 2)) 

        table_height = self.row_count * self.row_height
        if not self.is_overflow:
            table_height += self.header_height
        if remaining_height == None:
            remaining_height = self.content_display.elementHeight
        remaining_height -= (self.header_bar.elementHeight + MARGIN)
        if table_height > remaining_height:
            num_rows = int(remaining_height / self.row_height)
            self.overflow_rows = self.rows[num_rows - 1:]
            self.rows = self.rows[:num_rows - 1]
            self.row_count = len(self.rows)
            self.table_height = self.row_height * num_rows
            return True
        else:
            self.table_height = table_height
    
    def init_elements(self, elements, remaining_height):
        #Get Map Elelments
        self.header_element = elements['headerText']
        self.cell_element = elements['cellText']
        self.header_bar = elements['tableAreaHeader']
        self.header_bar_text = elements['tableAreaHeaderText']
        self.row_background = elements['evenRowBackground']
        self.content_display = elements['contentDisplayArea']

        #Calculate the row/column width/heights
        overflow = self.calc_table_height(remaining_height)
        self.calc_column_widths()
        self.calc_row_heights()
        return overflow

class ReportData:
    def __init__(self, report_title, aprx, map_element_names, map_layout_name, overflow_element_names, overflow_layout_name):
        self.tables = []
        self.pdf_paths = []
        self.pdfs = []
        self.idx = 0
        self.cur_x = 0
        self.cur_y = 0
        self.base_y = None
        self.base_x = None
        self.remaining_height = None
        self.overflow_row = None
        self.overflow = False
        self.place_holder = None

        self.report_title = report_title     
        self.aprx = aprx
        self.map_element_names = map_element_names
        self.map_layout_name = map_layout_name
        self.overflow_element_names = overflow_element_names
        self.overflow_layout_name = overflow_layout_name

        self.elements = {}
        self.elements[map_layout_name] = self.find_elements('map')
        self.elements[overflow_layout_name] = self.find_elements('overflow')

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

    def find_elements(self, type):
        #find elements from layout based on name and returns name/object dict
        # elements ['elmName1', 'elmName2']
        # returns {'elmName1': elm1, 'elmName2': elm2}
        # element type is based on naming conventions of elements
        layout_name = self.map_layout_name if type == 'map' else self.overflow_layout_name
        element_names = self.map_element_names if type == 'map' else self.overflow_element_names
        layouts = self.aprx.listLayouts(layout_name)
        if len(layouts) > 0:
            out_elements = {}
            for elm_name in element_names:
                elms = layouts[0].listElements(wildcard = elm_name)
                if len(elms) > 0:
                    print("Found: " + elm_name)
                    elm = elms[0]
                    out_elements[elm.name] = elm
                else:
                    print("Cannot find: " + elm_name)
            return out_elements
        else:
            print("Cannot locate Layout: " + layout_name)

    def update_layouts(self):
        self.set_layout(self.map_layout_name)
        x = 0
        a = False
        for table in self.tables:
            if table.is_overflow:
                if not a:
                    self.delete_elements()
                    self.set_layout(self.overflow_layout_name)
                    a = True
            overflow = table.init_elements(self.cur_elements, self.remaining_height)
            x += 1
            if overflow:
                overflow_table = Table(table.title, table.overflow_rows, table.fields)
                overflow_table.is_overflow = True
                self.tables.insert(x, overflow_table)

            if table.row_width > self.place_holder.elementWidth:
                print("TODO Need to adjust defined font internally so things will fit")

            header_bar = table.header_bar
            if self.base_y == None:
                if table.is_overflow:
                    height = self.place_holder.elementHeight
                else:
                    height = self.place_holder.elementHeight - header_bar.elementHeight
                max_num_rows = int(float(height) / float(table.row_height))
            else:
                height = self.remaining_height - header_bar.elementHeight
                max_num_rows = int(float(height) / float(table.row_height))

            if table.row_count > max_num_rows:
                print('Will need overflow...this should not happen any more')

            if not table.is_overflow:
                header_bar = table.header_bar.clone('header_bar_clone')
                header_bar_text = table.header_bar_text.clone('header_bar_text_clone')
                header_bar_text.text = table.title
                if self.base_y == None:
                    header_bar.elementPositionX = self.cur_x
                    header_bar.elementPositionY = self.cur_y
                               
                    header_bar_text.elementPositionX = self.cur_x + MARGIN
                    header_bar_text.elementPositionY = self.cur_y    
                else:                  
                    header_bar.elementPositionX = self.base_x
                    header_bar.elementPositionY = self.base_y

                    header_bar_text.elementPositionX = self.base_x + MARGIN
                    header_bar_text.elementPositionY = self.base_y
                    self.cur_y = self.base_y
                    self.cur_x = self.base_x

                self.cur_y -= (header_bar.elementHeight + MARGIN)
            start_y = self.cur_y 

            print("Generating Table: " + table.title)     
            self.add_row_backgrounds(table)
            self.add_table_lines('vertLine', True, table)
            self.add_table_lines('horzLine', False, table)

            self.base_y = self.cur_y - MARGIN
            if self.remaining_height == None:
                eh = self.place_holder.elementHeight
                if table.is_overflow:
                    self.remaining_height = eh - (eh - self.base_y)
                else:
                    self.remaining_height = eh - self.base_y
            else:
                self.remaining_height -= self.base_y

            #first reset the x/y
            self.cur_x = self.place_holder.elementPositionX
            self.cur_y = start_y
            self.add_values(table)
        self.delete_elements()

    def set_layout(self, layout_name):
        layouts = self.aprx.listLayouts(layout_name)
        if len(layouts) > 0:
            self.cur_elements = self.elements[layout_name]
            self.place_holder = self.cur_elements['contentDisplayArea']
            #expecting this to be the upper left
            self.cur_x = self.place_holder.elementPositionX
            self.cur_y = self.place_holder.elementPositionY
            self.base_x = self.cur_x
            self.base_y = None
            self.remaining_height = None
            self.layout = layouts[0]
            print("Layout Set: " + layout_name)
            self.add_pdf()
        else:
            print("Cannot find overflow layout template.")
            sys.exit()
            self.remaining_height = None

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
                    temp_row_background.elementHeight = table.row_height         
                temp_row_background.elementPositionX = cur_x
                temp_row_background.elementPositionY = cur_y
                cur_y -= temp_row_background.elementHeight
            else:
                cur_y -= table.row_height
            x += 1  

    def add_table_lines(self, element_name, vert, table): 
        line = self.cur_elements[element_name].clone(element_name + "_clone")
        if vert:
            line.elementHeight = table.table_height
            collection = table.column_widths
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
        cell_element = table.cell_element
        header_element = table.header_element
        base_x = self.cur_x
        x = 0
        if not table.is_overflow:
            for f in table.fields:
                elm = header_element.clone("header_clone")
                elm.text = f
                elm.elementPositionX = self.cur_x + (MARGIN / 2)
                elm.elementPositionY = self.cur_y - (MARGIN / 4)
                self.cur_x += table.column_widths[x]
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
                        self.cur_y -= table.row_height
                    new_row = False
                elm = cell_element.clone("cell_clone")
                elm.text = v
                elm.elementPositionX = self.cur_x + (MARGIN / 2)
                elm.elementPositionY = self.cur_y - (MARGIN / 4)
                self.cur_x += table.column_widths[x]
                x += 1
            new_row = True
            if table.is_overflow:
                self.cur_y -= table.row_height
            xx += 1

    def delete_elements(self):
        for elm in self.cur_elements:
            self.cur_elements[elm].delete()

    def add_pdf(self):
        self.pdfs.append({'title': self.report_title, 'layout': self.layout})

    def export_pdf(self):
        x = 0
        for pdf in self.pdfs:
            unique = str(x) if len(self.pdfs) > 1 else ''
            pdf_path = os.path.join(self.path, pdf['title'] + unique + ".pdf")
            pdf['layout'].exportToPDF(pdf_path)
            self.pdf_paths.append(pdf_path)
            x += 1
        if len(self.pdfs) > 1:
            self.append_pages("", "", "", "", "", "")

    def append_pages(self, pdf_title, pdf_author, pdf_subject, pdf_keywords, pdf_open_view, pdf_layout):
        if os.path.isfile(self.pdf_path):
            os.remove(self.pdf_path)
        pdf_doc = arcpy.mp.PDFDocumentCreate(self.pdf_path)

        for pdf in self.pdf_paths:
            pdf_doc.appendPages(pdf)
            os.remove(pdf)

        pdf_doc.saveAndClose()

    def generate_report(self, folder):
        self.path = folder
        self.pdf_path = os.path.join(folder, self.report_title + ".pdf")
        self.update_layouts()
        self.export_pdf()
        return self.pdf_path

def main():
    ####Setup report properties..these will come from the previous tools####
    tables = arcpy.GetParameterAsText(0).split(';')
    report_title = arcpy.GetParameterAsText(1)
    aprx = arcpy.mp.ArcGISProject(arcpy.GetParameterAsText(2))
    out_folder = arcpy.GetParameterAsText(3)
    
    report = ReportData(report_title, aprx, ELEMENTS, MAP_LAYOUT_NAME, ELEMENTS, OVERFLOW_LAYOUT_NAME)
   
    for table in tables:
        table_title = os.path.splitext(os.path.basename(table))[0]
        test_fields = [f.name for f in arcpy.Describe(table).fields]
        cur = arcpy.da.SearchCursor(table, test_fields)
        test_rows = [[v for v in r] for r in cur]
        report.add_table(table_title, test_rows, test_fields)

    pdf = report.generate_report(r"C:\temp")
    print("PDF Generated: " + pdf)
    os.startfile(pdf)

if __name__ == '__main__':
    main()