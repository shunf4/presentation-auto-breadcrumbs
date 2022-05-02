import uno
import typing
from com.sun.star.awt import Size
from com.sun.star.awt import Point
from com.sun.star.beans import PropertyValue

# bcx millimeter
BREADCRUMB_X = 0
# bcy millimeter
BREADCRUMB_Y = 0
# delimit (string) (with parentheses) / nodelimit
BREADCRUMB_DELIMITER = " > "
BREADCRUMB_STYLE_NAME = "Breadcrumb (Auto-generated)"
TOC_STYLE_NAME = "TOC (Auto-generated)"
# toccolorina
TOC_COLOR_INACTIVE = "CFCFCF"
# toccolora
TOC_COLOR_ACTIVE = "111111"

class TocEntry(object):
    def __init__(self, text):
        self.text = text
        self.shape = None
        self.children = []

    def __str__(self):
        return "TocEntry<" + self.text + "> [ " + ", ".join(list(map(lambda x: x.__str__(), self.children))) + " ]"

    def __repr__(self):
        return self.__str__()

def insert_child_and_switch_to(toc_list_stack: typing.List[TocEntry], text: str):
    new_entry = TocEntry(text)
    toc_list_stack[-1].children.append(new_entry)
    toc_list_stack.append(new_entry)

class IsFirstLine(object):
    def __init__(self, value: bool):
        self.value = value

    def set_value(self, value: bool):
        self.value = value

    def __bool__(self):
        return self.value

def do_recurse_write_toc_tree(depth: int, is_first_line: IsFirstLine, will_stress: bool, toc_root: TocEntry, curr_toc_entry: TocEntry, target_toc_entry: TocEntry):
    if curr_toc_entry is not toc_root:
        if not is_first_line:
            target_toc_entry.shape.finishParagraph([])
        else:
            is_first_line.set_value(False)

        para_props = [
            PropertyValue(Name = "NumberingLevel", Value = depth - 1)
        ]
        if will_stress:
            if curr_toc_entry is target_toc_entry:
                if TOC_COLOR_ACTIVE is not None and TOC_COLOR_ACTIVE != "":
                    para_props.append(PropertyValue(Name = "CharColor", Value = int(TOC_COLOR_ACTIVE, 16)))
            else:
                if TOC_COLOR_INACTIVE is not None and TOC_COLOR_INACTIVE != "":
                    para_props.append(PropertyValue(Name = "CharColor", Value = int(TOC_COLOR_INACTIVE, 16)))
                
        target_toc_entry.shape.appendTextPortion(curr_toc_entry.text, para_props)

    for child_entry in curr_toc_entry.children:
        do_recurse_write_toc_tree(depth + 1, is_first_line, will_stress, toc_root, child_entry, target_toc_entry)

def recurse_write_toc_tree(toc_root: TocEntry, target_toc_entry: TocEntry):
    if target_toc_entry.shape is None:
        return
        
    target_toc_entry.shape.setString("")
    do_recurse_write_toc_tree(0, IsFirstLine(True), target_toc_entry is not toc_root, toc_root, toc_root, target_toc_entry)


def do_recurse_toc_entry(depth: int, toc_root: TocEntry, curr_toc_entry: TocEntry):
    print(("  " * depth) + curr_toc_entry.text, end='')
    if curr_toc_entry.shape is not None:
        recurse_write_toc_tree(toc_root, curr_toc_entry)
        print(" (has TOC shape, written)", end='')
    print("")

    for child_entry in curr_toc_entry.children:
        do_recurse_toc_entry(depth + 1, toc_root, child_entry)

def recurse_toc_entry(toc_root: TocEntry):
    do_recurse_toc_entry(0, toc_root, toc_root)

def automatic_breadcrumbs():
    global BREADCRUMB_X
    global BREADCRUMB_Y
    global BREADCRUMB_DELIMITER
    global BREADCRUMB_STYLE_NAME
    global TOC_STYLE_NAME
    global TOC_COLOR_INACTIVE
    global TOC_COLOR_ACTIVE

    doc = XSCRIPTCONTEXT.getDocument()
    ctx = XSCRIPTCONTEXT.getComponentContext()
    ctl = doc.getCurrentController()
    sm = ctx.ServiceManager
    pages = doc.DrawPages
    tdm = ctx.getByName("/singletons/com.sun.star.reflection.theTypeDescriptionManager")
    tha_enum = tdm.getByHierarchicalName("com.sun.star.drawing.TextHorizontalAdjust")
    tha_dict = {name: value for name, value in zip(tha_enum.getEnumNames(), tha_enum.getEnumValues())}
    tva_enum = tdm.getByHierarchicalName("com.sun.star.drawing.TextVerticalAdjust")
    tva_dict = {name: value for name, value in zip(tva_enum.getEnumNames(), tva_enum.getEnumValues())}

    # breadcrumbs stack
    bc_stack = []

    toc_root = TocEntry("<Root>")
    toc_list_stack: typing.List[TocEntry] = []
    toc_list_stack.append(toc_root)

    graph_styles = doc.StyleFamilies.getByName("graphics")
    default_graph_style = graph_styles.getByName("standard")

    if graph_styles.hasByName(BREADCRUMB_STYLE_NAME):
        bc_graph_style = graph_styles.getByName(BREADCRUMB_STYLE_NAME)
    else:
        bc_graph_style = graph_styles.createInstance()
        graph_styles.insertByName(BREADCRUMB_STYLE_NAME, bc_graph_style)
        bc_graph_style.setParentStyle("standard")
        bc_graph_style.TextHorizontalAdjust = tha_dict["LEFT"]
        bc_graph_style.TextVerticalAdjust = tva_dict["TOP"]

    if graph_styles.hasByName(TOC_STYLE_NAME):
        toc_graph_style = graph_styles.getByName(TOC_STYLE_NAME)
    else:
        toc_graph_style = graph_styles.createInstance()
        graph_styles.insertByName(TOC_STYLE_NAME, toc_graph_style)
        toc_graph_style.setParentStyle("standard")
        

    for page in pages:
        for shape in page:
            if not shape.supportsService("com.sun.star.drawing.TextShape"):
                continue

    for page in pages:
        largest_shape = None
        largest_shape_area = 0
        top_shape = None
        top_shape_y = 999999
        toc_shape = None

        is_toc = False
        should_push_title = False
        push_extra_list = []
        should_hide_bc = False
        pop_count = 0

        for shape in page:
            if not shape.supportsService("com.sun.star.drawing.TextShape"):
                continue

            s: str = shape.getString()
            s = s.strip()
            if s == "#toc":
                is_toc = True
            elif s == "#push":
                should_push_title = True
            elif s.startswith("#push "):
                push_strs = s[len("#push "):].split("|")
                push_extra_list += push_strs
            elif s == "#pop":
                pop_count += 1
            elif s.startswith("#popto "):
                pop_count = len(bc_stack) - int(s[len("#popto "):])
            elif s.startswith("#pop "):
                pop_count += int(s[len("#pop "):])

            elif s == "#poppush":
                pop_count += 1
                should_push_title = True
            elif s.startswith("#poppush "):
                pop_count += 1
                push_strs = s[len("#poppush "):].split("|")
                push_extra_list += push_strs

            elif s == "#poppoppush":
                pop_count += 2
                should_push_title = True
            elif s.startswith("#poppoppush "):
                pop_count += 2
                push_strs = s[len("#poppoppush "):].split("|")
                push_extra_list += push_strs

            elif s == "#poppoppoppush":
                pop_count += 3
                should_push_title = True
            elif s.startswith("#poppoppoppush "):
                pop_count += 3
                push_strs = s[len("#poppoppoppush "):].split("|")
                push_extra_list += push_strs

            elif s == "#hidebc":
                should_hide_bc = True
            elif s.startswith("#bcx "):
                BREADCRUMB_X = int(s[len("#bcx "):])
            elif s.startswith("#bcy "):
                BREADCRUMB_Y = int(s[len("#bcy "):])
            elif s == "#nodelimit":
                BREADCRUMB_DELIMITER = ""
            elif s.startswith("#delimit ("):
                BREADCRUMB_DELIMITER = s[len("#delimit ("):-1]
            elif s.startswith("#toccolora "):
                TOC_COLOR_ACTIVE = s[len("#toccolora "):]
            elif s.startswith("#toccolorina "):
                TOC_COLOR_INACTIVE = s[len("#toccolorina "):]
            elif shape.Style.Name == TOC_STYLE_NAME:
                toc_shape = shape
            elif shape.Style.Name == BREADCRUMB_STYLE_NAME:
                bc_shape = shape
            else:
                area = shape.Size.Width * shape.Size.Height
                if area > largest_shape_area:
                    largest_shape = shape
                    largest_shape_area = area

                if shape.Position.X >= 0 and shape.Position.Y >= 0:
                    if shape.Position.Y < top_shape_y:
                        top_shape = shape
                        top_shape_y = shape.Position.Y
                    
        if pop_count < 0 or pop_count > len(bc_stack):
            raise ValueError("pop too much")

        if toc_shape is None:
            toc_shape = largest_shape
        title_shape = top_shape

        for i in range(pop_count):
            bc_stack.pop()
            toc_list_stack.pop()

        if should_push_title:
            title_text = title_shape.getString().strip()
            bc_stack.append(title_text)
            insert_child_and_switch_to(toc_list_stack, title_text)

        bc_stack += push_extra_list
        for push_extra in push_extra_list:
            insert_child_and_switch_to(toc_list_stack, push_extra)

        if should_hide_bc:
            if bc_shape is not None:
                page.remove(bc_shape)
        else:
            if len(bc_stack) > 0:
                if bc_shape is None:
                    bc_shape = doc.createInstance("com.sun.star.drawing.TextShape")
                    page.add(bc_shape)
                bc_shape.TextAutoGrowHeight = True
                bc_shape.TextAutoGrowWidth = True
                bc_shape.setString(BREADCRUMB_DELIMITER.join(bc_stack))
                bc_shape.setPosition(Point(BREADCRUMB_X, BREADCRUMB_Y))
                bc_shape.Style = bc_graph_style

        if is_toc:
            toc_list_stack[-1].shape = toc_shape
            toc_shape.setString("<TOC>")
            toc_shape.Style = toc_graph_style

    recurse_toc_entry(toc_root)

g_exportedScripts = automatic_breadcrumbs,

if __name__ == '__main__':
    print()
    print()
    print()

    from IDE_utils import Runner, XSCRIPTCONTEXT

    runner = {
        "C:/Program Files/LibreOffice/program/swriter.exe": [
            "--accept=\"socket,host=localhost,port=2017;urp;\""
        ],
    }

    with Runner(soffice=None) as jesse_owens:  # Start/Stop, Connect/Adapt
        automatic_breadcrumbs()  # Run
    
    print()
    print()
    print()
