# -*- coding: utf-8 -*-

import clr
import System
from System.Collections.Generic import List
from System.Windows import *
from System.Windows.Forms import OpenFileDialog, DialogResult

# Add Revit API references
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import ObjectType

# Import pyRevit forms for easy UI handling
from pyrevit import forms
from pyrevit import revit, script

# ------------------------------------------------------------------------------
# Setup & Globals
# ------------------------------------------------------------------------------
doc = revit.doc
uidoc = revit.uidoc
app = doc.Application

# --- CONFIGURATION ---
SWATCH_FAMILY_NAME = "material_swatch"
SWATCH_MATERIAL_PARAM = "Swatch Material"
TAG_FAMILY_NAME = "Material Tag"
SWATCH_WIDTH_FEET = 2.0 / 12.0
SWATCH_HEIGHT_FEET = 2.0 / 12.0
SPACING_FEET = 1.0 / 12.0
DESIGN_OPTION_SET_NAME = "Material Boards"

# ------------------------------------------------------------------------------
# Helper Classes & Functions
# ------------------------------------------------------------------------------
class MaterialSelection:
    """A wrapper class for materials to be used in the UI list."""
    def __init__(self, material):
        self.Name = material.Name
        self.IsSelected = False
        self.Element = material

def get_all_materials():
    """Returns a list of all materials in the project."""
    materials = FilteredElementCollector(doc).OfClass(Material).ToElements()
    return sorted([MaterialSelection(m) for m in materials], key=lambda x: x.Name)

def get_family_symbol(family_name, symbol_name=None):
    """Finds and returns a family symbol."""
    collector = FilteredElementCollector(doc).OfClass(FamilySymbol)
    symbols = [s for s in collector if s.Family.Name == family_name]
    if symbols:
        return symbols[0]
    else:
        forms.alert("Family '{}' not found. Please load it.".format(family_name), exitscript=True)

def get_or_create_design_option_set(name):
    """Gets or creates a Design Option Set by name."""
    sets = DesignOptionSet.GetDesignOptionSets(doc)
    for s in sets:
        if s.Name == name:
            return s
    # If not found, create it
    return DesignOptionSet.Create(doc, name)

# ------------------------------------------------------------------------------
# Main UI and Logic Class
# ------------------------------------------------------------------------------
class MaterialBoardGenerator(forms.WPFWindow):
    def __init__(self, xaml_file_name):
        forms.WPFWindow.__init__(self, xaml_file_name)
        self.SheetTitle = "MATERIAL BOARD"
        self.SheetNumber = "A-901"
        self.TitleBlockName = "Select Title Block..."
        self.titleblock_id = ElementId.InvalidElementId
        self.Materials = get_all_materials()

    @property
    def selected_materials(self):
        return [m.Element for m in self.Materials if m.IsSelected]

    def select_titleblock_click(self, sender, args):
        """Handle the title block selection button click."""
        collector = FilteredElementCollector(doc).OfClass(FamilySymbol).OfCategory(BuiltInCategory.OST_TitleBlocks)
        titleblocks = {t.Family.Name + " : " + t.Name: t.Id for t in collector}
        selected_key = forms.SelectFromList.show(titleblocks.keys(), title='Select Title Block', multiselect=False)
        if selected_key:
            self.titleblock_id = titleblocks[selected_key]
            self.TitleBlockName = selected_key
            self.get_control('TitleBlockName').Content = selected_key

    def generate_click(self, sender, args):
        """Main logic to generate the material board."""
        # 1. Input Validation
        if not self.SheetTitle or not self.SheetNumber:
            forms.alert("Please provide a Sheet Title and Number.", exitscript=True)
        if self.titleblock_id == ElementId.InvalidElementId:
            forms.alert("Please select a Title Block.", exitscript=True)
        if not self.selected_materials:
            forms.alert("Please select at least one material.", exitscript=True)

        self.Hide()

        # 2. Get required families
        swatch_symbol = get_family_symbol(SWATCH_FAMILY_NAME)
        tag_symbol = get_family_symbol(TAG_FAMILY_NAME)

        # 3. Use a Transaction Group
        with revit.TransactionGroup('Generate Material Board'):
            with revit.Transaction('Create Design Option, 3D View, and Swatches'):
                # 4. ⭐ Create a Design Option to host the swatches
                dos = get_or_create_design_option_set(DESIGN_OPTION_SET_NAME)
                option_name = "Board - {} ({})".format(self.SheetNumber, self.SheetTitle)
                design_option = DesignOption.Create(doc, option_name, dos.Id)
                
                # 5. ⭐ Create a new 3D view
                view_family_type = FilteredElementCollector(doc).OfClass(ViewFamilyType).First(lambda v: v.ViewFamily == ViewFamily.ThreeDimensional)
                view_3d = View3D.CreateIsometric(doc, view_family_type.Id)
                view_3d.Name = "Material Board - " + self.SheetNumber
                
                # Set view to show our specific design option
                view_3d.DesignOptionId = design_option.Id
                
                # 6. Place swatches in a grid (within the design option)
                placed_swatches = []
                paper_width_feet = 15.0 / 12.0
                cols = int(paper_width_feet / (SWATCH_WIDTH_FEET + SPACING_FEET)) or 1
                
                for i, material in enumerate(self.selected_materials):
                    col_index = i % cols
                    row_index = i // cols
                    x = col_index * (SWATCH_WIDTH_FEET + SPACING_FEET)
                    y = -row_index * (SWATCH_HEIGHT_FEET + SPACING_FEET)
                    
                    # Create swatch instance in the active design option
                    pt = XYZ(x, y, 0)
                    # We pass the design_option.Id to NewFamilyInstance
                    swatch_instance = doc.Create.NewFamilyInstance(pt, swatch_symbol, doc.ActiveView)
                    # Manually assign the swatch to the correct design option
                    swatch_instance.DesignOption = design_option

                    # Set material parameter
                    mat_param = swatch_instance.LookupParameter(SWATCH_MATERIAL_PARAM)
                    if mat_param:
                        mat_param.Set(material.Id)
                    
                    placed_swatches.append(swatch_instance)

                # 7. ⭐ Configure and lock the 3D view
                view_3d.SetOrientation(ViewOrientation3D(XYZ(0, 0, 0), XYZ(0, -1, 0), XYZ(0, 0, 1))) # Look from front
                view_3d.get_BoundingBox(None).Max += XYZ(paper_width_feet, 8.0/12.0, 0) # Attempt to fit crop
                view_3d.CropBoxActive = True
                view_3d.CropBoxVisible = False
                view_3d.DisplayStyle = DisplayStyle.Realistic
                view_3d.SaveOrientationAndLock()

            with revit.Transaction('Create Sheet and Place Viewport'):
                # 8. Create the Sheet
                new_sheet = ViewSheet.Create(doc, self.titleblock_id)
                new_sheet.Name = self.SheetTitle
                new_sheet.SheetNumber = self.SheetNumber

                # 9. Prompt user to pick a point
                try:
                    forms.alert("Please select the bottom-left corner for the material schedule on the new sheet.", title="Place Schedule")
                    revit.active_view = new_sheet
                    pick_point = uidoc.Selection.PickPoint("Select viewport bottom-left corner")
                except Exception as e:
                    if 'Operation canceled' in str(e):
                        forms.alert("Operation cancelled by user.", exitscript=True)
                    else:
                        raise e

                # 10. Place the 3D View as a Viewport
                viewport_center = pick_point + XYZ(view_3d.Outline.Width / 2.0, view_3d.Outline.Height / 2.0, 0)
                viewport = Viewport.Create(doc, new_sheet.Id, view_3d.Id, viewport_center)

            # Tagging is disabled for 3D views. This is a Revit limitation.
            # with revit.Transaction('Place Material Tags'):
            #     # Revit API does not support creating tags for elements in a 3D view on a sheet.
            #     # You would need to add text notes or use a workaround.
            #     pass

        self.Close()
        forms.alert("Material Board '{}' created successfully! Note: Automatic tagging is not supported in 3D views; text notes must be added manually.".format(self.SheetNumber), title="Success")


# ------------------------------------------------------------------------------
# Script Execution
# ------------------------------------------------------------------------------
if __name__ == '__main__':
    ui = MaterialBoardGenerator('ui.xaml')
    ui.ShowDialog()
