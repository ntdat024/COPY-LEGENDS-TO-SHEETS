#region library
import clr 
import os
import sys

clr.AddReference("System")
clr.AddReference("System.Data")
clr.AddReference("RevitServices")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference("System.Windows.Forms")

import math
import System
import RevitServices
import Autodesk
import Autodesk.Revit
import Autodesk.Revit.DB

from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import *
from Autodesk.Revit.DB.Mechanical import *

from System.Collections.Generic import *
from System.Windows import MessageBox
from System.Windows.Controls import Button, CheckBox, ListBox
from System.IO import FileStream, FileMode, FileAccess
from System.Windows.Markup import XamlReader

#endregion

#region revit infor
# Get the directory path of the script.py & the Window.xaml
dir_path = os.path.dirname(os.path.realpath(__file__))
xaml_file_path = os.path.join(dir_path, "Window.xaml")

#Get UIDocument, Document, UIApplication, Application
uidoc = __revit__.ActiveUIDocument
uiapp = UIApplication(uidoc.Document.Application)
app = uiapp.Application
doc = uidoc.Document
activeView = doc.ActiveView

all_viewPorts = FilteredElementCollector(doc, activeView.Id).OfCategory(BuiltInCategory.OST_Viewports).WhereElementIsNotElementType().ToElements()
all_sheets = FilteredElementCollector(doc).OfClass(ViewSheet).WhereElementIsNotElementType().ToElements()
all_viewSets = FilteredElementCollector(doc).OfClass(ViewSheetSet).ToElements()

#endregion


#region method
class Utils:
    def __init__(self):
        pass
    
    def get_all_legend_names(self):
        legends = []
        for view_port in all_viewPorts:
            view = doc.GetElement(view_port.ViewId)
            if view.ViewType == ViewType.Legend:
                legends.append(view.Name)
        
        legends.sort()
        return legends
    
    def get_sheet_full_name(self, sheet):
        return sheet.SheetNumber + " - " + sheet.Name
    
    def get_all_sheetSet (self):
        view_set = ["<all>"]
        for vs in all_viewSets:
            if vs.Name != "" and vs.Views.Size > 0 :
                view_set.append(vs.Name)
        return view_set
    
    def get_viewset_by_name (self, name):
        for vs in all_viewSets:
            if name == vs.Name:
                return vs
        return None
    
    def get_viewsheet_in_viewset (self, view_set):
        sheets = []
        try:
            for v in view_set:
                if isinstance(v, ViewSheet):
                    sheets.append(v)
        except:
            pass
        return sheets


    def get_all_sheet_in_model (self):
        sheets = sorted(list(all_sheets), key=lambda vs: vs.SheetNumber)
        sheet_names = []
        for vs in sheets:
            full_name = self.get_sheet_full_name(vs)
            sheet_names.append(full_name)
        return sheet_names
    
    def get_sheet_by_view_set (self, viewset_name):

        if viewset_name == "<all>":
            return self.get_all_sheet_in_model()
        else:
            view_set = self.get_viewset_by_name(viewset_name)

            sheet_in_view_set = []
            if view_set is not None:
                for sheet in view_set.Views:
                    if isinstance(sheet, ViewSheet):
                        sheet_in_view_set.append(sheet)
                        
            
            sheet_names = []
            sheet_in_view_set = sorted(sheet_in_view_set, key=lambda vs: vs.SheetNumber)
            for sheet in sheet_in_view_set:
                full_name = self.get_sheet_full_name(sheet)
                sheet_names.append(full_name)

            return sheet_names


    # end def
    def get_all_sheet_names (self):
        sheets = sorted(list(all_sheets), key=lambda vs: vs.SheetNumber)
        sheet_names = []
        for vs in sheets:
            full_name = self.get_sheet_full_name(vs)
            sheet_names.append(full_name)
        return sheet_names
    

    def get_sheet_element_by_name (self, list_sheet_names):
        sheet_elements = []
        for vs in all_sheets:
            full_name = self.get_sheet_full_name(vs)
            if list(list_sheet_names).__contains__(full_name):
                sheet_elements.append(vs)
        return sheet_elements
    
    def get_legend_viewport_by_name (self, list_legend_names):
        viewport_inSheet = []
        for view_port in all_viewPorts:
            view = doc.GetElement(view_port.ViewId)
            if view.ViewType == ViewType.Legend and list(list_legend_names).__contains__(view.Name):
                viewport_inSheet.append(view_port)
        return viewport_inSheet
    
    def copy_legends_to_sheet (self, list_legend_names, list_sheet_names):
        try:
            t = Transaction(doc," ")
            t.Start()

            for sheet in self.get_sheet_element_by_name(list_sheet_names):
                for view_port in self.get_legend_viewport_by_name(list_legend_names):
                    if sheet.Id != activeView.Id:
                        try:
                            new_vp = Viewport.Create(doc, sheet.Id, view_port.ViewId, view_port.GetBoxCenter())
                            new_vp.ChangeTypeId(view_port.GetTypeId())
                        except:
                            pass

            t.Commit()
            
        except Exception as e:
            MessageBox.Show(str(e), "Message")
  
class WPFListBoxUtils:
    def __init__(self, listbox):
        self.listbox = listbox

    def To_Checked_ListBox (self, list_strings):
        self.listbox.Items.Clear()
        for content in list_strings:
            checkBox = CheckBox()
            checkBox.Content = content
            checkBox.Click += self.CheckBox_Click
            self.listbox.Items.Add(checkBox)
        return self.listbox.Items
    
    def CheckBox_Click(self, sender, e):
        checkBox = sender
        if isinstance(checkBox, CheckBox):
            for item in self.listbox.SelectedItems:
                if isinstance(item, CheckBox):
                    item.IsChecked = checkBox.IsChecked
    
    def Checked_Items_To_List (self):
        items_checked = []
        for item in self.listbox.Items:
            if item.IsChecked:
                items_checked.append(item)
        return items_checked
    
    def Items_To_List_CheckBox (self):
        items = []
        for item in self.listbox.Items:
            items.append(item)
        return items

    def Show_Unchecked (self, list_unchecked):
        for key in list_unchecked.keys():
            item = list_unchecked[key]
            self.listbox.Items.Insert(key, item)
        
    def Select_All(self, is_select_all):
        for checkBox in self.listbox.Items:
            if isinstance(checkBox, CheckBox):
                checkBox.IsChecked = is_select_all

#region defind window
class WPFWindow:

    def load_window (self):
        
        #import window from .xaml file path
        file_stream = FileStream(xaml_file_path, FileMode.Open, FileAccess.Read)
        window = XamlReader.Load(file_stream)

        #get/set controls
        self.cbb_SheetSet = window.FindName("cbb_SheetSet")
        self.lbx_Legends = WPFListBoxUtils(window.FindName("lbx_Legends"))
        self.lbx_Sheets = WPFListBoxUtils(window.FindName("lbx_Sheets"))
        self.bt_SelectAllLegends = window.FindName("bt_SelectAllLegends")
        self.bt_SelectNoneLegends = window.FindName("bt_SelectNoneLegends")
        self.bt_SelectAllSheets = window.FindName("bt_SelectAllSheets")
        self.bt_SelectNoneSheets = window.FindName("bt_SelectNoneSheets")
        self.tb_Filter = window.FindName("tb_Filter")
        self.bt_Cancel = window.FindName("bt_Cancel")
        self.bt_OK = window.FindName("bt_OK")
        
        #binding data
        self.binding_data()
        self.window = window

        return window


    def binding_data (self):
        self.sheet_names = Utils().get_sheet_by_view_set("<all>")
        legend_names = Utils().get_all_legend_names()

        self.cbb_SheetSet.ItemsSource = Utils().get_all_sheetSet()
        self.cbb_SheetSet.SelectedIndex = 0
        self.lbx_Legends.To_Checked_ListBox(legend_names)
        self.lbx_Sheets.To_Checked_ListBox(self.sheet_names)
        

        #button
        self.bt_Cancel.Click += self.Cancel_Click
        self.bt_OK.Click += self.Ok_Click
        self.tb_Filter.TextChanged += self.Text_Filter_Changed
        self.cbb_SheetSet.SelectionChanged += self.SheetSet_Changed
        self.bt_SelectAllLegends.Click += self.Select_All_Legends
        self.bt_SelectNoneLegends.Click += self.Select_None_Legends
        self.bt_SelectAllSheets.Click += self.Select_All_Sheets
        self.bt_SelectNoneSheets.Click += self.Select_None_Sheets

        self.sheet_backup = self.lbx_Sheets.Items_To_List_CheckBox()
        self.count = 0

    def SheetSet_Changed (self, sender, e):
        view_set_name = self.cbb_SheetSet.SelectedItem
        self.sheet_names = Utils().get_sheet_by_view_set(view_set_name)
        self.lbx_Sheets.To_Checked_ListBox(self.sheet_names)
        self.tb_Filter.Text = ""


    def Ok_Click(self, sender, e):
        legend_names = []
        sheet_names = []

        for item in self.lbx_Legends.Checked_Items_To_List():
            legend_names.append(item.Content)
       
        for item in self.lbx_Sheets.Checked_Items_To_List():
            sheet_names.append(item.Content)


        if len(legend_names) > 0 and len(sheet_names) > 0:
            self.window.Close()
            Utils().copy_legends_to_sheet(legend_names, sheet_names)
            MessageBox.Show("Completed!","Message")
        else: 
            MessageBox.Show("Please select legends/sheets to copy!","Message")


    def Cancel_Click (self, sender, e):
        self.window.Close()
    
    def Text_Filter_Changed (self, sender, e):
        if self.count == 0: self.sheet_backup = self.lbx_Sheets.Items_To_List_CheckBox()

        list_contain = []
        filter_name = str(self.tb_Filter.Text).lower()
        if filter_name is not None or filter_name != "":
            for item in self.sheet_backup:
                if str(item.Content).lower().__contains__(filter_name):
                    list_contain.append(item)
            
            self.window.FindName("lbx_Sheets").Items.Clear()
            for item in list_contain:
                self.window.FindName("lbx_Sheets").Items.Add(item)
            self.count +=1
        else: self.count= 0


    
    def Select_All_Legends(self, sender, e):
        self.lbx_Legends.Select_All(True)

    def Select_None_Legends(self, sender, e):
        self.lbx_Legends.Select_All(False)

    def Select_All_Sheets(self, sender, e):
        self.lbx_Sheets.Select_All(True)

    def Select_None_Sheets(self, sender, e):
        self.lbx_Sheets.Select_All(False)



#endregion

def main_task():
    try:
        if activeView.ViewType == ViewType.DrawingSheet:
            legends = Utils().get_all_legend_names()
            if len(legends) > 0:
                window = WPFWindow().load_window()
                window.ShowDialog()
            else: MessageBox.Show("No legends in active sheet!", "Message")
        else:
            MessageBox.Show("Active a sheet to run tool!", "Message")
    except Exception as e:
        MessageBox.Show(str(e), "Message")

if __name__ == "__main__":
    main_task()
        






