# -*- coding: utf-8 -*-

import clr
clr.AddReference("System")
# List<ElementType>() <- it is a special type of list that RevitAPI often requires.
from System.Collections.Generic import List
import os, sys, math, datetime, time
from Autodesk.Revit.DB import *
# pyRevit
from pyrevit import revit, forms

doc     = __revit__.ActiveUIDocument.Document
uidoc   = __revit__.ActiveUIDocument
app     = __revit__.Application

titleblock = FilteredElementCollector(doc)\
            .OfCategory(BuiltInCategory.OST_TitleBlocks)\
            .WhereElementIsElementType()\
            .FirstElementId()


def change_case(sheetlist, upper=True, verbose=False):
    with revit.Transaction('Rename Sheets to Upper'):
        for el in sheetlist:
            sheetnameparam = el.Parameter[BuiltInParameter.SHEET_NAME]
            orig_name = sheetnameparam.AsString()
            new_name = orig_name.upper() if upper else orig_name.lower()
            if verbose:
                print('RENAMING:\t{0}\n'
                      '      to:\t{1}\n'.format(orig_name, new_name))
            sheetnameparam.Set(new_name)


sel_sheets = forms.select_sheets(title="Select Desired Sheets", use_selection = True)

if sel_sheets:
    selected_option, switches = \
        forms.CommandSwitchWindow.show(
            ['to UPPERCASE',
             'to lowercase'],
            switches=['Show Report'],
            message='Select rename option:'
            )

    if selected_option:
        change_case(sel_sheets,
                    upper=True if selected_option == 'to UPPERCASE' else False,
                    verbose=switches['Show Report'])

