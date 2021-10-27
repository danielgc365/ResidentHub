################################################################################
# ExcelParser takes a GM workbook (SBR or main) and obtains the necessary data
# to create a calibration
#
#
# Usage for SDM Veoneer Residents ONLY
# Created by Daniel Gomez
# Date 10/27/2021
##################################################

# Imports
from openpyxl import Workbook, load_workbook


#############################################################
# Class GM workbook
# Takes the GM Workbook excel file and parse necessary data
# Input: Path to GM Workbook
#############################################################
class GMWorkbook:
    def __init__(self, path):
        self.path = path
        self.wb = load_workbook(self.path, keep_vba=True, read_only=True)
        pass

    def open_sheet(self, sheet="Release Form"):
        self.Sheet = self.wb[sheet]

    def get_sheet_list(self):
        for sheet in self.wb:
            print(sheet.title)

    def parse(self):
        try:
            self.CalPN = self.Sheet['E21'].value
            self.UtilityPN = self.Sheet['E19'].value
            self.AppPN = self.Sheet['E25'].value
            self.CalAlphaCode = self.Sheet['F21'].value + self.Sheet['G21'].value
            self.AppAlphaCode = self.Sheet['F25'].value + self.Sheet['G25'].value
            self.ModelYear = self.Sheet['E4'].value
            self.EMPN = self.Sheet['E13'].value
            self.SDMType = self.Sheet['C13'].value
        except AttributeError:
            print("No sheet selected")
        pass


#############################################################
# Class GM SBR workbook
# Inherits from the GM Workbook class
# it parses the data from the SBR excel file
# Input: Path to GM SBR Workbook
#############################################################
class SBRWorkbook(GMWorkbook):
    def __init__(self, path):
        super().__init__(path)

    def parse(self):
        try:
            self.CalPN = self.Sheet['D34'].value
            self.CalAlphaCode = self.Sheet['E34'].value + self.Sheet['F34'].value
        except AttributeError:
            print("Error")

#test_sbr = SBRWorkbook(r'C:\Python\ResidentHub\85541969AA_SBR.xlsm')
#test_sbr.open_sheet()
#test_sbr.parse()
