import pickle
import pandas as pd
import win32com.client as win32
import os
import logging
import logging.handlers
import datetime as dt

class excelFormatColumns():
        
    def __init__(self):
        self.set_up_Logger()
        self.logger.info("Initiation")
        
        self.__wrap_text_cols_ColWidth_small=[]
        self.__wrap_text_cols_ColWidth_medium=[]
        self.__wrap_text_cols_ColWidth_large=[]
        self.__date_cols=[]
        self.__autofit_cols=[]
        self.__date_format="dd/mm/yyyy"
        self.colWidths=[30,50,80] #small, medium and large col widths

    def set_up_Logger(self):
        # delete old log
        now=dt.datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        
        # create logger
        self.logger = logging.getLogger('Log')
        self.logger.setLevel(logging.DEBUG)         
        # create formatter
        formatter = logging.Formatter('%(asctime)s - %(name)-4s - %(levelname)-8s - function: %(funcName)-15s LineNum: %(lineno)-5d - %(message)s',
                                          datefmt='%m-%d-%Y %H:%M')
        logpath="log/"+now+__name__
        handler = logging.handlers.RotatingFileHandler(logpath+".log", maxBytes=1000000, backupCount=5)  
        
        handler.setLevel(logging.DEBUG)
        handler.setFormatter(formatter)
        self.logger.addHandler(handler)
        self.logger.addHandler(handler)
            
        return
        
    def excelFile(self,xlPath):
        self.excelApp = win32.gencache.EnsureDispatch("Excel.Application")
        self.wb = self.excelApp.Workbooks.Open(xlPath)
        self.excelApp.Visible = False
        self.logger.info("%s",xlPath)
    
    def formatExcel(self, listOfSheets):
        self.logger.debug("Processing these Worksheets %s",listOfSheets)
        for sheet in reversed(listOfSheets):
            self.logger.debug("process sheet %s",sheet)
            ws = self.wb.Worksheets(sheet)
            ws.Activate()
            self.excelApp.ActiveWindow.Zoom = 60
    
            for col in self.__date_cols:
                ws.Columns(col).NumberFormat=self.__date_format
    
            for col in self.__autofit_cols:
                ws.Columns(col).AutoFit()
    
            for col in self.__wrap_text_cols_ColWidth_small:
                ws.Columns(col).ColumnWidth = self.colWidths[0]
                ws.Columns(col).WrapText = True 
    
            for col in self.__wrap_text_cols_ColWidth_medium:
                ws.Columns(col).ColumnWidth = self.colWidths[1]
                ws.Columns(col).WrapText = True 
    
            for col in self.__wrap_text_cols_ColWidth_large:
                ws.Columns(col).ColumnWidth = self.colWidths[2]
                ws.Columns(col).WrapText = True           
        
            ws.Rows.AutoFit()
            ws.Rows.VerticalAlignment = -4160
        
        self.wb.Close(SaveChanges=1)
        return 
    
    @property    
    def wrap_text_cols_ColWidth_small(self):
        """returns the list of columns requiring the smallest width"""
        return self.__wrap_text_cols_ColWidth_small
   
    @wrap_text_cols_ColWidth_small.setter
    def wrap_text_cols_ColWidth_small(self,colList):
        """sets the list of columns requiring the largest width"""
        self.__wrap_text_cols_ColWidth_small = colList
   
    @property    
    def wrap_text_cols_ColWidth_medium(self):
        """returns the list of columns requiring the medium width"""
        return self.__wrap_text_cols_ColWidth_medium
   
    @wrap_text_cols_ColWidth_medium.setter
    def wrap_text_cols_ColWidth_medium(self,colList):
        """sets the list of columns requiring the medium width"""
        self.__wrap_text_cols_ColWidth_medium = colList   
   
    @property    
    def wrap_text_cols_ColWidth_large(self):
        """returns the list of columns requiring the largest width"""
        return self.__wrap_text_cols_ColWidth_large
   
    @wrap_text_cols_ColWidth_large.setter
    def wrap_text_cols_ColWidth_large(self,colList):
        """sets the list of columns requiring the largest width"""
        self.__wrap_text_cols_ColWidth_large = colList      
   
    @property    
    def date_cols(self):
        """returns the list of columns requiring date format"""
        return self.__date_cols
   
    @date_cols.setter
    def date_cols(self,colList):
        """sets the list of columns requiring date format"""
        self.__date_cols = colList
        
    @property    
    def autofit_cols(self):
        """returns the list of columns requiring autofit width"""
        return self.__autofit_cols
   
    @autofit_cols.setter
    def autofit_cols(self,colList):
        """sets the list of columns requiring autofit width"""
        self.__autofit_cols = colList        
    
    @property    
    def date_format(self):
        """date format"""
        return self.__date_format
   
    @date_format.setter
    def date_format(self, dateFormat):
        """sets the date format"""
        self.__date_format = dateFormat
        
    @property    
    def colWidths(self):
        """returns the list of column widths [small, medium, large]"""
        return self.__colWidths
   
    @colWidths.setter
    def colWidths(self, ColW_list):
        """sets list of small, medium and large column widths [30,50,80]"""
        self.__colWidths = ColW_list      
   
   
   

