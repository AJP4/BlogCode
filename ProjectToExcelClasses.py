import win32com.client
import win32ui
import pandas as pd
import datetime as dt
import sys
import collections
import os.path
import pathlib
import dateutil.parser as dt_parse
import logging
import logging.handlers
import datetime as dt

class DataFrameOfMSProject(object):
       
    def __init__(self, headers=None, ms_project_file=None, logging_level="INFO", UniqueIDs_to_Ignore=[]):
        """Creates and returns a Pandas dataframe of a "flattened" MSProject file.  By "Flattened"
        means a table of tasks with summary tasks collapsed to one line, e.g. "Level 1 > Level 2 > Level 3"
        
        Keyword Arguments:
            headers {list} -- List of headers using MSProject exact field names (default: {None})
            ms_project_file {str} -- full path string to MSPorject File (r"C/path/to/MyProjectFile.mpp" (default: {None})
            logging_level {str} -- Options: "DEBUG" or "INFO" (default: {"INFO"})
            UniqueIDs_to_Ignore {list} -- Provide a list (Unique Task Ids) of any tasks that need to be ignored (default: {[]})
        """
        #create directory for log files if one does not exist
        pathlib.Path('log').mkdir(parents=True, exist_ok=True)        

        self.logging_level=logging_level
        self.set_up_Logger()
        self.logger.info("Initiation")
        self.ms_project_file = ms_project_file
        self.UniqueIDs_to_Ignore=UniqueIDs_to_Ignore
        

        if self.ms_project_file:
            """If ms_project_file is a valid MS Project File setup MSProject application,
            create MSProject object reference, load pjFields, then close the MSProject file
            """

            self.logger.debug("inside: def __init__ loop:  if self.ms_project_file")
            self.__mspApplication, self.__project = self.ms_project_object(self.__ms_project_file)

            self.headers = headers
            if self.headers:
                # tested if HeadersList is a valid set of MS Project Headers
                # can only do after MS Project file has been loaded as need to use Projec to check
                self.logger.debug("inside: def __init__:  if self.headers")
                self.__create_project_data_frame()
                self.ms_project_object_close()
            else:
                self.logger.error("Headers list provided contains error")
                print("Headers list provided contains error")
                self.ms_project_object_close()
        else:
            # tested if ms_project_file is a valid MS Project File
            print("Not a MS Project File or File does not exist")
            self.logger.error("Not a MS Project File or File does not exist")
        
        
     
    def set_up_Logger(self):
        # delete old log
        now=dt.datetime.now().strftime("%Y-%m-%d %H-%M-%S")        
        
        # create logger
        self.logger = logging.getLogger('Log')
        if self.logging_level=="DEBUG":
            self.logger.setLevel(logging.DEBUG)
        elif self.logging_level=="INFO":
            self.logger.setLevel(logging.INFO)            
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
        
    @property
    def ms_project_file(self):
        """returns the MSProject File String"""
        self.logger.debug("Entered: @Property > ms_project_file")
        return self.__ms_project_file
    
    @ms_project_file.setter
    def ms_project_file(self, f):
        """
        Sets the ms_project_file property once checking that the file exists and is a project file
        If None is entered, it prompts for the file via a dialog box
        """
        if f is None:
            f = self.select_file(ext="*.mpp", filters="MS Project Files (*.mpp)|*.mpp||")
            # see nullege.com/codes/search/win32ui.CreateFileDialog
        self.logger.debug("Entered: @Property.setter > ms_project_file")
        file = pathlib.Path(f)
        self.logger.debug("file.is_file() ==True %s ", file.is_file())
        self.logger.debug("file.suffix == .mpp %s ", file.suffix == ".mpp")
        if file.is_file() is True and file.suffix == ".mpp":
            self.logger.debug("Entered: @Property.setter > ms_project_file is a proper project file")
            self.__ms_project_file = f
        else:
            self.logger.error("Entered: @Property.setter > ms_project_file is NOT a proper project file")
            #print("Entered: @Property.setter > ms_project_file is NOT a proper project file")  # debug
            self.__ms_project_file = False

    @staticmethod
    def select_file(ext, filters):
        dia = win32ui.CreateFileDialog(1, ext, None, 0, filters)
        # see nullege.com/codes/search/win32ui.CreateFileDialog
        dia.DoModal()
        return dia.GetPathName()     

    @property
    def headers(self):
        """
        returns headers
        """
        self.logger.debug("Entered: @Property > headers")
        return self.__headers
    
    @headers.setter
    def headers(self, extra_headers):
        """
        sets headers
        These headers "UniqueID", "SummaryTask","Name","Start","Finish","% Complete" will always be output
        """
        self.logger.debug("Entered: @Property.setter > headers")
        const_header = ["UniqueID", "SummaryTask", "Name", "Start", "Finish", "% Complete"]
        if extra_headers is None or extra_headers == []:
            self.logger.debug("Inside def headers.setter : TRUE, h==None or h==[]")
            self.logger.info("No extra headers set because extra_headers = %s",extra_headers)
            self.__headers = const_header+["Resource Names", "Notes", "Predecessors"]
        else:
            self.logger.info("***Extra headers requested: %s", extra_headers)
            if self.__do_list_of_headers_exist_in_ms_project(extra_headers):
                self.logger.debug("****in: if self.__doListOfHeadersExistInMSProject(h)")
                self.__headers = []
                [self.__headers.append(v) for v in const_header + extra_headers if v not in self.__headers]
                self.logger.info("****** final list of headers %s", self.__headers)
            else:
                self.__headers = False
            
    def __do_list_of_headers_exist_in_ms_project(self, header_list):
        """Checks to see if all the headers listed exist in MS Project
        
        Arguments:
            header_list {list} -- list of strings of the MSProject field names that are to be included in the Pandas Dataframe
        
        Returns:
            [boolean] -- True if headers are valid and have been loaded, False if the list contains non valid headers
        """

        self.logger.debug("def __doListOfHeadersExistInMSProject")
        for header in header_list:
            self.logger.debug("%s",header)
            try:
                self.__mspApplication.FieldNameToFieldConstant(header)
            except:
                self.logger.error("%s", header, " is not a valid MS Project Heading")
                print ("%s", header, " is not a valid MS Project Heading")
                return False
        self.logger.info("header list is good") 
        return True

    def ms_project_object(self, path_to_ms_project):
        """load MS Project File and create reference to MS Project File object
        return the project object and msp application object
        
        Arguments:
            path_to_ms_project {str} -- [full filepath to MSProject File]
        
        Returns:
            [win32com object] -- [win32com object for the project and the MSP application]
        """

        self.logger.debug("Entered: defMSProjectObject")
        self.logger.info("path to MS Project file entered is %s", path_to_ms_project)
        msp = win32com.client.Dispatch("MSProject.Application")
        msp.DisplayAlerts=False
        msp.FileOpen(path_to_ms_project)
        project = msp.ActiveProject
        msp.DisplayAlerts=True
        self.logger.debug("Inside def MSProjectObject > created msp application object and project Object")
        return msp, project
    
    def ms_project_object_close(self):
        """Closes the MS Project Application
        """

        self.logger.debug("Entered: def MSProjectObjectClose")
        self.__mspApplication.FileSave()
        self.__mspApplication.Quit()
        self.logger.debug("Inside def MSProjectObjectClose > closed application")
        return
    
    @property
    def project_data_frame(self):
        """
        returns project DataFrame
        """
        self.logger.debug("Entered: @Property > projectDataFrame")
        return self.__projectDataFrame       
    
    @project_data_frame.setter
    def project_data_frame(self, headers):
        """
        Sets the projectDataFrame
        """
        self.logger.debug("Entered: __setDataFrame")
        pass
        
    def __create_project_data_frame(self):
        self.__projectDataFrame = pd.DataFrame(columns=self.__headers)
        summary_tasks_to_task = []
        task_collection= self.__project.Tasks
        
        # A list containing the tasks that were not ignored, initially contains all task to be ignored
        notIgnoredTask=self.UniqueIDs_to_Ignore[:]
        
        for t in task_collection:
            # print to log if as task to be ignored was found
            if (t.UniqueID in self.UniqueIDs_to_Ignore):
                self.logger.info("Task %s was ignored as requested", str(t.UniqueID))
                notIgnoredTask.remove(t.UniqueID)
            if (not t.Summary) & ~(t.UniqueID in self.UniqueIDs_to_Ignore):  # i.e. it is a task line not a Summary Task
                # find dependent task
                dep = []  # an empty list to add dependent task id
                self.logger.debug("Collecting Task Dependencies for %s", t.UniqueID)
                for d in t.TaskDependencies:
                    if int(d.From) != t.UniqueID:  # a task can have multiple references to itself, not sure why, but this removes them
                        dep.append(str(d.From) + "-" + str(d.From.Name))
                
                # collect resource names        
                res = []  # an empty list to add resources
                self.logger.debug("Collecting Task Resources for %s", t.UniqueID)
                for r in t.Assignments:
                    self.logger.debug("ResourceName is %s", r.ResourceName)
                    res.append(r.ResourceName)    
                
                # it is not good practic but it is possible to have project tasks at the top level (outline level 1)
                # So this if statement catches those occurances and empties summary_tasks_to_task list                 
                if t.OutlineLevel ==1:
                    summary_tasks_to_task = []                
                
                # create a temporary "temp" list variable holding the entries for the dataframe row.
                sum_task = ">".join(summary_tasks_to_task)
                dependencies = [", ".join(dep)]
                resources = [", ".join(res)]
                
                temp = [t.UniqueID, sum_task]
                for head_title in self.__headers:
                    if head_title != "UniqueID" and head_title != "SummaryTask":
                        # note that dependencies and resources have been created by iterating over their
                        # respective collection objects and are therefore not found via Task.GetField
                        if head_title == "Predecessors":
                            temp = temp+dependencies
                        elif head_title=="Resource Names":
                            temp = temp+resources
                        else:
                            temp = temp+[t.GetField(self.__mspApplication.FieldNameToFieldConstant(head_title))]
                       
                self.__projectDataFrame = self.__projectDataFrame.append(pd.Series(temp, index=self.__headers), ignore_index=True)
            
            elif t.Summary & (t.OutlineLevel > len(summary_tasks_to_task)):
                # if tasks is a summary task and its outline level is greater than number of summary tasks in the list
                # summaryTasksToTask then add that summary task to the list
                if not(t.UniqueID in self.UniqueIDs_to_Ignore):
                    summary_tasks_to_task.append(t.Name)
                
            else:
                if not(t.UniqueID in self.UniqueIDs_to_Ignore):
                    while not len(summary_tasks_to_task) == t.OutlineLevel - 1:
                        # if tasks is a summary task and its outline level is less than number of summary tasks in the list
                        # summaryTasksToTask then remove last summary task from list and add new summary task to the list
                        summary_tasks_to_task.pop()
                
                    summary_tasks_to_task.append(t.Name)
        
        # print to log the to be ignored tasks that were not ignored as not in the project file
        for t in notIgnoredTask:
            self.logger.info("Task %s was not ignored, as not in project file", str(t))
        
        # finally, set the index of the dataframe to the unique MS Project Task ID
        self.__projectDataFrame = self.__projectDataFrame.set_index("UniqueID")
        self.__projectDataFrame["Finish"] = pd.to_datetime(self.__projectDataFrame["Finish"],dayfirst=True)
        self.__projectDataFrame["Start"] = pd.to_datetime(self.__projectDataFrame["Start"],dayfirst=True)
        return        
    
    def output_dictionary_of_data_frames_FINISHING(self, due_date=None, header_to_filter=None, filter_text=None,
                                         duration_of_periods=7, num_of_periods=5, flag_incomplete_only=True):
        """
        (1) Outputs a dictionary of DataFrames based duration of the periods
        (e.g. duration_of_periods = 7 for 7 days (or by week))
        (2) number of periods from the due_date
        (3) It will also produce the output dataframe based on a filter.  use this when wanting to 
        only output certain tasks (e.g. filter by resource name).  The filter uses regular expressions
        (4) to note whether to put all tasks or only incomplete tasks.  flag_incomplete_only=True surpressing output of complete tasks
        """
        data_frame_collection = collections.OrderedDict()
        df_filtered = self.project_data_frame

        if due_date is None:  # if no date offered, use todays date
            due_date = dt.datetime.today().date()
        else:
            due_date = dt.datetime.strptime(due_date, "%d/%m/%Y").date()
        
        if flag_incomplete_only is True:
            df_filtered=df_filtered[df_filtered["% Complete"] != "100%"]     

        """Process the periods"""
        self.logger.info("dueDate = %s", due_date)
        self.logger.info("header to filter = %s", header_to_filter)
        self.logger.info("header to filter = %s", header_to_filter)
        self.logger.info("Filter Text = %s", filter_text)
        self.logger.info("Duration of periods = %s", duration_of_periods)
        self.logger.info("Number of Periods = %s", num_of_periods)
        self.logger.info("Flag = %s", flag_incomplete_only)
        
        key="Overdue"
        data_frame_collection[key] = df_filtered[df_filtered["Finish"] <= due_date]
        
        date = due_date + dt.timedelta(days = 1)
        self.logger.info("due_date = %s",due_date)
        self.logger.info("date = %s",date)
        
        # create worksheet for each period (other than "OverDue")
        for reportingWindow in range(0, num_of_periods-1):
            from_date = date
            to_date = date + dt.timedelta(days = duration_of_periods-1)
            
            # create sheet name based on period range
            key = str(from_date - dt.timedelta(days=0)) + "<>" + str(to_date - dt.timedelta(days=0))
            self.logger.info("key = %s", key)

            # filter for tasks between the date range
            data_frame_collection[key] = df_filtered[(df_filtered["Finish"] >= from_date) &
                                                     (df_filtered["Finish"] <= to_date)]
            
            date = to_date + dt.timedelta(days=1)
        return data_frame_collection
    def output_dictionary_of_data_frames_WIP(self, due_date=None, header_to_filter=None, filter_text=None,
                                             duration_of_periods=7, num_of_periods=5, flag_incomplete_only=True, flag_OUTPUT_WIP_COLUMN=True):
        """
        (1) Outputs a dictionary of DataFrames based duration of the periods
        (e.g. duration_of_periods = 7 for 7 days (or by week))
        (2) number of periods from the due_date
        (3) It will also produce the output dataframe based on a filter.  use this when wanting to 
        only output certain tasks (e.g. filter by resource name).  The filter uses regular expressions
        (4) to note whether to put all tasks or only incomplete tasks.  flag_incomplete_only=True surpressing output of complete tasks
        """
        data_frame_collection = collections.OrderedDict()
        df_filtered = self.project_data_frame
        self.logger.debug("Entered : output_dictionary_of_data_frames_WIP")
        if due_date is None:  # if no date offered, use todays date
            due_date = dt.datetime.today().date()
        else:
            due_date = dt.datetime.strptime(due_date, "%d/%m/%Y").date()

        if flag_incomplete_only is True:
            df_filtered=df_filtered[df_filtered["% Complete"] != "100%"]     

        """Process the periods"""
        self.logger.info("dueDate = %s", due_date)
        self.logger.info("header to filter = %s", header_to_filter)
        self.logger.info("header to filter = %s", header_to_filter)
        self.logger.info("Filter Text = %s", filter_text)
        self.logger.info("Duration of periods = %s", duration_of_periods)
        self.logger.info("Number of Periods = %s", num_of_periods)
        self.logger.info("Flag = %s", flag_incomplete_only)

        key="Overdue"
        data_frame_collection[key] = df_filtered[df_filtered["Finish"] <= due_date]

        date = due_date + dt.timedelta(days = 1)
        self.logger.info("due_date = %s",due_date)
        self.logger.info("date = %s",date)        

        # create worksheet for each period (other than "OverDue")
        for reportingWindow in range(0, num_of_periods):
            from_date = date
            to_date = date + dt.timedelta(days = duration_of_periods-1)
            self.logger.info("Inside reportingWindow: from_date = %s",from_date)
            self.logger.info("Inside reportingWindow: to_date = %s",to_date)
            # create sheet name based on period range
            key = str(from_date - dt.timedelta(days=0)) + "<>" + str(to_date - dt.timedelta(days=0))
            self.logger.info("key = %s", key)
            # filter for tasks between the date range
            df_temp=df_filtered.copy(deep=True)
            
            df_temp=df_temp[((df_temp["Start"] >= from_date) & (df_temp["Start"] <= to_date) & (df_temp["Finish"]>=from_date)) |
                        ((df_temp["Start"] <= from_date) & (df_temp["Finish"] >= to_date) & (df_temp["Finish"]>=to_date) & (df_temp["Start"]<=to_date)) |
                        ((df_temp["Start"] <=from_date) & (df_temp["Finish"]>=from_date) & (df_temp["Finish"]<=to_date)& (df_temp["Start"]<=to_date))]
        
            self.logger.debug("flag_OUTPUT_WIP_COLUMN = %s",flag_OUTPUT_WIP_COLUMN)
            if flag_OUTPUT_WIP_COLUMN:
                df_temp["WIP"]="WIP"
                df_temp.ix[df_temp["Start"] >= from_date,"WIP"]="Starting in Period"
                df_temp.ix[df_temp["Finish"]<to_date,"WIP"]="Finishing in Period"
                df_temp.ix[(df_temp["Start"] >= from_date)&(df_temp["Finish"]<=to_date),"WIP"]="Starting & Finishing in Period"
                      
                self.logger.debug("to_date = %s",to_date)
                self.logger.info("Added WIP Column")
            data_frame_collection[key] = df_temp.copy(deep=True)
            date = to_date + dt.timedelta(days=1)   
        return data_frame_collection
            


        
        
        
    
    