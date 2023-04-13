# Author: Felix He (GN Audio)
# Contact: fhe@jabra.com

import threading
import datetime
import logging
import os
#a singleton class

class Logger(object):
    instance = None
    curr_log_name = ""
    __fh = None
    __ch = None
    mutex=threading.Lock()
    
    def __get_fh(self):
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        self.__fh = logging.FileHandler('./trace_log/%s'%self.curr_log_name, encoding='utf-8')
        self.__fh.setFormatter(formatter)
    
    def __get_ch(self):
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        self.__ch = logging.StreamHandler()
        self.__ch.setFormatter(formatter)
        
    def __init__(self):
        if not os.path.exists("./trace_log/"):
            os.mkdir("./trace_log/")
    
        self.fileLogger=logging.getLogger('FileLogger')
        self.fileLogger.setLevel(logging.DEBUG)
        
        self.stdLogger=logging.getLogger('StdLogger ')
        self.stdLogger.setLevel(logging.DEBUG)
        
        self.curr_log_name = "AutoTest_%s.log"%datetime.datetime.now().strftime('%Y-%m-%d')
        self.__get_fh()
        self.__get_ch()
        
        self.fileLogger.addHandler(self.__fh)
        self.fileLogger.addHandler(self.__ch)
        self.stdLogger.addHandler(self.__ch)
        self.stdLogger.addHandler(self.__fh)
        #log.removeHandler(fileTimeHandler)
        
    @staticmethod
    def ins():
        if Logger.instance == None:
            Logger.mutex.acquire()
            Logger.instance = Logger()
            Logger.mutex.release()
        return Logger.instance
        
    def file_logger(self):
        if self.curr_log_name != "AutoTest_%s.log"%datetime.datetime.now().strftime('%Y-%m-%d'):
            self.curr_log_name = "AutoTest_%s.log"%datetime.datetime.now().strftime('%Y-%m-%d')
            self.fileLogger.removeHandler(self.__fh)
            self.stdLogger.removeHandler(self.__fh)
            self.__get_fh()
            self.fileLogger.addHandler(self.__fh)
            #self.stdLogger.addHandler(self.__fh)
            
        return self.fileLogger
        
    def std_logger(self):
        if self.curr_log_name != "AutoTest_%s.log"%datetime.datetime.now().strftime('%Y-%m-%d'):
            self.curr_log_name = "AutoTest_%s.log"%datetime.datetime.now().strftime('%Y-%m-%d')
            self.fileLogger.removeHandler(self.__fh)
            self.stdLogger.removeHandler(self.__fh)
            self.__get_fh()
            self.fileLogger.addHandler(self.__fh)
            self.stdLogger.addHandler(self.__fh)
        return self.stdLogger

    
    
