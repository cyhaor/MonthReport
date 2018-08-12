# -*- coding:utf-8 -*-
import os
import win32com.client
from utils.mylogger import logger


def check_kill_process(process_name):
    try:
        wmi = win32com.client.GetObject('winmgmts:')
        process_code_cov = wmi.ExecQuery('select * from Win32_Process where Name="%s"' % process_name)
        if len(process_code_cov) > 0:
            os.system('taskkill /IM ' + process_name + ' /F')
            logger.info("kill %s process success. " % process_name)
        else:
            logger.info("%s process is not exist. " % process_name)
    except Exception as e:
        logger.error("check %s process fail: %s" % (process_name, e))
