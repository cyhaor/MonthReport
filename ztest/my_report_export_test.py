import os
import unittest
from utils.check_kill_process import check_kill_process
from utils.operator_excel import EasyExcel


class TestExcelExport(unittest.TestCase):
    check_kill_process('EXCEL.exe')
    input_template = os.getcwd() + r'\input\MR Healthy Report Template.xlsx'
    output_pdf = os.getcwd() + r'\output\MR Healthy Report Template.pdf'
    output_excel = os.getcwd() + r'\output\MR Healthy Report Template_output.xlsx'
    xls = EasyExcel(input_template)
    xls.set_para("maxErrorModuleNumPercentDevice1", "100%")
    num = "tableDeviceModuleErrorNum1"
    xls.set_table_para(num, ['9', '9', '9', '9', '9', '9', '9'])
    xls.save_as_pdf(output_pdf)
    xls.save(output_excel)
    xls.close()
    check_kill_process('EXCEL.exe')

