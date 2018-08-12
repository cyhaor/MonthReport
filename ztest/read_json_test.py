import json
import unittest


class TestReadJson(unittest.TestCase):
    with open("input/template_parameters.json", 'r') as load_f:
        excel_para_dict = json.load(load_f)
        print(excel_para_dict["deviceDateRunTime1"])
