import unittest


class TestSqlAlchemy(unittest.TestCase):
    engine = create_engine('mysql+mysqlconnector://root:a123456@localhost:3306/test')

