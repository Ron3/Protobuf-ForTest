#coding=utf-8
"""
Create On 2018/12/7

@author: Ron2
"""

import unittest
from game.skill_pb2 import Skill


class TestProtobuf(unittest.TestCase):
    """
    """

    def setUp(self):
        """
        :return: 
        """


    def tearDown(self):
        """
        :return: 
        """


    def test_writeData(self):
        """
        测试写数据的
        :return: 
        """
        skillObj = Skill()
        skillObj.skillId = 5
        skillObj.name = u"激光浮游炮"
        skillObj.desc = u"此技能会造成5点伤害"
        # print type(skillObj), dir(skillObj)
        data = skillObj.SerializeToString()
        print u"data => ", data
        # print "size => ", len(data)

        readObj = Skill()
        readObj.ParseFromString(data)
        print u"skillId => ", readObj.skillId
        print u"name => ", readObj.name
        print u"desc => ", readObj.desc







