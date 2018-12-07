#coding=utf-8
"""
Create On 2018/12/7

@author: Ron2
"""

import unittest
import time
import json
from game.skill_pb2 import Skill
from game.skill_pb2 import SkillArray


class TestProtobuf(unittest.TestCase):
    """
    """

    def setUp(self):
        """
        :return: 
        """
        self.testNum = 100000


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
        # print u"data => ", data
        # print "size => ", len(data)

        ''' 模式测试多个的 '''
        skillArrayObj = SkillArray()
        for x in xrange(self.testNum):
            skillObj = skillArrayObj.skill.add()
            skillObj.skillId = 6 + x
            skillObj.name = u"冲击对手"
            skillObj.desc = u"能打断boss技能"

        f = open("./skill.dat",  "wb")
        f.write(skillArrayObj.SerializeToString())
        f.close()


    def test_readData(self):
        """
        测试读取数据
        :return: 
        """
        beginTime = self.getNow()

        skillArrayObj = SkillArray()
        f = open("./skill.dat", "rb")
        skillArrayObj.ParseFromString(f.read())
        f.close()


        for skillObj in skillArrayObj.skill:
            if skillObj.skillId % 10000 == 0:
                print skillObj.skillId
                # print "id => ", skillObj.skillId
                # print "name => ", skillObj.name
                # print "desc => ", skillObj.desc

        print "Protocol CostTime => ", self.getNow() - beginTime



    def test_writeJson(self):
        """
        :return: 
        """
        ar = []
        for x in xrange(self.testNum):
            dic = {}
            dic["skillId"] = x
            dic["name"] = u"冲击对手"
            dic["desc"] = u"能打断boss技能"
            ar.append(dic)

        data = json.dumps(ar)
        f = open("./json.dat", "wb")
        f.write(data)
        f.close()


    def test_readJson(self):
        """
        :return: 
        """
        class MySkill(object):
            def __init__(self):
                self.skillId = 0
                self.name = ""
                self.desc = ""

            def fromDic(self, dic):
                """
                :param dic: 
                :return: 
                """
                self.skillId = dic.get("skillId")
                self.name = dic.get("name")
                self.desc = dic.get("desc")

        # 测试
        self.test_readData()

        beginTime = self.getNow()

        f = open("./json.dat", "rb")
        data = f.read()
        f.close()

        ar = json.loads(data)
        for dic in ar:
            skill = MySkill()

            skill.fromDic(dic)
            if skill.skillId % 10000 == 0:
                print skill.skillId

        print "json CostTime ==> ", self.getNow() - beginTime


    def getNow(self):
        """
        :return:
        """
        return time.time()

