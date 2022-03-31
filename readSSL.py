#!/usr/bin/env python
# -*- coding: utf-8 -*-
# ==============
#      Main script file
# ==============
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
sys.path.append("../../")
import ctypes as c
from pym1com import pym1com
import time
import pandas as pd

class ParGui():
    def __init__(self, parent=None):
        self.ip = ""
        self.config = {}
        try:
            execfile('c:\\LW\\LW_config\\hmi.conf', self.config)
            self.ip = self.config["plc_ip"]
        except Exception as e:
            print(str(e))

    def update(self):
        pass

    def connectM1(self, ip):
        print("connecting.. ")
        self.ip = ip
        # self.m1 = pym1com.M1(self.ip, None, 'M1', 'bachmann', None, 'PAR_GUI')
        ssl = ('protocol' in self.config and self.config['protocol'] == "ssl")
        pkcs12 = None
        pkcs12Password = None
        if ssl:
            pkcs12 = self.config["pkcs12"]
            pkcs12Password = self.config["pkcs12_passwd"]

        self.m1 = pym1com.M1(self.ip, None, 'admin', 'windy-EEWLL',
                             console=None, ssl=ssl, pkcs12=pkcs12,
                             pkcs12Password=pkcs12Password)

        self.SSL = self.m1.readData('VALIDATE/endDateCertificate')
        
        #self.lub = self.m1.readData('SUPERVIS/.settings.pSetMainBearLubLvlTimeout.actual')
        
       
log = []
for i in range(1, 25):
    try:
        ip = '10.2.{}.1'.format(i)
        M1 = ParGui()
        M1.connectM1(ip)
        note = {'WPT':i, 'SSL':M1.SSL}
        log.append(note)
        
    except Exception as err:
        print(err)
        continue


data = pd.DataFrame(log)
data.to_excel('SSL.xlsx', index=False)
     
