#!/bin/env python
# -*- encoding: cp932 -*-

##
#   @file PowerPointRTC.py
#   @brief PortPointControl Component

import win32com
import pythoncom
import pdb
from win32com.client import *
import pprint
import datetime
import msvcrt


import optparse
import sys,os,platform
import re
import time
import random
import commands
import math



import RTC
import OpenRTM_aist

from OpenRTM_aist import CorbaNaming
from OpenRTM_aist import RTObject
from OpenRTM_aist import CorbaConsumer
from omniORB import CORBA
import CosNaming

from ImpressControl import *


powerpointcontrol_spec = ["implementation_id", "PowerPointControl",
                  "type_name",         "PowerPointControl",
                  "description",       "PowerPoint Component",
                  "version",           "0.1",
                  "vendor",            "Miyamoto Nobuhiko",
                  "category",          "example",
                  "activity_type",     "DataFlowComponent",
                  "max_instance",      "10",
                  "language",          "Python",
                  "lang_type",         "script",
                  "conf.default.file_path", "NewFile",
                  "conf.default.SlideNumberInRelative", "1",
                  "conf.default.SlideFileInitialNumber", "0",
                  "conf.__widget__.file_path", "text",
                  "conf.__widget__.SlideNumberInRelative", "radio",
                  "conf.__widget__.SlideFileInitialNumber", "spin",
                  "conf.__constraints__.SlideNumberInRelative", "(0,1)",
                  "conf.__constraints__.SlideFileInitialNumber", "0<=x<=1000",
                  ""]




   


##
# @class PowerPointControl
# @brief PowerPointを操作するためのRTCのクラス
#

class PowerPointControl(ImpressControl):
    ##
    # @brief コンストラクタ
    # @param self 
    # @param manager マネージャーオブジェクト
    #
  def __init__(self, manager):
    ImpressControl.__init__(self, manager)
    
    prop = OpenRTM_aist.Manager.instance().getConfig()
    fn = self.getProperty(prop, "powerpoint.filename", "")
    self.m_powerpoint = PowerPointObject()
    
    if fn != "":
      str1 = [fn]
      OpenRTM_aist.replaceString(str1,"/","\\")
      fn = os.path.abspath(str1[0])
    self.m_powerpoint.Open(fn)

    self.conf_filename = ["NewFile"]

    self.slidenum = 0
    
    
    
    
    
    return

  ##
  # @brief rtc.confの設定を取得する関数
  #
  def getProperty(self, prop, key, value):
        
        if  prop.findNode(key) != None:
            #print value
            value = prop.getProperty(key)
        return value

  ##
  # @brief コンフィギュレーションパラメータが変更されたときに呼び出される関数
  # @param self 
  #
  def configUpdate(self):
      return
      """self._configsets.update("default","file_path")
      str1 = [self.conf_filename[0]]
      OpenRTM_aist.replaceString(str1,"/","\\")
      sfn = str1[0]
      tfn = os.path.abspath(sfn)
      if sfn == "NewFile":
        self.m_powerpoint.Open("")
      else:
        print sfn,tfn
        self.m_powerpoint.initCom()
        self.m_powerpoint.Open(tfn)"""
        #self.m_powerpoint.closeCom()


  ##
  # @brief 線描画
  # @param self 
  # @param bx 
  # @param by 
  # @param ex 
  # @param ey 
  def drawLine(self, bx, by, ex, ey):
    self.m_powerpoint.drawLine(bx, by, ex, ey)

  ##
  # @brief スライド番号の進める、戻す
  # @param self 
  # @param num 進めるスライド数
  def changeSlideNum(self, num):
    if self.SlideNumberInRelative[0] == 0:
      self.m_powerpoint.gotoSlide(num)
      self.slidenum = num
    else:
      
      self.m_powerpoint.gotoSlide(self.slidenum+num)
      self.slidenum += num


  ##
  # @brief アニメーションを進める、戻す
  # @param self 
  # @param num 進めるアニメーションの数
  def changeEffeceNum(self, num):
    if num > 0:
      for i in range(0, num):
        self.m_powerpoint.next()
    else:
      for i in range(0, -num):
        self.m_powerpoint.previous()

  ##
  # @brief スライド番号取得
  # @param self 
  # @return スライド番号
  def getSlideNum(self):
    return self.slidenum


  
  
  def onActivated(self, ec_id):
    ImpressControl.onActivated(self, ec_id)
    
    

    #self.file = open('text3.txt', 'w')

    self.m_powerpoint.initCom()
    self.m_powerpoint.run()
    self.m_powerpoint.gotoSlide(self.SlideFileInitialNumber[0])
    self.slidenum = self.SlideFileInitialNumber[0]
    
    return RTC.RTC_OK

  def onDeactivated(self, ec_id):
    ImpressControl.onDeactivated(self, ec_id)
    self.m_powerpoint.end()
    #self.file.close()
    return RTC.RTC_OK


  ##
  # @brief 周期処理用コールバック関数
  # @param self 
  # @param ec_id target ExecutionContext Id
  # @return RTC::ReturnCode_t
  
  def onExecute(self, ec_id):
    ImpressControl.onExecute(self, ec_id)
        

    return RTC.RTC_OK

  
  ##
  # @brief 終了処理用コールバック関数
  # @param self 
  # @param ec_id target ExecutionContext Id
  # @return RTC::ReturnCode_t
  
  def on_shutdown(self, ec_id):
      ImpressControl.onExecute(self, ec_id)
      return RTC.RTC_OK


  
  
  
      

  
##
# @class PowerPointObject
# @brief PowerPointを操作するクラス
#
class PowerPointObject:
    ppLayoutBlank = 12
    ppLayoutChart = 8
    ppLayoutChartAndText = 6
    ppLayoutClipartAndText = 10
    ppLayoutClipArtAndVerticalText = 26
    ppLayoutCustom = 32
    ppLayoutFourObjects = 24
    ppLayoutLargeObject = 15
    ppLayoutMediaClipAndText = 18
    ppLayoutMixed = -2
    ppLayoutObject = 16
    ppLayoutObjectAndText = 14
    ppLayoutObjectAndTwoObjects = 30
    ppLayoutObjectOverText = 19
    ppLayoutOrgchart = 7
    ppLayoutTable = 4
    ppLayoutText = 2
    ppLayoutTextAndChart = 5
    ppLayoutTextAndClipart = 9
    ppLayoutTextAndMediaClip = 17
    ppLayoutTextAndObject = 13
    ppLayoutTextAndTwoObjects = 21
    ppLayoutTextOverObject = 20
    ppLayoutTitle = 1
    ppLayoutTitleOnly = 11
    ppLayoutTwoColumnText = 3
    ppLayoutTwoObjects = 29
    ppLayoutTwoObjectsAndObject = 31
    ppLayoutTwoObjectsAndText = 22
    ppLayoutTwoObjectsOverText = 23
    ppLayoutVerticalText = 25
    ppLayoutVerticalTitleAndText = 27
    ppLayoutVerticalTitleAndTextOverChart = 28

    ##
    # @brief コンストラクタ
    # @param self 
    #
    def __init__(self):
        self.filename = " "
        
        self.ptApplication = None
        self.ptPresentations = None
        self.ptPresentation = None
        self.ptSlideShowWindow = None
        self.ptSlideShowView = None
        

        self.thread_ptApplication = None
        self.thread_ptPresentations = None
        self.thread_ptPresentation = None

        self.t_ptApplication = None
        self.t_ptPresentations = None
        self.t_ptPresentation = None

    ##
    # @brief 
    # @param self
    #
    def run(self):
      self.ptSlideShowWindow = self.ptPresentation.SlideShowSettings.Run()
      self.ptSlideShowView = self.ptSlideShowWindow.View
      
    ##
    # @brief 
    # @param self
    #
    def end(self):
      self.ptSlideShowView.Exit()
      self.ptSlideShowWindow = None
      self.ptSlideShowView = None
      
    ##
    # @brief 
    # @param self
    #
    def gotoSlide(self, num):
      
      if 0 < num and num <=  self.ptPresentation.Slides.Count:
        self.ptSlideShowView.GotoSlide(num)
        return True
      else:
        return False


        
    ##
    # @brief 
    # @param self 
    #
    def next(self):
      self.ptSlideShowView.Next()
    ##
    # @brief 
    # @param self 
    #
    def previous(self):
      self.ptSlideShowView.Previous()
    ##
    # @brief 
    # @param self
    # @param bx
    # @param by
    # @param ex
    # @param ey 
    #
    def drawLine(self, bx, by, ex, ey):
      
      self.ptSlideShowView.DrawLine(bx, by, ex, ey)

    ##
    # @brief 
    # @param self 
    #
    def eraseDrawing(self):
      self.ptSlideShowView.EraseDrawing()
    
    ##
    # @brief 
    # @param self 
    #
    def preInitCom(self):
        self.thread_ptApplication = pythoncom.CoMarshalInterThreadInterfaceInStream (pythoncom.IID_IDispatch, self.t_ptApplication)
        self.thread_ptPresentations = pythoncom.CoMarshalInterThreadInterfaceInStream (pythoncom.IID_IDispatch, self.t_ptPresentations)
        self.thread_ptPresentation = pythoncom.CoMarshalInterThreadInterfaceInStream (pythoncom.IID_IDispatch, self.t_ptPresentation)

    ##
    # @brief 
    # @param self 
    #
    def initCom(self):
        if self.ptApplication == None:
          pythoncom.CoInitialize()
          self.ptApplication = win32com.client.Dispatch ( pythoncom.CoGetInterfaceAndReleaseStream (self.thread_ptApplication, pythoncom.IID_IDispatch))
          self.ptPresentations = win32com.client.Dispatch ( pythoncom.CoGetInterfaceAndReleaseStream (self.thread_ptPresentations, pythoncom.IID_IDispatch))
          self.ptPresentation = win32com.client.Dispatch ( pythoncom.CoGetInterfaceAndReleaseStream (self.thread_ptPresentation, pythoncom.IID_IDispatch))

    ##
    # @brief 
    # @param self 
    #
    def closeCom(self):
        pythoncom.CoUninitialize()

    ##
    # @brief PowerPointファイルを開く関数
    # @param self 
    # @param fn ファイルパス
    #
    def Open(self, fn):
        if self.filename == fn:
            return
        self.filename = fn

        

        try:
            if self.ptApplication == None:
              t_ptApplication = win32com.client.Dispatch("PowerPoint.Application")
            else:
              t_ptApplication = self.ptApplication
            
            
            t_ptApplication.Visible = True
            try:
                t_ptPresentations = t_ptApplication.Presentations
                

                try:
                    t_ptPresentation = None
                    if self.filename == "":
                        t_ptPresentation = t_ptPresentations.Add()
                        t_ptPresentation.Slides.Add(1,PowerPointObject.ppLayoutTitleOnly)
                    else:
                        t_ptPresentation = t_ptPresentations.Open(self.filename)

                    
                    self.t_ptApplication = t_ptApplication
                    self.t_ptPresentations = t_ptPresentations
                    self.t_ptPresentation = t_ptPresentation

                    self.preInitCom()

                    
                except:
                    return
            except:
                return
        except:
            return


##
# @brief
# @param manager マネージャーオブジェクト
def MyModuleInit(manager):
    profile = OpenRTM_aist.Properties(defaults_str=powerpointcontrol_spec)
    manager.registerFactory(profile,
                            PowerPointControl,
                            OpenRTM_aist.Delete)
    comp = manager.createComponent("PowerPointControl")

def main():
    """po = PowerPointObject()
    fn = os.path.abspath("1.pptx")
    po.Open(fn)
    po.run()
    po.gotoSlide(2)
    return"""
    
    
    mgr = OpenRTM_aist.Manager.init(sys.argv)
    mgr.setModuleInitProc(MyModuleInit)
    mgr.activateManager()
    mgr.runManager()
    
if __name__ == "__main__":
    main()
