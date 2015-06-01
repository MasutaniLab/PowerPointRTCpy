#!/bin/env python
# -*- encoding: utf-8 -*-

##
#
# @file ImpressControl.py

import optparse
import sys,os,platform
import re
import time
import random
import commands
import math



from os.path import expanduser
sv = sys.version_info


if os.name == 'posix':
    home = expanduser("~")
    sys.path += [home+'/OOoRTC', '/usr/lib/python2.' + str(sv[1]) + '/dist-packages', '/usr/lib/python2.' + str(sv[1]) + '/dist-packages/rtctree/rtmidl']
elif os.name == 'nt':
    sys.path += ['.\\OOoRTC', 'C:\\Python2' + str(sv[1]) + '\\lib\\site-packages', 'C:\\Python2' + str(sv[1]) + '\\Lib\\site-packages\\OpenRTM_aist\\RTM_IDL', 'C:\\Python2' + str(sv[1]) + '\\lib\\site-packages\\rtctree\\rtmidl']




import RTC
import OpenRTM_aist

from OpenRTM_aist import CorbaNaming
from OpenRTM_aist import RTObject
from OpenRTM_aist import CorbaConsumer
from omniORB import CORBA
import CosNaming





##
# @class ImpressControl
# @brief プレゼンテーションソフトを操作するためのRTCのクラス
#

class ImpressControl(OpenRTM_aist.DataFlowComponentBase):
    ##
    # @brief コンストラクタ
    # @param self 
    # @param manager マネージャーオブジェクト
    #
  def __init__(self, manager):
    OpenRTM_aist.DataFlowComponentBase.__init__(self, manager)
    
    self._d_m_SlideNumin = RTC.TimedShort(RTC.Time(0,0),0)
    self._m_SlideNumIn = OpenRTM_aist.InPort("SlideNumberIn", self._d_m_SlideNumin)

    self._d_m_EffectNum = RTC.TimedShort(RTC.Time(0,0),0)
    self._m_EffectNumIn = OpenRTM_aist.InPort("EffectNumberIn", self._d_m_EffectNum)
    
    self._d_m_SlideNumout = RTC.TimedShort(RTC.Time(0,0),0)
    self._m_SlideNumOut = OpenRTM_aist.OutPort("SlideNumberOut", self._d_m_SlideNumout)

    self._d_m_Pen = RTC.TimedShortSeq(RTC.Time(0,0),[])
    self._m_PenIn = OpenRTM_aist.InPort("Pen", self._d_m_Pen)


    self.SlideFileInitialNumber = [0]
    self.SlideNumberInRelative = [1]
    

    
    
    
    
    
    return

  

  ##
  # @brief 実行周期を設定する関数
  # @param self 
  # @param rate：実行周期
  def mSetRate(self, rate):
      m_ec = self.get_owned_contexts()
      m_ec[0].set_rate(rate)

  ##
  # @brief 活性化するための関数
  # @param self 
  def mActivate(self):
      m_ec = self.get_owned_contexts()
      m_ec[0].activate_component(self._objref)

  ##
  # @brief 不活性化するための関数
  # @param self 
  def mDeactivate(self):
      m_ec = self.get_owned_contexts()
      m_ec[0].deactivate_component(self._objref)

      

  

  
  ##
  # @brief 不活性化時のコールバック関数
  # @param self 
  # @param ec_id target ExecutionContext Id
  # @return RTC::ReturnCode_t
  def onDeactivated(self, ec_id):
    


    
    return RTC.RTC_OK

  ##
  # @brief 活性化処理用コールバック関数
  # @param self 
  # @param ec_id target ExecutionContext Id
  # @return RTC::ReturnCode_t
  def onActivated(self, ec_id):
    
    return RTC.RTC_OK
  
  ##
  # @brief 初期化処理用コールバック関数
  # @param self 
  # @return RTC::ReturnCode_t
  def onInitialize(self):

    
    
    self.addInPort("SlideNumberIn",self._m_SlideNumIn)
    self.addInPort("EffectNumberIn",self._m_EffectNumIn)
    self.addOutPort("SlideNumberOut",self._m_SlideNumOut)
    self.addOutPort("Pen",self._m_PenIn)

    self.bindParameter("SlideFileInitialNumber", self.SlideFileInitialNumber, "0")
    self.bindParameter("SlideNumberInRelative", self.SlideNumberInRelative, "1")
    
    
    return RTC.RTC_OK


  ##
  # @brief 線描画
  # @param self 
  # @param bx 
  # @param by 
  # @param ex 
  # @param ey 
  def drawLine(self, bx, by, ex, ey):
    pass

  ##
  # @brief スライド番号の進める、戻す
  # @param self 
  # @param num 進めるスライド数
  def changeSlideNum(self, num):
    pass

  ##
  # @brief アニメーションを進める、戻す
  # @param self 
  # @param num 進めるアニメーションの数
  def changeEffeceNum(self, num):
    pass

  ##
  # @brief スライド番号取得
  # @param self 
  # @return スライド番号
  def getSlideNum(self):
    return 0

  ##
  # @brief 周期処理用コールバック関数
  # @param self 
  # @param ec_id target ExecutionContext Id
  # @return RTC::ReturnCode_t
  
  def onExecute(self, ec_id): 
    if self._m_SlideNumIn.isNew():
      data = self._m_SlideNumIn.read()
      self.changeSlideNum(data.data)

      self._d_m_SlideNumout.data = self.getSlideNum()
      self._m_SlideNumOut.write()
    

    if self._m_EffectNumIn.isNew():
      data = self._m_EffectNumIn.read()
      self.changeEffeceNum(data.data)
        
    if self._m_PenIn.isNew():
      data = self._m_PenIn.read()
      
      for i in range(0, int(len(data.data)/2)):
        if i*2+1 < len(data.data):
            if i == 0:
              self.drawLine(data.data[i*2],data.data[i*2+1],data.data[i*2],data.data[i*2+1])
            else:
              self.drawLine(data.data[i*2-2],data.data[i*2-1],data.data[i*2],data.data[i*2+1])

    

    return RTC.RTC_OK

  ##
  # @brief 終了処理用コールバック関数
  # @param self 
  # @param 
  # @return RTC::ReturnCode_t
  
  def onFinalize(self):
      return RTC.RTC_OK
