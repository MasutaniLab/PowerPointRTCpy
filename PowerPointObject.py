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



