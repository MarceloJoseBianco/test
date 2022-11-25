# -*- coding: utf-8 -*-
"""
Created on Fri Mar 13 07:28:55 2020

@author: fiw
"""

#from LoadSCDMAPITypesV22 import*
#from UtilitiesOnLoadV22 import*
import pickle
import os
#from myTools import *
import math
import time
import ctypes
#from PIL import Image
#from tkinter import *  
import sys
#from Tkinter import *  


#from SCTools import*




import clr

clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")



from System.Windows.Forms import Application, Form, TextBox, Label, Button, CheckBox, BorderStyle, Panel, ComboBox
from System.Windows.Forms import ToolBar, ToolBarButton, OpenFileDialog, FolderBrowserDialog
from System.Windows.Forms import DialogResult, ScrollBars, DockStyle, PictureBox

from System.Drawing import Point, Color, Size, Image



def openWindowC1Geo(token=''):
    
    #timeSteps=(0,1,2,3)
    form = getInputGeo()
    form.AutoSize=True
    #form.Size.Height=500
    #form.Size.Width=1000
    #form = IForm()
    form.ShowDialog()
    
    return True
    
class getInputGeo(Form):


    def __init__(self):
            #
        self.CenterToScreen()
        self.geoDic={}

#        self.pb=PictureBox()
#        self.pb.Size = Size(188, 200)
#        self.pb.Location = Point(300, 0)
#        self.pb.Image = Image.FromFile(r"C:\Users\filip.wegener\OneDrive - JBO GmbH\dropbox\Dropbox\Data\Ansys\Extensions\C1Tool\C1Tool\images\upperFlangePara_klein.png") 
#        self.Controls.Add(self.pb)

        
        ### Upper Flange 
        col=0
#        self.L_upperFlange = Label()
#        self.L_upperFlange.Text = "Upper Flange"
#        self.L_upperFlange.Location = Point(20, col)
#        self.L_upperFlange.Height = 20
#        self.L_upperFlange.Width = 100
#        self.Controls.Add(self.L_upperFlange)
        #Create Geo button 
        self.openSketchUpperFlange = Button()
        self.openSketchUpperFlange.Text = "Upper Flange"
        self.openSketchUpperFlange.Location = Point(20, col)
        self.openSketchUpperFlange.Height = 22
        self.openSketchUpperFlange.Width = 100      
        self.openSketchUpperFlange.Click += self.showSketchUpperFlange
        self.Controls.Add(self.openSketchUpperFlange)        

        #flange_w
        col+=25
        self.L_upperFlange_w = Label()
        self.L_upperFlange_w.Text = "flange_w [mm]"
        self.L_upperFlange_w.Location = Point(50, col)
        self.L_upperFlange_w.Height = 20
        self.L_upperFlange_w.Width = 100
        self.Controls.Add(self.L_upperFlange_w)        

        self.TB_upperFlange_w = TextBox()
        self.TB_upperFlange_w.Text ='285'
        self.TB_upperFlange_w.Location = Point(200, col)
        self.TB_upperFlange_w.Height = 10
        self.TB_upperFlange_w.Width = 50
        self.Controls.Add(self.TB_upperFlange_w)   

        #flange_h
        col+=20
        self.L_upperFlange_h = Label()
        self.L_upperFlange_h.Text = "flange_h [mm]"
        self.L_upperFlange_h.Location = Point(50, col)
        self.L_upperFlange_h.Height = 20
        self.L_upperFlange_h.Width = 100
        self.Controls.Add(self.L_upperFlange_h)        

        self.TB_upperFlange_h = TextBox()
        self.TB_upperFlange_h.Text ='128'
        self.TB_upperFlange_h.Location = Point(200, col)
        self.TB_upperFlange_h.Height = 10
        self.TB_upperFlange_h.Width = 50
        self.Controls.Add(self.TB_upperFlange_h)  

        #web_h
        col+=20
        self.L_web_h = Label()
        self.L_web_h.Text = "web_h [mm]"
        self.L_web_h.Location = Point(50, col)
        self.L_web_h.Height = 20
        self.L_web_h.Width = 100
        self.Controls.Add(self.L_web_h)        

        self.TB_web_h = TextBox()
        self.TB_web_h.Text ='363'
        self.TB_web_h.Location = Point(200, col)
        self.TB_web_h.Height = 10
        self.TB_web_h.Width = 50
        self.Controls.Add(self.TB_web_h) 

        #web_w
        col+=20
        self.L_web_w = Label()
        self.L_web_w.Text = "web_w [mm]"
        self.L_web_w.Location = Point(50, col)
        self.L_web_w.Height = 20
        self.L_web_w.Width = 100
        self.Controls.Add(self.L_web_w)        

        self.TB_web_w = TextBox()
        self.TB_web_w.Text ='75'
        self.TB_web_w.Location = Point(200, col)
        self.TB_web_w.Height = 10
        self.TB_web_w.Width = 50
        self.Controls.Add(self.TB_web_w) 

        #neck_s
        col+=20
        self.L_upperFlange_neck_s = Label()
        self.L_upperFlange_neck_s.Text = "neck_s [mm]"
        self.L_upperFlange_neck_s.Location = Point(50, col)
        self.L_upperFlange_neck_s.Height = 20
        self.L_upperFlange_neck_s.Width = 100
        self.Controls.Add(self.L_upperFlange_neck_s)        

        self.TB_upperFlange_neck_s = TextBox()
        self.TB_upperFlange_neck_s.Text ='100'
        self.TB_upperFlange_neck_s.Location = Point(200, col)
        self.TB_upperFlange_neck_s.Height = 10
        self.TB_upperFlange_neck_s.Width = 50
        self.Controls.Add(self.TB_upperFlange_neck_s)  

        #neck_r
        col+=20
        self.L_upperFlange_neck_r = Label()
        self.L_upperFlange_neck_r.Text = "neck_r [mm]"
        self.L_upperFlange_neck_r.Location = Point(50, col)
        self.L_upperFlange_neck_r.Height = 20
        self.L_upperFlange_neck_r.Width = 100
        self.Controls.Add(self.L_upperFlange_neck_r)        

        self.TB_upperFlange_neck_r = TextBox()
        self.TB_upperFlange_neck_r.Text ='15'
        self.TB_upperFlange_neck_r.Location = Point(200, col)
        self.TB_upperFlange_neck_r.Height = 10
        self.TB_upperFlange_neck_r.Width = 50
        self.Controls.Add(self.TB_upperFlange_neck_r)         

        #neck_h
        col+=20
        self.L_upperFlange_neck_h = Label()
        self.L_upperFlange_neck_h.Text = "neck_h [mm]"
        self.L_upperFlange_neck_h.Location = Point(50, col)
        self.L_upperFlange_neck_h.Height = 20
        self.L_upperFlange_neck_h.Width = 100
        self.Controls.Add(self.L_upperFlange_neck_h)        

        self.TB_upperFlange_neck_h = TextBox()
        self.TB_upperFlange_neck_h.Text ='200'
        self.TB_upperFlange_neck_h.Location = Point(200, col)
        self.TB_upperFlange_neck_h.Height = 10
        self.TB_upperFlange_neck_h.Width = 50
        self.Controls.Add(self.TB_upperFlange_neck_h)         


        
#        self.flangeGeoTop['cutOut_w']  = 115
#        self.flangeGeoTop['cutOut_topL']  = 157
#        self.flangeGeoTop['cutOut_h']  = 287
#        self.flangeGeoTop['cutOut_topR']  = 30
#        self.flangeGeoTop['cutOut_midR']  = 300
#        self.flangeGeoTop['cutOut_botR']  = 50
        
        ### upper cutOut
        col+=20
        self.L_upperCutOut = Label()
        self.L_upperCutOut.Text = "Upper Flange cutOut"
        self.L_upperCutOut.Location = Point(20, col)
        self.L_upperCutOut.Height = 20
        self.L_upperCutOut.Width = 200
        self.Controls.Add(self.L_upperCutOut)
        #
#        self.openSketchUpperCutOut = Button()
#        self.openSketchUpperCutOut.Text = "Upper Flange cutOut"
#        self.openSketchUpperCutOut.Location = Point(20, col)
#        self.openSketchUpperCutOut.Height = 22
#        self.openSketchUpperCutOut.Width = 150      
#        self.openSketchUpperCutOut.Click += self.showSketchUpperCutOut
#        self.Controls.Add(self.openSketchUpperCutOut)           

        #cutOut_w
        col+=25
        self.L_upperCutOut_w = Label()
        self.L_upperCutOut_w.Text = "cutOut w [mm]"
        self.L_upperCutOut_w.Location = Point(50, col)
        self.L_upperCutOut_w.Height = 20
        self.L_upperCutOut_w.Width = 100
        self.Controls.Add(self.L_upperCutOut_w)        

        self.TB_upperCutOut_w = TextBox()
        self.TB_upperCutOut_w.Text ='115'
        self.TB_upperCutOut_w.Location = Point(200, col)
        self.TB_upperCutOut_w.Height = 10
        self.TB_upperCutOut_w.Width = 50
        self.Controls.Add(self.TB_upperCutOut_w)  

        #cutOut_h
        col+=20
        self.L_upperCutOut_h = Label()
        self.L_upperCutOut_h.Text = "cutOut h [mm]"
        self.L_upperCutOut_h.Location = Point(50, col)
        self.L_upperCutOut_h.Height = 20
        self.L_upperCutOut_h.Width = 100
        self.Controls.Add(self.L_upperCutOut_h)        

        self.TB_upperCutOut_h = TextBox()
        self.TB_upperCutOut_h.Text ='287'
        self.TB_upperCutOut_h.Location = Point(200, col)
        self.TB_upperCutOut_h.Height = 10
        self.TB_upperCutOut_h.Width = 50
        self.Controls.Add(self.TB_upperCutOut_h)   

        #cutOut_E1
        col+=20
        self.L_upperCutOut_E1 = Label()
        self.L_upperCutOut_E1.Text = "cutOut E1 [mm]"
        self.L_upperCutOut_E1.Location = Point(50, col)
        self.L_upperCutOut_E1.Height = 20
        self.L_upperCutOut_E1.Width = 100
        self.Controls.Add(self.L_upperCutOut_E1)        

        self.TB_upperCutOut_E1 = TextBox()
        self.TB_upperCutOut_E1.Text ='20'
        self.TB_upperCutOut_E1.Location = Point(200, col)
        self.TB_upperCutOut_E1.Height = 10
        self.TB_upperCutOut_E1.Width = 50
        self.Controls.Add(self.TB_upperCutOut_E1)                
        
        #cutOut_E2
        col+=20
        self.L_upperCutOut_E2 = Label()
        self.L_upperCutOut_E2.Text = "cutOut E2 [mm]"
        self.L_upperCutOut_E2.Location = Point(50, col)
        self.L_upperCutOut_E2.Height = 20
        self.L_upperCutOut_E2.Width = 100
        self.Controls.Add(self.L_upperCutOut_E2)        

        self.TB_upperCutOut_E2 = TextBox()
        self.TB_upperCutOut_E2.Text ='40'
        self.TB_upperCutOut_E2.Location = Point(200, col)
        self.TB_upperCutOut_E2.Height = 10
        self.TB_upperCutOut_E2.Width = 50
        self.Controls.Add(self.TB_upperCutOut_E2)     
        
        #cutOut_midR
        col+=20
        self.L_upperCutOut_midR = Label()
        self.L_upperCutOut_midR.Text = "cutOut midR [mm]"
        self.L_upperCutOut_midR.Location = Point(50, col)
        self.L_upperCutOut_midR.Height = 20
        self.L_upperCutOut_midR.Width = 100
        self.Controls.Add(self.L_upperCutOut_midR)        

        self.TB_upperCutOut_midR = TextBox()
        self.TB_upperCutOut_midR.Text ='300'
        self.TB_upperCutOut_midR.Location = Point(200, col)
        self.TB_upperCutOut_midR.Height = 10
        self.TB_upperCutOut_midR.Width = 50
        self.Controls.Add(self.TB_upperCutOut_midR) 

        #cutOut_botR
        col+=20
        self.L_upperCutOut_botR = Label()
        self.L_upperCutOut_botR.Text = "cutOut botR [mm]"
        self.L_upperCutOut_botR.Location = Point(50, col)
        self.L_upperCutOut_botR.Height = 20
        self.L_upperCutOut_botR.Width = 100
        self.Controls.Add(self.L_upperCutOut_botR)        

        self.TB_upperCutOut_botR = TextBox()
        self.TB_upperCutOut_botR.Text ='50'
        self.TB_upperCutOut_botR.Location = Point(200, col)
        self.TB_upperCutOut_botR.Height = 10
        self.TB_upperCutOut_botR.Width = 50
        self.Controls.Add(self.TB_upperCutOut_botR) 
        
        ### Upper Sheets 
#        col+=20
#        self.L_upperSheets = Label()
#        self.L_upperSheets.Text = "Upper Sheets"
#        self.L_upperSheets.Location = Point(20, col)
#        self.L_upperSheets.Height = 20
#        self.L_upperSheets.Width = 200
#        self.Controls.Add(self.L_upperSheets)
#
#        #turmTop_s
#        col+=20
#        self.L_turmTop_s = Label()
#        self.L_turmTop_s.Text = "upper sheet s [mm]"
#        self.L_turmTop_s.Location = Point(50, col)
#        self.L_turmTop_s.Height = 20
#        self.L_turmTop_s.Width = 100
#        self.Controls.Add(self.L_turmTop_s)        
#
#        self.TB_turmTop_s = TextBox()
#        self.TB_turmTop_s.Text ='100'
#        self.TB_turmTop_s.Location = Point(200, col)
#        self.TB_turmTop_s.Height = 10
#        self.TB_turmTop_s.Width = 50
#        self.Controls.Add(self.TB_turmTop_s)  
#
#        #turmTop_l
#        col+=20
#        self.L_turmTop_l = Label()
#        self.L_turmTop_l.Text = "upper sheet l [mm]"
#        self.L_turmTop_l.Location = Point(50, col)
#        self.L_turmTop_l.Height = 20
#        self.L_turmTop_l.Width = 100
#        self.Controls.Add(self.L_turmTop_l)        
#
#        self.TB_turmTop_l = TextBox()
#        self.TB_turmTop_l.Text ='1000'
#        self.TB_turmTop_l.Location = Point(200, col)
#        self.TB_turmTop_l.Height = 10
#        self.TB_turmTop_l.Width = 50
#        self.Controls.Add(self.TB_turmTop_l)  


        ### Lower Flange 
        col+=20
#        self.L_lowerFlange = Label()
#        self.L_lowerFlange.Text = "Lower Flange"
#        self.L_lowerFlange.Location = Point(20, col)
#        self.L_lowerFlange.Height = 20
#        self.L_lowerFlange.Width = 200
#        self.Controls.Add(self.L_lowerFlange)
        
        #Create Geo button 
        self.openSketchLowerFlange = Button()
        self.openSketchLowerFlange.Text = "Lower Flange"
        self.openSketchLowerFlange.Location = Point(20, col)
        self.openSketchLowerFlange.Height = 22
        self.openSketchLowerFlange.Width = 100      
        self.openSketchLowerFlange.Click += self.showSketchLowerFlange
        self.Controls.Add(self.openSketchLowerFlange)        
        
        #flange_w
        col+=25
        self.L_lowerFlange_w = Label()
        self.L_lowerFlange_w.Text = "flange_w [mm]"
        self.L_lowerFlange_w.Location = Point(50, col)
        self.L_lowerFlange_w.Height = 20
        self.L_lowerFlange_w.Width = 100
        self.Controls.Add(self.L_lowerFlange_w)        

        self.TB_lowerFlange_w = TextBox()
        self.TB_lowerFlange_w.Text ='120'
        self.TB_lowerFlange_w.Location = Point(200, col)
        self.TB_lowerFlange_w.Height = 10
        self.TB_lowerFlange_w.Width = 50
        self.Controls.Add(self.TB_lowerFlange_w)         

        #flange_h
        col+=20
        self.L_lowerFlange_h = Label()
        self.L_lowerFlange_h.Text = "flange_h [mm]"
        self.L_lowerFlange_h.Location = Point(50, col)
        self.L_lowerFlange_h.Height = 20
        self.L_lowerFlange_h.Width = 100
        self.Controls.Add(self.L_lowerFlange_h)        

        self.TB_lowerFlange_h = TextBox()
        self.TB_lowerFlange_h.Text ='500'
        self.TB_lowerFlange_h.Location = Point(200, col)
        self.TB_lowerFlange_h.Height = 10
        self.TB_lowerFlange_h.Width = 50
        self.Controls.Add(self.TB_lowerFlange_h) 
        
        #flange_topL
        col+=20
        self.L_lowerFlange_topL = Label()
        self.L_lowerFlange_topL.Text = "flange_topL [mm]"
        self.L_lowerFlange_topL.Location = Point(50, col)
        self.L_lowerFlange_topL.Height = 20
        self.L_lowerFlange_topL.Width = 100
        self.Controls.Add(self.L_lowerFlange_topL)        

        self.TB_lowerFlange_topL = TextBox()
        self.TB_lowerFlange_topL.Text ='85'
        self.TB_lowerFlange_topL.Location = Point(200, col)
        self.TB_lowerFlange_topL.Height = 10
        self.TB_lowerFlange_topL.Width = 50
        self.Controls.Add(self.TB_lowerFlange_topL)         

        #lowerFlange_neck_s
        col+=20
        self.L_lowerFlange_neck_s = Label()
        self.L_lowerFlange_neck_s.Text = "neck_s [mm]"
        self.L_lowerFlange_neck_s.Location = Point(50, col)
        self.L_lowerFlange_neck_s.Height = 20
        self.L_lowerFlange_neck_s.Width = 100
        self.Controls.Add(self.L_lowerFlange_neck_s)        

        self.TB_lowerFlange_neck_s = TextBox()
        self.TB_lowerFlange_neck_s.Text ='100'
        self.TB_lowerFlange_neck_s.Location = Point(200, col)
        self.TB_lowerFlange_neck_s.Height = 10
        self.TB_lowerFlange_neck_s.Width = 50
        self.Controls.Add(self.TB_lowerFlange_neck_s)  

        #neck_h
        col+=20
        self.L_lowerFlange_neck_h = Label()
        self.L_lowerFlange_neck_h.Text = "neck_h [mm]"
        self.L_lowerFlange_neck_h.Location = Point(50, col)
        self.L_lowerFlange_neck_h.Height = 20
        self.L_lowerFlange_neck_h.Width = 100
        self.Controls.Add(self.L_lowerFlange_neck_h)        

        self.TB_lowerFlange_neck_h = TextBox()
        self.TB_lowerFlange_neck_h.Text ='30'
        self.TB_lowerFlange_neck_h.Location = Point(200, col)
        self.TB_lowerFlange_neck_h.Height = 10
        self.TB_lowerFlange_neck_h.Width = 50
        self.Controls.Add(self.TB_lowerFlange_neck_h)         

        ### lower cutOut
        col+=20
        self.L_lowerCutOut = Label()
        self.L_lowerCutOut.Text = "lower Flange cutOut"
        self.L_lowerCutOut.Location = Point(20, col)
        self.L_lowerCutOut.Height = 20
        self.L_lowerCutOut.Width = 200
        self.Controls.Add(self.L_lowerCutOut)
        
#        self.openSketchLowerCutOut = Button()
#        self.openSketchLowerCutOut.Text = "Lower Flange cutOut"
#        self.openSketchLowerCutOut.Location = Point(20, col)
#        self.openSketchLowerCutOut.Height = 22
#        self.openSketchLowerCutOut.Width = 150      
#        self.openSketchLowerCutOut.Click += self.showSketchLowerCutOut
#        self.Controls.Add(self.openSketchLowerCutOut)              

        #cutOut_w
        col+=20
        self.L_lowerCutOut_w = Label()
        self.L_lowerCutOut_w.Text = "CutOut w [mm]"
        self.L_lowerCutOut_w.Location = Point(50, col)
        self.L_lowerCutOut_w.Height = 20
        self.L_lowerCutOut_w.Width = 100
        self.Controls.Add(self.L_lowerCutOut_w)        

        self.TB_lowerCutOut_w = TextBox()
        self.TB_lowerCutOut_w.Text ='115'
        self.TB_lowerCutOut_w.Location = Point(200, col)
        self.TB_lowerCutOut_w.Height = 10
        self.TB_lowerCutOut_w.Width = 50
        self.Controls.Add(self.TB_lowerCutOut_w)  

        #cutOut_h
        col+=20
        self.L_lowerCutOut_h = Label()
        self.L_lowerCutOut_h.Text = "cutOut h [mm]"
        self.L_lowerCutOut_h.Location = Point(50, col)
        self.L_lowerCutOut_h.Height = 20
        self.L_lowerCutOut_h.Width = 100
        self.Controls.Add(self.L_lowerCutOut_h)        

        self.TB_lowerCutOut_h = TextBox()
        self.TB_lowerCutOut_h.Text ='200'
        self.TB_lowerCutOut_h.Location = Point(200, col)
        self.TB_lowerCutOut_h.Height = 10
        self.TB_lowerCutOut_h.Width = 50
        self.Controls.Add(self.TB_lowerCutOut_h)   

        #cutOut_E1
        col+=20
        self.L_lowerCutOut_E1 = Label()
        self.L_lowerCutOut_E1.Text = "cutOut E1 [mm]"
        self.L_lowerCutOut_E1.Location = Point(50, col)
        self.L_lowerCutOut_E1.Height = 20
        self.L_lowerCutOut_E1.Width = 100
        self.Controls.Add(self.L_lowerCutOut_E1)        

        self.TB_lowerCutOut_E1 = TextBox()
        self.TB_lowerCutOut_E1.Text ='80'
        self.TB_lowerCutOut_E1.Location = Point(200, col)
        self.TB_lowerCutOut_E1.Height = 10
        self.TB_lowerCutOut_E1.Width = 50
        self.Controls.Add(self.TB_lowerCutOut_E1)                
        
        #cutOut_midR
        col+=20
        self.L_lowerCutOut_midR = Label()
        self.L_lowerCutOut_midR.Text = "cutOut midR [mm]"
        self.L_lowerCutOut_midR.Location = Point(50, col)
        self.L_lowerCutOut_midR.Height = 20
        self.L_lowerCutOut_midR.Width = 100
        self.Controls.Add(self.L_lowerCutOut_midR)        

        self.TB_lowerCutOut_midR = TextBox()
        self.TB_lowerCutOut_midR.Text ='290'
        self.TB_lowerCutOut_midR.Location = Point(200, col)
        self.TB_lowerCutOut_midR.Height = 10
        self.TB_lowerCutOut_midR.Width = 50
        self.Controls.Add(self.TB_lowerCutOut_midR) 

        #cutOut_topR
        col+=20
        self.L_lowerCutOut_topR = Label()
        self.L_lowerCutOut_topR.Text = "cutOut botR [mm]"
        self.L_lowerCutOut_topR.Location = Point(50, col)
        self.L_lowerCutOut_topR.Height = 20
        self.L_lowerCutOut_topR.Width = 100
        self.Controls.Add(self.L_lowerCutOut_topR)        

        self.TB_lowerCutOut_topR = TextBox()
        self.TB_lowerCutOut_topR.Text ='50'
        self.TB_lowerCutOut_topR.Location = Point(200, col)
        self.TB_lowerCutOut_topR.Height = 10
        self.TB_lowerCutOut_topR.Width = 50
        self.Controls.Add(self.TB_lowerCutOut_topR) 

        ### Lower Sheets 
#        col+=20
#        self.L_lowerSheets = Label()
#        self.L_lowerSheets.Text = "Lower Sheets"
#        self.L_lowerSheets.Location = Point(20, col)
#        self.L_lowerSheets.Height = 20
#        self.L_lowerSheets.Width = 200
#        self.Controls.Add(self.L_lowerSheets)
#
#        #turmBot_s
#        col+=20
#        self.L_turmBot_s = Label()
#        self.L_turmBot_s.Text = "lower sheet s [mm]"
#        self.L_turmBot_s.Location = Point(50, col)
#        self.L_turmBot_s.Height = 20
#        self.L_turmBot_s.Width = 100
#        self.Controls.Add(self.L_turmBot_s)        
#
#        self.TB_turmBot_s = TextBox()
#        self.TB_turmBot_s.Text ='100'
#        self.TB_turmBot_s.Location = Point(200, col)
#        self.TB_turmBot_s.Height = 10
#        self.TB_turmBot_s.Width = 50
#        self.Controls.Add(self.TB_turmBot_s)  
#
#        #turmBot_l
#        col+=20
#        self.L_turmBot_l = Label()
#        self.L_turmBot_l.Text = "lower sheet l [mm]"
#        self.L_turmBot_l.Location = Point(50, col)
#        self.L_turmBot_l.Height = 20
#        self.L_turmBot_l.Width = 100
#        self.Controls.Add(self.L_turmBot_l)        
#
#        self.TB_turmBot_l = TextBox()
#        self.TB_turmBot_l.Text ='1000'
#        self.TB_turmBot_l.Location = Point(200, col)
#        self.TB_turmBot_l.Height = 10
#        self.TB_turmBot_l.Width = 50
#        self.Controls.Add(self.TB_turmBot_l)  

        ### Segment Options 
        col+=20
        self.L_Segment = Label()
        self.L_Segment.Text = "Segment"
        self.L_Segment.Location = Point(20, col)
        self.L_Segment.Height = 20
        self.L_Segment.Width = 200
        self.Controls.Add(self.L_Segment)     

        #segment_angle
        col+=20
        self.L_segment_angle = Label()
        self.L_segment_angle.Text = "segment angle [°]"
        self.L_segment_angle.Location = Point(50, col)
        self.L_segment_angle.Height = 20
        self.L_segment_angle.Width = 100
        self.Controls.Add(self.L_segment_angle)        

        self.TB_segment_angle = TextBox()
        self.TB_segment_angle.Text ='30'
        self.TB_segment_angle.Location = Point(200, col)
        self.TB_segment_angle.Height = 10
        self.TB_segment_angle.Width = 50
        self.Controls.Add(self.TB_segment_angle) 

        #tower_Rm
        col+=20
        self.L_segment_Rm = Label()
        self.L_segment_Rm.Text = "tower sheet Rm [mm]"
        self.L_segment_Rm.Location = Point(50, col)
        self.L_segment_Rm.Height = 20
        self.L_segment_Rm.Width = 150
        self.Controls.Add(self.L_segment_Rm)        

        self.TB_segment_Rm = TextBox()
        self.TB_segment_Rm.Text ='4000'
        self.TB_segment_Rm.Location = Point(200, col)
        self.TB_segment_Rm.Height = 10
        self.TB_segment_Rm.Width = 50
        self.Controls.Add(self.TB_segment_Rm)          


        #anzahl fastener
        col+=20
        self.L_number_fastener = Label()
        self.L_number_fastener.Text = "number fastener"
        self.L_number_fastener.Location = Point(50, col)
        self.L_number_fastener.Height = 20
        self.L_number_fastener.Width = 100
        self.Controls.Add(self.L_number_fastener)        

        self.TB_number_fastener = TextBox()
        self.TB_number_fastener.Text ='10'
        self.TB_number_fastener.Location = Point(200, col)
        self.TB_number_fastener.Height = 10
        self.TB_number_fastener.Width = 50
        self.Controls.Add(self.TB_number_fastener) 
        
#        # geoInput
#        self.label2 = Label()
#        self.label2.Text = "Input Files Geometry"
#        self.label2.Location = Point(20, 50)
#        self.label2.Height = 30
#        self.label2.Width = 200
#        self.Controls.Add(self.label2)

        ### Create Geo button 
        col+=20
        self.buttonCreate = Button()
        self.buttonCreate.Text = "Create Geometry"
        self.buttonCreate.Location = Point(20, col)
        self.buttonCreate.Height = 30
        self.buttonCreate.Width = 200      
        self.buttonCreate.Click += self.createGeo
        self.Controls.Add(self.buttonCreate)
#     
#        #
#
#        self.label2 = Label()
#        self.label2.Text = "Define start and end time for Interpolation:"
#        self.label2.Location = Point(20, 90)
#        self.label2.Height = 20
#        self.label2.Width = 250
#        self.Controls.Add(self.label2)
#        #
#        
#        self.label3 = Label()
#        self.label3.Text = "start time:"
#        self.label3.Location = Point(20, 120)
#        self.label3.Height = 20
#        self.label3.Width = 60
#        self.Controls.Add(self.label3)
#        #       
#        self.startTime = ComboBox()
#        self.startTime.Parent = self
#        self.startTime.Location = Point(80, 120)        
#        self.startTime.Items.AddRange(timeSteps)  
#        #
#
#        self.label4 = Label()
#        self.label4.Text = "ende time:"
#        self.label4.Location = Point(20, 150)
#        self.label4.Height = 20
#        self.label4.Width = 60
#        self.Controls.Add(self.label4)
#        #        
#        self.endTime = ComboBox()
#        self.endTime.Parent = self
#        self.endTime.Location = Point(80, 150)        
#        self.endTime.Items.AddRange(timeSteps) 
#        #
#        self.check = CheckBox()
#        self.check.Text = "Bolts_DC36"
#        self.check.Location = Point(20, 180)
#        self.check.Height = 20
#        self.check.Width = 125
#        self.Controls.Add(self.check) 
#        
#        #
#        self.check2 = CheckBox()
#        self.check2.Text = "Bolts_DC50"
#        self.check2.Location = Point(150, 180)
#        self.check2.Height = 20
#        self.check2.Width = 100
#        self.Controls.Add(self.check2)         
#        #
#        self.check3 = CheckBox()
#        self.check3.Text = "Weld_inside"
#        self.check3.Location = Point(20, 200)
#        self.check3.Height = 20
#        self.check3.Width = 125
#        self.Controls.Add(self.check3)  
#        #
#        self.check4 = CheckBox()
#        self.check4.Text = "Weld_outside"
#        self.check4.Location = Point(150, 200)
#        self.check4.Height = 20
#        self.check4.Width = 100
#        self.Controls.Add(self.check4)      
#        #
#        self.check5 = CheckBox()
#        self.check5.Text = "Flange_rounding"
#        self.check5.Location = Point(20, 220)
#        self.check5.Height = 20
#        self.check5.Width = 125
#        self.Controls.Add(self.check5)   
#        
#        self.check6 = CheckBox()
#        self.check6.Text = "Bolt_holes"
#        self.check6.Location = Point(150, 220)
#        self.check6.Height = 20
#        self.check6.Width = 100
#        self.Controls.Add(self.check6)    
#        
#        self.check7 = CheckBox()
#        self.check7.Text = "getGaps"
#        self.check7.Location = Point(20, 240)
#        self.check7.Height = 20
#        self.check7.Width = 125
#        self.Controls.Add(self.check7)   
#        
#        self.button4 = Button()
#        self.button4.Text = "Calc FLS"
#        self.button4.Location = Point(20, 260)
#        self.button4.Height = 30
#        self.button4.Width = 200    
#        #self.button4.BackColor = Color.Red
#        self.button4.Click += self.returnInput
#        self.Controls.Add(self.button4)  

    def showSketchUpperFlange(self, sender, event):
        form = showTopFlange()
        #form.AutoSize=True
        #form.Size.Height=500
        #form.Size.Width=1000
        #form = IForm()
        
        #form.ShowDialog()
        
        #file_name = raw_input("File Name: ")
        #def setup(self, filename, title):
        # adjust the form's client area size to the picture
        #   self.ClientSize = Size(300, 300)
        #   self.Text = title
        #  self.filename = filename
        #   self.image = Image.FromFile(self.filename)
        #   pictureBox = PictureBox()
           # this will fit the image to the form
       #    pictureBox.SizeMode = PictureBoxSizeMode.StretchImage
       #    pictureBox.Image = self.image
           # fit the picture box to the frame
       #    pictureBox.Dock = DockStyle.Fill
            
       #     self.Controls.Add(pictureBox)
       #   self.Show()
           
        #MessageBox.Show("High five! ExtSample1 is a success!")
        #print('getcwd:      ', os.getcwd())
        #print('__file__:    ', __file__)
        #os.startfile(file_name)
        #os.startfile("O:\\Users\\Marcelo\\C1 Connection\\Source Code\\C1Tool\\images\\upperFlangePara.jpg") 
        
        #import cv2 #EDIT, this line added

        #im = Image.open('O:\Users\Marcelo\C1 Connection\Source Code\C1Tool\images\upperFlangePara.jpg',0)
        #im.show('image',img)

    def showSketchLowerFlange(self, sender, event):
        os.startfile(r"C:\Users\filip.wegener\OneDrive - JBO GmbH\dropbox\Dropbox\Data\Ansys\Extensions\C1Tool\C1Tool\images\lowerFlangePara.jpg") 
        

        
    def createGeo(self, sender, event):
        
        ### upper Flange
        self.flangeGeoTop={}
        self.turmGeoTop={}
        self.flangeGeoTop['side']='top'
        self.turmGeoTop['side']='top' 
        
        self.flangeGeoTop['neck_s'] = float(self.TB_upperFlange_neck_s.Text.replace(',','.'))
        self.flangeGeoTop['neck_r']  = float(self.TB_upperFlange_neck_r.Text.replace(',','.'))
        self.flangeGeoTop['neck_h']  = float(self.TB_upperFlange_neck_h.Text.replace(',','.'))
        
        self.flangeGeoTop['flange_h']  = float(self.TB_upperFlange_h.Text.replace(',','.'))   
        self.flangeGeoTop['flange_w']  = float(self.TB_upperFlange_w.Text.replace(',','.'))   
        
        self.flangeGeoTop['fork_h']  = float(self.TB_web_h.Text.replace(',','.')) 
        self.flangeGeoTop['fork_w']  = float(self.TB_web_w.Text.replace(',','.')) 
        
#        self.turmGeoTop['s'] = float(self.TB_turmTop_s.Text.replace(',','.')) 
#        self.turmGeoTop['l'] = float(self.TB_turmTop_l.Text.replace(',','.')) 
        self.turmGeoTop['Rm'] = float(self.TB_segment_Rm.Text.replace(',','.'))   
        self.turmGeoTop['Ra'] = self.turmGeoTop['Rm']+self.flangeGeoTop['flange_w']*0.5   

        self.flangeGeoTop['cutOut_w']  = float(self.TB_upperCutOut_w.Text.replace(',','.')) 
        self.flangeGeoTop['cutOut_h']  = float(self.TB_upperCutOut_h.Text.replace(',','.')) 
        self.flangeGeoTop['cutOut_topE1']  = float(self.TB_upperCutOut_E1.Text.replace(',','.')) 
        self.flangeGeoTop['cutOut_topE2']  = float(self.TB_upperCutOut_E2.Text.replace(',','.'))
        self.flangeGeoTop['cutOut_midR']  = float(self.TB_upperCutOut_midR.Text.replace(',','.'))
        self.flangeGeoTop['cutOut_botR']  = float(self.TB_upperCutOut_botR.Text.replace(',','.'))     
        
        
        ### bottom Flange
        self.flangeGeoBot={}
        self.turmGeoBot={}
        
        self.flangeGeoBot['side']='bot'
        self.turmGeoBot['side']='bot' 
        
        self.flangeGeoBot['neck_s'] = float(self.TB_lowerFlange_neck_s.Text.replace(',','.'))
        self.flangeGeoBot['neck_h']  = float(self.TB_lowerFlange_neck_h.Text.replace(',','.'))
        
        self.flangeGeoBot['flange_h']  = float(self.TB_lowerFlange_h.Text.replace(',','.'))        
        self.flangeGeoBot['flange_w']  = float(self.TB_lowerFlange_w.Text.replace(',','.'))            
        self.flangeGeoBot['flange_topL']  = float(self.TB_lowerFlange_topL.Text.replace(',','.'))            
        
        
#        self.turmGeoBot['s'] = float(self.TB_turmBot_s.Text.replace(',','.')) 
#        self.turmGeoBot['l'] = float(self.TB_turmBot_l.Text.replace(',','.')) 
        self.turmGeoBot['Rm'] = float(self.TB_segment_Rm.Text.replace(',','.'))   
        self.turmGeoBot['Ra'] = self.turmGeoBot['Rm']+self.flangeGeoBot['flange_w']*0.5
        
        self.flangeGeoBot['cutOut_w']  = float(self.TB_lowerCutOut_w.Text.replace(',','.')) 
        self.flangeGeoBot['cutOut_h']  = float(self.TB_lowerCutOut_h.Text.replace(',','.')) 
        self.flangeGeoBot['cutOut_E1']  = float(self.TB_lowerCutOut_E1.Text.replace(',','.')) 
        self.flangeGeoBot['cutOut_midR']  = float(self.TB_lowerCutOut_midR.Text.replace(',','.'))
        self.flangeGeoBot['cutOut_topR']  = float(self.TB_lowerCutOut_topR.Text.replace(',','.'))   
        
        ###Segment
        self.segment={}
        self.segment['angle'] = float(self.TB_segment_angle.Text.replace(',','.')) 
        ### Fastener
        self.fastener={}
        self.fastener['n']= int(self.TB_number_fastener.Text.replace(',','.')) 
        #
        self.dicGeo={'flangeGeoBot':self.flangeGeoBot , 'turmGeoBot':self.turmGeoBot, 'flangeGeoTop':self.flangeGeoTop , 'turmGeoTop':self.turmGeoTop, 'fastener': self.fastener , 'segment': self.segment}
        #form=openWindowC1Geo()
        self.C1_Connection=C1_Geo(self.dicGeo)
        self.C1_Connection.createPart()        
        
        #self.Close() 
            
        

        
#        self.inputDic['Bolts_DC36']=self.check.Checked
#        self.inputDic['Bolts_DC50']=self.check2.Checked
#        self.inputDic['Weld_inside']=self.check3.Checked
#        self.inputDic['Weld_outside']=self.check4.Checked
#        self.inputDic['Bolt_holes']=self.check6.Checked
#        self.inputDic['Flange_rounding']=self.check5.Checked
#        self.inputDic['getGaps']=self.check7.Checked

        #ExtAPI.Log.WriteMessage('######## Start calcFLS ########')
        #self.Close()     

### Skript SCTools

def oninit(context):
    return

class showTopFlange(Form):


    def __init__(self):
            #
   
 
        ctypes.windll.user32.MessageBoxW(0,sys.version, "User Current Version:-", 1)
        
        posi = __file__.rfind(chr(92))
        #file_name = raw_input("File Name: ")
        self.filename = __file__[0:posi]+"\\images\\upperFlangePara2.jpg"
        ctypes.windll.user32.MessageBoxW(0,self.filename, "Development Msg", 1)
        self.image = Image.FromFile(self.filename)
        
        
        #self.CenterToScreen()
        #self.geoDic={}
        
        #self.image.siy
        pictureBox = PictureBox()
        
        
        pictureBox.Image = self.image
        #pictureBox.SizeMode = PictureBoxSizeMode.CenterImage;
        #pictureBox.Size=Size(4160,2102)
        
        #pictureBox.Size.Height=500
        #pictureBox.Size.Width=1000
        
        pictureBox.Dock = DockStyle.Fill
        
        #self.Size.Height=500
        #self.Size.Width=1000
        self.Controls.Add(pictureBox)
        #self.AutoSize=True
        #self.SizeMode = AutoSize
        
        
        #form.Size.Height=500
        #form.Size.Width=1000
        
        #        self.pb=PictureBox()
        #        self.pb.Size = Size(188, 200)
        #        self.pb.Location = Point(300, 0)
        #        self.pb.Image = Image.FromFile(r"C:\Users\filip.wegener\OneDrive - JBO GmbH\dropbox\Dropbox\Data\Ansys\Extensions\C1Tool\C1Tool\images\upperFlangePara_klein.png") 
        #        self.Controls.Add(self.pb)
        
        #timeSteps=(0,1,2,3)
        #form = getInputGeo()
        #form.AutoSize=True
        #form.Size.Height=500
        #form.Size.Width=1000
        #form = IForm()
        #form.ShowDialog()
        
        self.Show()
        
        
        
class C1_Geo():
    #
    def __init__(self,dicGeo): 
        
        self.dicGeo=dicGeo

        GetRootPart().ClearAllPartData()
        
        # Collect all property values in mm
        #Get tower and top flange parameter 

        self.flangeGeoTop=self.dicGeo['flangeGeoTop']
        self.turmGeoTop=self.dicGeo['turmGeoTop']
    
        #Get tower and bot flange parameter

        self.flangeGeoBot=self.dicGeo['flangeGeoBot']
        self.turmGeoBot=self.dicGeo['turmGeoBot']

    
        
        self.flangeGeoBot['cutOut_w']  = 115        
        self.flangeGeoBot['cutOut_h']  = 175        
        self.flangeGeoBot['cutOut_topR']  = 50  
        self.flangeGeoBot['cutOut_midR']  = 290  
        self.flangeGeoBot['cutOut_E1']  = 30  
        
        # fastener
        self.fastener=self.dicGeo['fastener']
        
        #segment
        self.segment=self.dicGeo['segment']
       


    def createPart(self):

        
        #dic=getProps_DM()
        #ExtAPI.Log.WriteMessage('modelPath2:'+self.modelPath)
        #flangeInfos=addDic2FlangeInfo({'geoSettings':dic})   
        #exportInfos2Excel(flangeInfos)       
        
        #ExtAPI.Log.WriteMessage('self.feature.Bodies:'+str(self.feature.Bodies))
        
        #ExtAPI.Log.WriteMessage('dir(self.feature.Bodies):'+str(dir(self.feature.Bodies)))
        #ExtAPI.Log.WriteMessage('dir(self.feature):'+str(dir(self.feature)))
        #create Flansches
        self.createtopFlansch(flangeGeo=self.flangeGeoTop,turmGeo=self.turmGeoTop,segment=self.segment,fastener=self.fastener)
        self.createbotFlansch(flangeGeo=self.flangeGeoBot,turmGeo=self.turmGeoBot,segment=self.segment,fastener=self.fastener)
        #create cutOuts
        self.createCutOutTopFlange(flangeGeo=self.flangeGeoTop, fastener=self.fastener,segment=self.segment)
        self.createCutOutBotFlange(flangeGeo=self.flangeGeoBot, fastener=self.fastener,segment=self.segment)
        #pattern
        self.patternTopFlansch(fastener=self.fastener,segment=self.segment)
        self.patternBotFlansch(fastener=self.fastener,segment=self.segment)
        #create Connector
        self.connector=self.createConnector(fastener=self.fastener)
        #assembly connector
        #self.patternConnector()
        #assembly helpBodies       
        #self.createHoles(flangeGeo=self.flangeGeoTop,turmGeo=self.turmGeoTop,side='top')
        #self.createHoles(flangeGeo=self.flangeGeoBot,turmGeo=self.turmGeoBot,side='bot')
        #create Bodies
        #self.combineBodies()
        
        #self.feature.Bodies = self.bodies
        #self.feature.MaterialType = MaterialTypeEnum.Freeze 
        
        #buildTime=time.clock()-start
        #ExtAPI.Log.WriteMessage('-- Finished Flange build up -- time:'+str(buildTime)+'s')
        
        #analytical
        #fitView()
        #
        
        #ExtAPI.Application.ScriptByName('jscript').ExecuteCommand('AG.ag.gb.FitView();')
        #ExtAPI.Graphics.Camera.InternalObject.zoomFit()
        #ExtAPI.Graphics.Redraw()
        #self.createInfoFile()
        #


    def combineBodies(self):

        flangeBot = self.builder.Operations.Tools.CreatePart(self.bodies2CombineBot)
        flangeBot.Name = "flangeBot"
        self.bodies.append(flangeBot)
        
        
        flangeTop = self.builder.Operations.Tools.CreatePart(self.bodies2CombineTop)
        flangeTop.Name = "flangeTop"
        self.bodies.append(flangeTop)
        

#        helpBody = self.builder.Operations.Tools.CreatePart(self.helpBodies)
#        helpBody.Name = "helpBody"                
#        self.bodies.append(helpBody)
                               
                            
    def patternConnector(self):
        
        pass


    def createCutOutTopFlange(self,flangeGeo,fastener,segment):

        # Set Sketch Plane
        sectionPlane = Plane.PlaneYZ
        result = ViewHelper.SetSketchPlane(sectionPlane, None)
        sketch = SketchHelper.StartConstraintSketching()

        #
        # Sketch Line
        p1=Point2D.Create( MM(0.0),MM(0.0))
        p2=Point2D.Create( MM(flangeGeo['cutOut_w']*0.5-flangeGeo['cutOut_topE1']),MM(0.0))
        p3=Point2D.Create( MM(flangeGeo['cutOut_w']*0.5),MM(-flangeGeo['cutOut_topE2']))
        
        alpha_rad=math.acos((flangeGeo['cutOut_w']*0.5-flangeGeo['cutOut_midR'])/(flangeGeo['cutOut_botR']-flangeGeo['cutOut_midR']))
        y_midR=math.sin(alpha_rad)*flangeGeo['cutOut_midR']
        y_botR=flangeGeo['cutOut_botR']-math.sin(alpha_rad)*flangeGeo['cutOut_botR']
        
        p4=Point2D.Create( MM(flangeGeo['cutOut_w']*0.5),MM(-flangeGeo['cutOut_h']+y_midR+y_botR))        

        p5y=-flangeGeo['cutOut_h']+y_botR
        p5x=math.cos(alpha_rad)*flangeGeo['cutOut_botR']
        
        p5=Point2D.Create( MM(p5x),MM(p5y))   
        p6=Point2D.Create( MM(0.0),MM(-flangeGeo['cutOut_h']))  
        
        # create lines and arcs
        l1=SketchLine.Create(p1, p2).CreatedCurves[0]
 
        # create lines and arcs
        origin = Point2D.Create(MM(flangeGeo['cutOut_w']*0.5-flangeGeo['cutOut_topE1']), MM(-flangeGeo['cutOut_topE2']))
        majorDir = DirectionUV.DirU
        minorDir = -DirectionUV.DirV
        ellipse = SketchEllipse.Create(origin, majorDir, minorDir, MM(flangeGeo['cutOut_topE1']), MM(flangeGeo['cutOut_topE2'])) 
       
        #curvePoint=SelectionPoint.Create(l1.GetChildren[ICurvePoint]()[1])
        #arc1=SketchArc.CreateTangentArc(curvePoint, p3).CreatedCurves[0]        
        
        l2=SketchLine.Create(p3, p4).CreatedCurves[0]

        curvePoint=SelectionPoint.Create(l2.GetChildren[ICurvePoint]()[1])
        arc2=SketchArc.CreateTangentArc(curvePoint, p5).CreatedCurves[0] 

        curvePoint=SelectionPoint.Create(arc2.GetChildren[ICurvePoint]()[1])
        arc3=SketchArc.CreateTangentArc(curvePoint, p6).CreatedCurves[0]   

        #trim ellipse
        curveSelPoint = SelectionPoint.Create(GetRootPart().DatumPlanes[-1].Curves[1], 4.70425724408983)
        result = TrimSketchCurve.Execute(curveSelPoint)

        l3=SketchLine.Create(p6, p1).CreatedCurves[0]
        
        #mirror
        selection = Selection.Create(GetRootPart().DatumPlanes[-1].Curves[:5])
        mirrorPlane = Selection.Create(l3)
        options = MoveOptions()
        options.Copy = True
        options.SnapAssociatedVertices = False
        result = Move.MirrorEntities(selection, mirrorPlane, options)
        
        selection = Selection.Create(l3)
        result = Delete.Execute(selection)

        # create face
        mode = InteractionMode.Solid
        face2Cut = ViewHelper.SetViewMode(mode).GetCreated[IDesignFace]()[0]
        
#        # Create Pattern
        selection = BodySelection.Create(GetRootPart().Bodies[-1])
        data = CircularPatternData()
        data.CircularAxis = Selection.Create(GetRootPart().CoordinateSystems[0].Axes[2])
        data.RadialDirection = Direction.Create(0, 0, 0)
        data.CircularCount = 2
        data.CircularAngle = DEG(segment['angle']/(fastener['n']-1))
        pattern = Pattern.CreateCircular(selection, data, None)

 
        # Change Object Visibility
        selection = BodySelection.Create(self.botFlange)
        visibility = VisibilityType.Hide
        inSelectedView = False
        faceLevel = False
        ViewHelper.SetObjectVisibility(selection, visibility, inSelectedView, faceLevel)

        #cut flange
        selection = FaceSelection.Create(GetRootPart().Components[-1].Components[0].Content.Bodies[0].Faces[0])
        options = ExtrudeFaceOptions()
        options.ExtrudeType = ExtrudeType.ForceCut
        result = ExtrudeFaces.Execute(selection, MM(1e6), options)
        
        # Change Object Visibility
        selection = BodySelection.Create(self.botFlange)
        visibility = VisibilityType.Show
        inSelectedView = False
        faceLevel = False
        ViewHelper.SetObjectVisibility(selection, visibility, inSelectedView, faceLevel)
        # EndBlock        

    def createCutOutBotFlange(self,flangeGeo,fastener,segment):

        # Set Sketch Plane
        sectionPlane = Plane.PlaneYZ
        result = ViewHelper.SetSketchPlane(sectionPlane, None)
        sketch = SketchHelper.StartConstraintSketching()

        # 
        offset_y=-flangeGeo['flange_topL']
        # Sketch Line
        p1=Point2D.Create( MM(0.0),MM(offset_y))

        alpha_rad=math.acos((flangeGeo['cutOut_w']*0.5-flangeGeo['cutOut_midR'])/(flangeGeo['cutOut_topR']-flangeGeo['cutOut_midR']))
        y_midR=math.sin(alpha_rad)*flangeGeo['cutOut_midR']
        y_topR=flangeGeo['cutOut_topR']-math.sin(alpha_rad)*flangeGeo['cutOut_topR'] 

        p2y=offset_y-y_topR
        p2x=math.cos(alpha_rad)*flangeGeo['cutOut_topR']
        p2=Point2D.Create( MM(p2x),MM(p2y))

        p3=Point2D.Create( MM(flangeGeo['cutOut_w']*0.5),MM(offset_y-y_topR-y_midR))
        
        p4=Point2D.Create( MM(flangeGeo['cutOut_w']*0.5),MM(offset_y-flangeGeo['cutOut_h']+flangeGeo['cutOut_E1']))        
        p5=Point2D.Create( MM(0.0),MM(offset_y-flangeGeo['cutOut_h']))       

        
        # create lines and arcs
        origin = Point2D.Create(MM(0.0), MM(offset_y-flangeGeo['cutOut_h']+flangeGeo['cutOut_E1']))
        majorDir = DirectionUV.DirU
        minorDir = -DirectionUV.DirV
        ellipse = SketchEllipse.Create(origin, majorDir, minorDir, MM(flangeGeo['cutOut_w']*0.5), MM(flangeGeo['cutOut_E1']))        

        l1=SketchLine.Create(p4, p3).CreatedCurves[0]
        
        curvePoint=SelectionPoint.Create(l1.GetChildren[ICurvePoint]()[1])
        arc1=SketchArc.CreateTangentArc(curvePoint, p2).CreatedCurves[0]        

        curvePoint=SelectionPoint.Create(arc1.GetChildren[ICurvePoint]()[1])
        arc2=SketchArc.CreateTangentArc(curvePoint, p1).CreatedCurves[0]

        l2=SketchLine.Create(p1, p5).CreatedCurves[0]
        
        #mirror
        selection = Selection.Create([l1,arc1,arc2])
        mirrorPlane = Selection.Create(l2)
        options = MoveOptions()
        options.Copy = True
        options.SnapAssociatedVertices = False
        result = Move.MirrorEntities(selection, mirrorPlane, options)
        
        selection = Selection.Create(l2)
        result = Delete.Execute(selection)

        #trim ellipse
        curveSelPoint = SelectionPoint.Create(GetRootPart().DatumPlanes[-1].Curves[0], 4.70425724408983)
        result = TrimSketchCurve.Execute(curveSelPoint)
        # create face
        mode = InteractionMode.Solid
        face2Cut = ViewHelper.SetViewMode(mode).GetCreated[IDesignFace]()[0]
        
        # Create Pattern
        selection = BodySelection.Create(GetRootPart().Bodies[-1])
        data = CircularPatternData()
        data.CircularAxis = Selection.Create(GetRootPart().CoordinateSystems[0].Axes[2])
        data.RadialDirection = Direction.Create(0, 0, 0)
        data.CircularCount = 2
        data.CircularAngle = DEG(segment['angle']/(fastener['n']-1))
        pattern = Pattern.CreateCircular(selection, data, None)
     
 
        # Change Object Visibility
        selection = BodySelection.Create(self.topFlange)
        visibility = VisibilityType.Hide
        inSelectedView = False
        faceLevel = False
        ViewHelper.SetObjectVisibility(selection, visibility, inSelectedView, faceLevel)

 
        selection = FaceSelection.Create(GetRootPart().Components[-1].Components[0].Content.Bodies[0].Faces[0]) 
        options = ExtrudeFaceOptions()
        options.ExtrudeType = ExtrudeType.ForceCut
        result = ExtrudeFaces.Execute(selection, MM(1e6), options)
        
        # Change Object Visibility
        selection = BodySelection.Create(self.topFlange)
        visibility = VisibilityType.Show
        inSelectedView = False
        faceLevel = False
        ViewHelper.SetObjectVisibility(selection, visibility, inSelectedView, faceLevel)
        # EndBlock       
            
    def createtopFlansch(self,flangeGeo,turmGeo,segment,fastener):
        ###########################################################################   
        # create Flansch Profile
      
        # Set Sketch Plane
        sectionPlane = Plane.PlaneZX
        result = ViewHelper.SetSketchPlane(sectionPlane, None)
        sketch = SketchHelper.StartConstraintSketching()
        
        # Sketch Line
        p1=Point2D.Create( MM(flangeGeo['neck_h']+flangeGeo['flange_h']),MM(turmGeo['Rm']), )
        p2=Point2D.Create( MM(flangeGeo['neck_h']+flangeGeo['flange_h']),MM(turmGeo['Rm']-flangeGeo['neck_s']*0.5))
        p3=Point2D.Create( MM(flangeGeo['flange_h']+flangeGeo['neck_r']),MM(turmGeo['Rm']-flangeGeo['neck_s']*0.5) )
        p4=Point2D.Create( MM(flangeGeo['flange_h']), MM(turmGeo['Rm']-flangeGeo['neck_s']*0.5-flangeGeo['neck_r']))
        p5=Point2D.Create( MM(flangeGeo['flange_h']), MM(turmGeo['Rm']-flangeGeo['flange_w']*0.5))
        p6=Point2D.Create( MM(-flangeGeo['fork_h']), MM(turmGeo['Rm']-flangeGeo['flange_w']*0.5))
        p7=Point2D.Create( MM(-flangeGeo['fork_h']), MM(turmGeo['Rm']-flangeGeo['flange_w']*0.5+flangeGeo['fork_w']))
        p8=Point2D.Create( MM(0.0),MM(turmGeo['Rm']-flangeGeo['flange_w']*0.5+flangeGeo['fork_w']))
        p9=Point2D.Create( MM(0.0),MM(turmGeo['Rm']))
        
        l1=SketchLine.Create(p1, p2).CreatedCurves[0]
        l2=SketchLine.Create(p2, p3).CreatedCurves[0]
        
        curvePoint=SelectionPoint.Create(l2.GetChildren[ICurvePoint]()[1])
        arc1=SketchArc.CreateTangentArc(curvePoint, p4).CreatedCurves[0]
                
        l3=SketchLine.Create(p4, p5).CreatedCurves[0]
        l4=SketchLine.Create(p5, p6).CreatedCurves[0]
        l5=SketchLine.Create(p6, p7).CreatedCurves[0]
        l6=SketchLine.Create(p7, p8).CreatedCurves[0]
        l7=SketchLine.Create(p8, p9).CreatedCurves[0]
        l8=SketchLine.Create(p9, p1).CreatedCurves[0]
        
        #mirror
        selection = Selection.Create([l1,l2,l3,l4,l5,l6,l7,arc1])
        mirrorPlane = Selection.Create(l8)
        options = MoveOptions()
        options.Copy = True
        options.SnapAssociatedVertices = False
        result = Move.MirrorEntities(selection, mirrorPlane, options)
        
        selection = Selection.Create(l8)
        result = Delete.Execute(selection)
        
        
        mode = InteractionMode.Solid
        face2Revovle = ViewHelper.SetViewMode(mode).GetCreated[IDesignFace]()[0]
                
        # revolve
                
        selection = FaceSelection.Create(face2Revovle)
        axisSelection = Selection.Create(GetRootPart().CoordinateSystems[0].Axes[2])
        axis = RevolveFaces.GetAxisFromSelection(selection, axisSelection)
        options = RevolveFaceOptions()
        options.ExtrudeType = ExtrudeType.Add
        
#        self.topFlange = RevolveFaces.Execute(selection, axis,DEG(segment['angle']), options).CreatedBodies[0]
#        self.topFlange.SetName('topFlange')

        self.topFlange = RevolveFaces.Execute(selection, axis,DEG(segment['angle']/(fastener['n']-1)), options).CreatedBodies[0]
        self.topFlange.SetName('topFlange')

    def createbotFlansch(self,flangeGeo,turmGeo,segment,fastener):
        ###########################################################################   
        # create Flansch Profile
              
        # Set Sketch Plane
        sectionPlane = Plane.PlaneZX
        result = ViewHelper.SetSketchPlane(sectionPlane, None)
        sketch = SketchHelper.StartConstraintSketching()
        
        # Sketch Line
        p1=Point2D.Create( MM(0.0),MM(turmGeo['Rm']-flangeGeo['flange_w']*0.5) )
        p2=Point2D.Create( MM(-flangeGeo['flange_h']),MM(turmGeo['Rm']-flangeGeo['flange_w']*0.5) )
        p3=Point2D.Create( MM(-flangeGeo['flange_h']-flangeGeo['neck_h']),MM(turmGeo['Rm']-flangeGeo['neck_s']*0.5) )
        p4=Point2D.Create( MM(-flangeGeo['flange_h']-flangeGeo['neck_h']),MM(turmGeo['Rm']+flangeGeo['neck_s']*0.5) )
        p5=Point2D.Create( MM(-flangeGeo['flange_h']),MM(turmGeo['Rm']+flangeGeo['flange_w']*0.5) )
        p6=Point2D.Create( MM(0.0),MM(turmGeo['Rm']+flangeGeo['flange_w']*0.5))
        
        l1=SketchLine.Create(p1, p2).CreatedCurves[0]
        l2=SketchLine.Create(p2, p3).CreatedCurves[0]
        l3=SketchLine.Create(p3, p4).CreatedCurves[0]
        l4=SketchLine.Create(p4, p5).CreatedCurves[0]
        l5=SketchLine.Create(p5, p6).CreatedCurves[0]
        l6=SketchLine.Create(p6, p1).CreatedCurves[0]
        
        #mirror
        
        mode = InteractionMode.Solid
        face2Revolve = ViewHelper.SetViewMode(mode).GetCreated[IDesignFace]()[0]
                
        # revolve
                
        selection = FaceSelection.Create(face2Revolve)
        axisSelection = Selection.Create(GetRootPart().CoordinateSystems[0].Axes[2])
        axis = RevolveFaces.GetAxisFromSelection(selection, axisSelection)
        options = RevolveFaceOptions()
        options.ExtrudeType = ExtrudeType.ForceIndependent
        
#        self.botFlange = RevolveFaces.Execute(selection, axis,DEG(segment['angle']), options).CreatedBodies[0]
#        self.botFlange.SetName('botFlange')      
        
        self.botFlange = RevolveFaces.Execute(selection, axis,DEG(segment['angle']/(fastener['n']-1)), options).CreatedBodies[0]
        self.botFlange.SetName('topFlange')        

    def patternTopFlansch(self,segment,fastener):
        ###########################################################################   
        # create Flansch Profile
        selection = BodySelection.Create(self.topFlange)
        data = CircularPatternData()
        data.CircularAxis = Selection.Create(GetRootPart().CoordinateSystems[0].Axes[2])
        data.RadialDirection = Direction.Create(0, 0, 0)
        data.CircularCount = fastener['n']-1
        data.CircularAngle = DEG(segment['angle']-(segment['angle']/(fastener['n']-1)))
        result = Pattern.CreateCircular(selection, data, None)

    def patternBotFlansch(self,segment,fastener):
        ###########################################################################   
        # create Flansch Profile
        selection = BodySelection.Create(self.botFlange)
        data = CircularPatternData()
        data.CircularAxis = Selection.Create(GetRootPart().CoordinateSystems[0].Axes[2])
        data.RadialDirection = Direction.Create(0, 0, 0)
        data.CircularCount = fastener['n']-1
        data.CircularAngle = DEG(segment['angle']-(segment['angle']/(fastener['n']-1)))
        result = Pattern.CreateCircular(selection, data, None)
        
    def createConnector(self,fastener):
        #Niklas hier kannst du Dich austoben
        
        return connector

def exportInfos2Excel(flangeInfos,dicFLS={'InputFileMarkov':['']}):

    for ext in ExtAPI.ExtensionManager.Extensions:
        if ext.Name == 'FlangeTool':
            dirExt=ext.InstallDir
            break
    
    modelPath=getModelPath()

    for MMPath in dicFLS['InputFileMarkov']:
        
        ExtAPI.Log.WriteMessage('MMPath:'+MMPath)

        anaName=ExtAPI.Application.InternalObject.Title.split(': ')[1].split(' -')[0]

        if '.mkv' in  MMPath:
            mmDic=loadMM_SGRE(pathFile=MMPath)
            flangeInfos=addDic2FlangeInfo(mmDic) 
            MMName=MMPath.split('\\')[-1].split('.')[0]
            savePath=os.path.join(modelPath,'FlangeInfo_'+anaName+'_'+MMName+'.xlsx')

        
        elif '.xls' in  MMPath:
            mmDic=loadMM_Excel(pathFile=MMPath)
            flangeInfos=addDic2FlangeInfo(mmDic) 
            MMName=MMPath.split('\\')[-1].split('.')[0]
            savePath=os.path.join(modelPath,'FlangeInfo_'+anaName+'_'+MMName+'.xlsx')
            
        else:
            savePath=os.path.join(modelPath,'FlangeInfo_'+anaName+'.xlsx')
      
        excelFile = xlsxwriter.Workbook(savePath)
                
        keyList=['geoSettings','mechSettings','geoPropsMech','impSettings','Bolt_Fs','Bolt_Mx','Bolt_My','topFlange_innerWeld','botFlange_innerWeld','topFlange_outerWeld','botFlange_outerWeld','topFlange_Rounding','botFlange_Rounding','topFlange_Holes','botFlange_Holes','Gaps Inner Radius','Gaps Outer Radius','Gaps Outer Radius2','markovArray','calcFLS']
            
        for key1 in keyList:
            if key1 in flangeInfos.keys(): 
                
                sheet=excelFile.add_worksheet(key1)
            
                row=0
                col=0
                for key2 in flangeInfos[key1]['keyList']:
                    #ExtAPI.Log.WriteMessage('key2:'+key2)
            
                    if isinstance(flangeInfos[key1][key2],list):
    #                    ExtAPI.Log.WriteMessage('hallo_1')                
                        #ExtAPI.Log.WriteMessage('key1:'+key1)
                        #ExtAPI.Log.WriteMessage('key2:'+key2)
                        
                        sheet.write(0,col,key2)
                        row=1
                        for value in flangeInfos[key1][key2]:
                            sheet.write(row,col,value)
                            row+=1
                        col+=1
                        
                    else:
    #                    ExtAPI.Log.WriteMessage('hallo_2')                
    #                    ExtAPI.Log.WriteMessage('key1:'+key1)
    #                    ExtAPI.Log.WriteMessage('key2:'+key2)
                        
                        sheet.write(row,0,key2)
                        sheet.write(row,1,flangeInfos[key1][key2])
                        row+=1 
        if dicFLS['InputFileMarkov'][0] != '':
            excelFile=createFLSExcel(excelFile,flangeInfos,dicFLS)
            
        excelFile.close() 
    
#    if dicFLS:
#        path=os.path.realpath(savePath)
#        os.startfile(savePath)     
    
def addDic2FlangeInfo(dic,noUpdate=False):

    for ext in ExtAPI.ExtensionManager.Extensions:
        if ext.Name == 'FlangeTool':
            dirExt=ext.InstallDir
            break
        
    modelPath=getModelPath()
    
    ExtAPI.Log.WriteMessage('modelPath3:'+modelPath)
    pathFile=os.path.join(modelPath,'flangeInfos')
    #load flangeInfo file
    
    try:
        file=open(pathFile, "rb" )
        flangeInfos=pickle.load(file)
        file.close()
    except:
        flangeInfos={} 
        
    if noUpdate:
        flangeInfos=dic
    else:
        flangeInfos.update(dic)
    
    # save infofile
    file=open(pathFile, "wb" )
    pickle.dump(flangeInfos,file)
    file.close()
    
    return flangeInfos
    # write excel        
        
#def createInterpolationFunctions4Excel():

def getModelPath():
    
    try:        
        title = ExtAPI.Application.InternalObject.Title
        sys_coord = title.split(":")[0].replace(' ','')
        cc='''
try:
    sys_coord = "{0}"
    comps = ACT.GetAllComponents()
    for comp in comps:
        ts = ACT.GetTaskForContainer(comp.GetContainer())
        if ts.CoordinateId.startswith(sys_coord):# and ts.Name == "Geometry":
            returnValue(ts.ActiveDirectory)
except:
    returnValue(None)            
        '''.format(sys_coord)
        z=wbjn.ExecuteCommand(ExtAPI,cc)
#        z=z.Replace('ACT','MECH')
#        
#        if not os.path.exists(z):
#            os.makedirs(z)        
        
    except:
        z=None
 
  
    if z == None: 
        for extension in ExtAPI.ExtensionManager.Extensions:
            if extension.Name == 'FlangeTool':
                FT_ext=extension
                break
                
        FT_obj=ExtAPI.DataModel.GetUserObjects(FT_ext)[0]
        props=FT_obj.AllProperties
        z=props[len(props)-1].InternalValue        
    
    return z


form=openWindowC1Geo()