# -*- coding: utf-8 -*-
"""
Created on Thu Oct 22 16:46:07 2017

Copyright 2017-2021 Javier Alejandro Cuartas Micieces

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in 
the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of 
the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS 
FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER 
IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

"""
import os
import time, datetime
import pandas as pd
import copy
import sys
import importlib
import json
import xlrd
import ctypes

from selenium import webdriver

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC 

import pyautogui as pa

class ToolBoxMethods:
    @staticmethod
    def capslockv(*args):
        dllv = ctypes.WinDLL ("User32.dll")
        Mayusc = 0x14
        return dllv.GetKeyState(Mayusc)

    @staticmethod
    def log_in_web(*args):#Arguments (filepath,portal,input_tag_name_attribute_value_user,input_tag_name_attribute_value_pass,input_tag_name_attribute_value_submit,username_string,password_string,hidden_window_mode_0)
        ch=webdriver.ChromeOptions()
        ch.add_argument("--start-maximized");
        if args[7]==0:
            ch.add_argument('--headless')#If you set last argument to 0, browser will be run hidden mode.
        browser=webdriver.Chrome(executable_path=args[0], options=ch)
        browser.get(args[1])
        user=browser.find_element_by_name(args[2])
        password=browser.find_element_by_name(args[3])
        subm=browser.find_element_by_name(args[4])
        user.send_keys(args[5])
        password.send_keys(args[6])
        subm.click()
        return browser
    
    @staticmethod
    def autotree(*args):#Arguments or tuple of lists with the arguments, like [browser,timeoutseconds,"xpath",number_of_element_in_xpath]
        if type(args[0]) is list:
            if (len(args)>1):
                browser=ToolBoxMethods.autotree(*tuple(args[0]))
                args=args[1:]
                browser=ToolBoxMethods.autotree(*args)
            else:
                browser=ToolBoxMethods.autotree(*tuple(args[0]))
        else:
            browser=args[0]
            WebDriverWait(browser,args[1]).until(EC.element_to_be_clickable((By.XPATH, args[2])))
            p1=browser.find_elements_by_xpath(args[2])
            p1[args[3]].click()
        return browser
    
    @staticmethod
    def img_autotree(*args):#Arguments or tuple of lists with the arguments, like [filepath,filename,waiting_time_between_clicks,foldernameending]
        if type(args[0]) is list:
            if (len(args)>1):
                ToolBoxMethods.img_autotree(*tuple(args[0]))
                args=args[1:]
                ToolBoxMethods.img_autotree(*args)
            else:
                ToolBoxMethods.img_autotree(*tuple(args[0]))
        else:
            pa.moveTo(50,50)
            focus1=ToolBoxMethods.try_imgs(args[0],args[1],1,args[3])
            pa.click(focus1[0],focus1[1])
            time.sleep(int(args[2]))
    
    @staticmethod
    def try_imgs(*args):#Arguments: (filepath, file, grayscale 1 for True,foldernameending)
        x0,y0=pa.size()
        pa.moveTo(x=x0-80,y=y0-4)
        h=[el for el in os.listdir(args[0]) if el.endswith(args[3])]
        i=len(h)-1
        for el in h:
            g=pa.locateCenterOnScreen(args[0]+h[i]+"\\\\"+args[1],grayscale=True) if args[2]==1 else pa.locateCenterOnScreen(args[0]+h[i]+"\\\\"+args[1])
            if g==None:
                g=None
                i=i-1
            else:
                break
        return g
        
    @staticmethod
    def get_static_html_as_string(*args):#Arguments: (filepath,foldernameending)
        focus0=ToolBoxMethods.try_imgs(args[0],"Focus15.png",0,args[1])
        print(focus0)
        if focus0!=None: 
            # time.sleep(1)
            pa.hotkey('ctrl','u')
            time.sleep(2)
            pa.keyDown("ctrl")
            pa.keyDown("a")
            pa.keyUp("ctrl")
            pa.keyUp("a")
            time.sleep(2)
            pa.keyDown("ctrl")
            pa.keyDown("c")
            pa.keyUp("ctrl")
            pa.keyUp("c")
            
            imglist=["Focus13.png","Focus5.png","Focus10.png"]
            
            ToolBoxMethods.checkemptybackground(args[0],imglist,args[1],"ctrl")
            
            A_l=args[0].split("\\")
            A_l=[value for value in A_l if value != '']
            gs='/'.join(A_l)
            os.startfile("notepad.exe")
            time.sleep(1)
            pa.keyDown("ctrl")
            pa.keyDown("v")
            pa.keyUp("ctrl")
            pa.keyUp("v")
            
            time.sleep(2)
            pa.keyDown("ctrl")
            pa.keyDown("g")
            pa.keyUp("ctrl")
            pa.keyUp("g")
            
            pa.write("auxFrete")
            pa.press(["tab","tab","tab","tab","tab","tab","tab","enter"])
            pa.write(gs)
            pa.press("enter")
            time.sleep(1)
            pa.press(["tab","tab","tab","tab","tab","tab","tab","tab","tab"])
            pa.press("enter")
            pa.press("enter")
            
            ToolBoxMethods.checkemptybackground(args[0],imglist,args[1],"alt")
            pa.press("enter")
            try:
                with open(args[0]+"auxFrete.txt","r", errors='ignore') as r:
                    data=r.read()
                os.remove("auxFrete.txt")
            except:
                data=""
                pass
            edft=0
            pa.moveTo(50,50)
        else:
            edft=1
        return edft,data

    
    @staticmethod
    def get_dynamic_html_as_string(*args):#Argument: (filepath,htmlfilename,foldernameending)
        focus0=ToolBoxMethods.try_imgs(args[0],"Focus15.png",0,args[2])
        if focus0!=None: 
            pa.hotkey('ctrl','shift','I')
            time.sleep(2)
            focus1=ToolBoxMethods.try_imgs(args[0],args[1],0,args[2])
            if focus1==None:
                time.sleep(3)
                focus1=ToolBoxMethods.try_imgs(args[0],args[1],0,args[2])
            pa.click(button='Right',x=focus1[0],y=focus1[1])
            time.sleep(1)
            pa.press(['down','down','enter'])
            time.sleep(1)
            pa.keyDown("ctrl")
            pa.keyDown("a")
            pa.keyUp("ctrl")
            pa.keyUp("a")
            time.sleep(2)
            pa.keyDown("ctrl")
            pa.keyDown("c")
            pa.keyUp("ctrl")
            pa.keyUp("c")
            pa.hotkey('ctrl','shift','I')
            time.sleep(1)
            A_l=args[0].split("\\")
            A_l=[value for value in A_l if value != '']
            gs='/'.join(A_l)
            print(gs)
            os.startfile("notepad.exe")
            time.sleep(1)
            pa.keyDown("ctrl")
            pa.keyDown("v")
            pa.keyUp("ctrl")
            pa.keyUp("v")
            
            imglist=["Focus13.png","Focus5.png","Focus10.png"]
            
            time.sleep(2)
            pa.keyDown("ctrl")
            pa.keyDown("g")
            pa.keyUp("ctrl")
            pa.keyUp("g")
            pa.write("auxFrete")
            pa.press(["tab","tab","tab","tab","tab","tab","tab","enter"])
            pa.write(gs)
            pa.press("enter")
            time.sleep(1)
            pa.press(["tab","tab","tab","tab","tab","tab","tab","tab","tab"])
            pa.press("enter")
            pa.press("enter")
            
            ToolBoxMethods.checkemptybackground(args[0],imglist,args[2],"alt")
            pa.press("enter")
            try:
                with open(args[0]+"auxFrete.txt","r", errors='ignore') as r:
                    data=r.read()
                os.remove("auxFrete.txt")
            except:
                data=""
                pass
            edft=0
            pa.moveTo(50,50)
        else:
            edft=1
        return edft,data
    
    @staticmethod
    def appendlist(inp,out):
        for el in inp:
            out.append(el)
    
    @staticmethod
    def extract(a1,a2,strin):
        try:
            sr0=strin
            ia1=sr0.index(a1)+len(a1)
            ia2=sr0[ia1:].index(a2)+ia1
            return sr0[ia1:ia2], ia2+len(a2)
        except:
            return 'EOF',-1
    
    @staticmethod
    def fullextract(a1,a2,html):
        output=''
        sr1=''
        sr0=html
        l=[]
        while(sr0!=sr1):
            sr1=sr0
            output,iax=ToolBoxMethods.extract(a1,a2,sr0)
            if output=='EOF':
    	        break
            l.append(output)
            sr0=sr0[iax:]
        return l
    
    @staticmethod
    def extractall(html,route):#Arguments: (Text for extraction, Dictionary in the form {colum_name1:[[delimiter1,delimiter2],[delimiter_inside_the_previous1_1,delimiter_inside_the_previous1_2],...],...})
        results=dict()
        for element in route:
            results[element]=ToolBoxMethods.deepextract(html,route,element)
        return results
    
    @staticmethod
    def deepextract(html,route,delimkey):#Arguments: (Text for extraction, Dictionary in the form {colum_name1:[[delimiter1,delimiter2],[delimiter_inside_the_previous1_1,delimiter_inside_the_previous1_2],...],...},result_key)
        delimkeyresults=list()
        if len(route[delimkey])>1:
            D=dict()
            D[route[delimkey][0][0]]=dict()
            D[route[delimkey][0][0]][0]=ToolBoxMethods.fullextract(route[delimkey][0][0],route[delimkey][0][1],html)
        else:
            delimkeyresults.extend(ToolBoxMethods.fullextract(route[delimkey][0][0],route[delimkey][0][1],html))
        for x in range(1,len(route[delimkey])):
            D[route[delimkey][0][0]][x]=list()
            for g in D[route[delimkey][0][0]][x-1]:
                s=ToolBoxMethods.fullextract(route[delimkey][x][0],route[delimkey][x][1],g)
                ToolBoxMethods.appendlist(s,D[route[delimkey][0][0]][x])
            if x==len(route[delimkey])-1:
                delimkeyresults.extend(D[route[delimkey][0][0]][x])
                del D
        return delimkeyresults
    
    @staticmethod
    def checkemptybackground(path,imagelist,foldernameending,closecommandbeginning):
        w=ToolBoxMethods.try_imgs(path,imagelist[0],0,foldernameending)
        if w==None:
            g=ToolBoxMethods.try_imgs(path,imagelist[1],0,foldernameending)
            h=ToolBoxMethods.try_imgs(path,imagelist[2],0,foldernameending)
            if any([g!=None,h!=None]):
                pass
            else:
                pa.keyDown(closecommandbeginning)
                pa.keyDown("f4")
                pa.keyUp(closecommandbeginning)
                pa.keyUp("f4")
        else:
            pa.keyDown(closecommandbeginning)
            pa.keyDown("f4")
            pa.keyUp(closecommandbeginning)
            pa.keyUp("f4")
