import tkinter as tk
from tkinter import filedialog
import pandas as pd
import numpy as numpy
import os, sys
from docxtpl import DocxTemplate
from tkinter import ttk
from tkinter import *
from tkinter.messagebox import showinfo



#os.chdir(sys.path[0])
#win = Tk()
##Label(win, text=" Select the monthly data to trasnform into tables: ", font=('Helvetica 14 bold')).pack(pady=20)
#win_width = 750
#win_height = 750
#x = int(int(win.winfo_screenwidth()/2) - int(win_width/2))
#y = int(int(win.winfo_screenheight()/2) - int(win_height/2))
#win.geometry(f"{win_width}x{win_height}+{x}+{y}")

root1 = Tk()
root1.title('Select monthly dataset')
root1.geometry('750x250')
#file_path = filedialog.askopenfilename()




def open():
    file_path = filedialog.askopenfilename()
    df = pd.read_excel(file_path,header=None)
    doc = DocxTemplate('C:/Users/ChaudhryU/Desktop/finishedScript/dependencies/Template.docx')
    try:
        doc.save('Template_Rendered.docx')
    except PermissionError:
        root2 = Tk()
        root2.title("Error!")
        Label(root2, text=" Pleaes close previous Template_Rendered before continuing! ", font=('Helvetica 14 bold')).pack(pady=40)
        root2_width = 750
        root2_height = 250
        x = int(int(root2.winfo_screenwidth()/2) - int(root2_width/2))
        y = int(int(root2.winfo_screenheight()/2) - int(root2_height/2))
        root2.geometry(f"{root2_width}x{root2_height}+{x}+{y}")
        root2.mainloop()
        root2.destroy()




    df = df.loc[df[0].str.len() == 8]


    results = df.iloc[:, :9]

    results[['line', 'location']] = results[1].str.split(pat ='-', expand = True)
    results[['comm', 'points']] = results[2].str.split(pat =' - PM - ', expand = True)


    print(results['points'].value_counts())

    #-------------------------------------------------------------------------------------------------
    EXPO_completed = len(results[(results.line == 'EXPO') & (results[8] == 'Yes')])
    EXPO_late = len(results[(results.line == 'EXPO') & (results[8] == 'Due')])
    EXPO_total = len(results[results.line == 'EXPO'])

    try:
        EXPO_late_percentage = "{:.0%}".format(EXPO_late/ EXPO_total)
    except ZeroDivisionError:
        EXPO_late_percentage = "{:.0%}".format(0)
        
    try:
        EXPO_complete_percentage ="{:.0%}".format(EXPO_completed/ EXPO_total)
    except ZeroDivisionError:
        EXPO_complete_perecentage = "{:.0%}".format(0)
        
    #-------------------------------------------------------------------------------------------------
    MBL_completed = len(results[(results.line == 'MBL') & (results[8] == 'Yes')])
    MBL_late = len(results[(results.line == 'MBL') & (results[8] == 'Due')])
    MBL_total = len(results[results.line == 'MBL'])

    try:
        MBL_late_percentage = "{:.0%}".format(MBL_late/ MBL_total)
    except ZeroDivisionError:
        MBL_late_percentage = "{:.0%}".format(0)
    try:
        MBL_complete_percentage ="{:.0%}".format(MBL_completed/ MBL_total)
    except ZeroDivisionError:
        MBL_complete_percentage = "{:.0%}".format(0)
        
    #-------------------------------------------------------------------------------------------------
    MRL_MPL_completed = len(results[(results.line.str.contains('MRL|MPL') & (results[8] == 'Yes'))])
    MRL_MPL_late = len(results[(results.line.str.contains('MRL|MPL') & (results[8] == 'Due'))])
    MRL_MPL_total = len(results[results.line.str.contains('MRL|MPL', na=False)] ) 


    try:
        MRL_MPL_late_percentage = "{:.0%}".format(MRL_MPL_late/ MRL_MPL_total)
    except ZeroDivisionError:
        MRL_MPL_late_percentage = "{:.0%}".format(0)
        
    try:
        MRL_MPL_complete_percentage ="{:.0%}".format(MRL_MPL_completed/ MRL_MPL_total)
    except ZeroDivisionError:
        MRL_MPL_complete_perecentage = "{:.0%}".format(0)
        
        
    #-------------------------------------------------------------------------------------------------
    MGL_completed = len(results[(results.line == 'MGL') & (results[8] == 'Yes')])
    MGL_late = len(results[(results.line == 'MGL') & (results[8] == 'Due')])
    MGL_total = len(results[results.line == 'MGL'])

    try:
        MGL_late_percentage = "{:.0%}".format(MGL_late/ MGL_total)
    except ZeroDivisionError:
        MGL_late_percentage = "{:.0%}".format(0)
    try:
        MGL_complete_percentage ="{:.0%}".format(MGL_completed/ MGL_total)
    except ZeroDivisionError:
        MGL_complete_perecentage = "{:.0%}".format(0)
        


    #-------------------------------------------------------------------------------------------------
    PGL_completed = len(results[(results.line == 'PGL') & (results[8] == 'Yes')])
    PGL_late = len(results[(results.line == 'PGL') & (results[8] == 'Due')])
    PGL_total =  len(results[results.line == 'PGL'])

    try:
        PGL_late_percentage = "{:.0%}".format(PGL_late/ PGL_total)
    except ZeroDivisionError:
        PGL_late_percentage = "{:.0%}".format(0)
    try:
        PGL_complete_percentage ="{:.0%}".format(PGL_completed/ PGL_total)
    except ZeroDivisionError:
        PGL_complete_perecentage = "{:.0%}".format(0)
        


    #-------------------------------------------------------------------------------------------------
    CRENSHAW_completed = len(results[(results.line == 'CRENSHAW') & (results[8] == 'Yes')])
    CRENSHAW_late = len(results[(results.line == 'CRENSHAW') & (results[8] == 'Due')])
    CRENSHAW_total = len(results[results.line == "CRENSHAW"])

    try:
        CRENSHAW_late_percentage = "{:.0%}".format(CRENSHAW_late/ CRENSHAW_total)
    except ZeroDivisionError:
        CRENSHAW_late_percentage = "{:.0%}".format(0)
    try:
        CRENSHAW_complete_percentage ="{:.0%}".format(CRENSHAW_completed/ CRENSHAW_total)
    except ZeroDivisionError:
        CRENSHAW_complete_percentage = "{:.0%}".format(0)
        

    #-------------------------------------------------------------------------------------------------
    MOL_completed = len(results[(results.line == 'MOL') & (results[8] == 'Yes')])
    MOL_late = len(results[(results.line == 'MOL') & (results[8] == 'Due')])
    MOL_total = len(results[results.line == "MOL"])

    try:
        MOL_late_percentage = "{:.0%}".format(MOL_late/ MOL_total)
    except ZeroDivisionError:
        MOL_late_percentage = "{:.0%}".format(0)
    try:
        MOL_complete_percentage ="{:.0%}".format(MOL_completed/ MOL_total)
    except ZeroDivisionError:
        MOL_complete_perecentage = "{:.0%}".format(0)


    #-------------------------------------------------------------------------------------------------
    OTHER_completed = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results[8] == 'Yes')])  
    OTHER_late = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results[8] == 'Due')])  
    OTHER_total = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False)] ) 

    try:
        OTHER_late_percentage = "{:.0%}".format(OTHER_late/ OTHER_total)
    except ZeroDivisionError:
        OTHER_late_percentage = "{:.0%}".format(0)
    try:
        OTHER_complete_percentage ="{:.0%}".format(OTHER_completed/ OTHER_total)
    except ZeroDivisionError:
        OTHER_complete_percentage = "{:.0%}".format(0)
        
        
        
        

        
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------











    CCTV_MBL_sched = len(results[(results.line == 'MBL') & (results.points == 'CCTV')])
    CCTV_MBL_comp = len(results[(results.line == 'MBL') & (results.points == 'CCTV') & (results[8] == 'Yes')])
    CCTV_MBL_late = len(results[(results.line == 'MBL') & (results.points == 'CCTV') & (results[8] == 'Due')])
    try:
        CCTV_MBL_late_percentage = "{:.0%}".format(CCTV_MBL_late/ CCTV_MBL_sched)
    except ZeroDivisionError:
        CCTV_MBL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    CCTV_MRL_MPL_sched = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'CCTV'))])
    CCTV_MRL_MPL_comp = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'CCTV')& (results[8] == 'Yes'))])
    CCTV_MRL_MPL_late = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'CCTV')& (results[8] == 'Due'))])
    try:
        CCTV_MRL_MPL_late_percentage = "{:.0%}".format(CCTV_MRL_MPL_late/ CCTV_MRL_MPL_sched)
    except ZeroDivisionError:
        CCTV_MRL_MPL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    CCTV_MGL_sched = len(results[(results.line == 'MGL') & (results.points == 'CCTV')])
    CCTV_MGL_comp = len(results[(results.line == 'MGL') & (results.points == 'CCTV') & (results[8] == 'Yes')])
    CCTV_MGL_late = len(results[(results.line == 'MGL') & (results.points == 'CCTV') & (results[8] == 'Due')])
    try:
        CCTV_MGL_late_percentage = "{:.0%}".format(CCTV_MGL_late/ CCTV_MGL_sched)
    except ZeroDivisionError:
        CCTV_MGL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    CCTV_PGL_sched = len(results[(results.line == 'PGL') & (results.points == 'CCTV')])
    CCTV_PGL_comp = len(results[(results.line == 'PGL') & (results.points == 'CCTV') & (results[8] == 'Yes')])
    CCTV_PGL_late = len(results[(results.line == 'PGL') & (results.points == 'CCTV') & (results[8] == 'Due')])
    try:
        CCTV_PGL_late_percentage = "{:.0%}".format(CCTV_PGL_late/ CCTV_PGL_sched)
    except ZeroDivisionError:
        CCTV_PGL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    CCTV_EXPO_sched = len(results[(results.line == 'EXPO') & (results.points == 'CCTV')])
    CCTV_EXPO_comp = len(results[(results.line == 'EXPO') & (results.points == 'CCTV') & (results[8] == 'Yes')])
    CCTV_EXPO_late = len(results[(results.line == 'EXPO') & (results.points == 'CCTV') & (results[8] == 'Due')])
    try:
        CCTV_EXPO_late_percentage = "{:.0%}".format(CCTV_EXPO_late/ CCTV_EXPO_sched)
    except ZeroDivisionError:
        CCTV_EXPO_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    CCTV_MOL_sched = len(results[(results.line == 'MOL') & (results.points == 'CCTV')])
    CCTV_MOL_comp = len(results[(results.line == 'MOL') & (results.points == 'CCTV') & (results[8] == 'Yes')])
    CCTV_MOL_late = len(results[(results.line == 'MOL') & (results.points == 'CCTV') & (results[8] == 'Due')])
    try:
        CCTV_MOL_late_percentage = "{:.0%}".format(CCTV_MOL_late/ CCTV_MOL_sched)
    except ZeroDivisionError:
        CCTV_MOL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    CCTV_OTHER_sched = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results.points == 'CCTV')])
    CCTV_OTHER_comp = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results[8] == 'Yes') & (results.points == 'CCTV')])
    CCTV_OTHER_late = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results[8] == 'Due') & (results.points == 'CCTV')])
    try:
        CCTV_OTHER_late_percentage = "{:.0%}".format(CCTV_OTHER_late/ CCTV_OTHER_sched)
    except ZeroDivisionError:
        CCTV_OTHER_late_percentage = "{:.0%}".format(0)  
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------    















    CTS_MBL_sched = len(results[(results.line == 'MBL') & (results.points == 'CABLE TRANSMISSION SYSTEM')])
    CTS_MBL_comp = len(results[(results.line == 'MBL') & (results.points == 'CABLE TRANSMISSION SYSTEM') & (results[8] == 'Yes')])
    CTS_MBL_late = len(results[(results.line == 'MBL') & (results.points == 'CABLE TRANSMISSION SYSTEM') & (results[8] == 'Due')])
    try:
        CTS_MBL_late_percentage = "{:.0%}".format(CTS_MBL_late/ CTS_MBL_sched)
    except ZeroDivisionError:
        CTS_MBL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    CTS_MRL_MPL_sched = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'CABLE TRANSMISSION SYSTEM'))])
    CTS_MRL_MPL_comp = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'CABLE TRANSMISSION SYSTEM')& (results[8] == 'Yes'))])
    CTS_MRL_MPL_late = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'CABLE TRANSMISSION SYSTEM')& (results[8] == 'Due'))])
    try:
        CTS_MRL_MPL_late_percentage = "{:.0%}".format(CTS_MRL_MPL_late/ CTS_MRL_MPL_sched)
    except ZeroDivisionError:
        CTS_MRL_MPL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    CTS_MGL_sched = len(results[(results.line == 'MGL') & (results.points == 'CABLE TRANSMISSION SYSTEM')])
    CTS_MGL_comp = len(results[(results.line == 'MGL') & (results.points == 'CABLE TRANSMISSION SYSTEM') & (results[8] == 'Yes')])
    CTS_MGL_late = len(results[(results.line == 'MGL') & (results.points == 'CABLE TRANSMISSION SYSTEM') & (results[8] == 'Due')])
    try:
        CTS_MGL_late_percentage = "{:.0%}".format(CTS_MGL_late/ CTS_MGL_sched)
    except ZeroDivisionError:
        CTS_MGL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    CTS_PGL_sched = len(results[(results.line == 'PGL') & (results.points == 'CABLE TRANSMISSION SYSTEM')])
    CTS_PGL_comp = len(results[(results.line == 'PGL') & (results.points == 'CABLE TRANSMISSION SYSTEM') & (results[8] == 'Yes')])
    CTS_PGL_late = len(results[(results.line == 'PGL') & (results.points == 'CABLE TRANSMISSION SYSTEM') & (results[8] == 'Due')])
    try:
        CTS_PGL_late_percentage = "{:.0%}".format(CTS_PGL_late/ CTS_PGL_sched)
    except ZeroDivisionError:
        CTS_PGL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    CTS_EXPO_sched = len(results[(results.line == 'EXPO') & (results.points == 'CABLE TRANSMISSION SYSTEM')])
    CTS_EXPO_comp = len(results[(results.line == 'EXPO') & (results.points == 'CABLE TRANSMISSION SYSTEM') & (results[8] == 'Yes')])
    CTS_EXPO_late = len(results[(results.line == 'EXPO') & (results.points == 'CABLE TRANSMISSION SYSTEM') & (results[8] == 'Due')])
    try:
        CTS_EXPO_late_percentage = "{:.0%}".format(CTS_EXPO_late/ CTS_EXPO_sched)
    except ZeroDivisionError:
        CTS_EXPO_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    CTS_MOL_sched = len(results[(results.line == 'MOL') & (results.points == 'CABLE TRANSMISSION SYSTEM')])
    CTS_MOL_comp = len(results[(results.line == 'MOL') & (results.points == 'CABLE TRANSMISSION SYSTEM') & (results[8] == 'Yes')])
    CTS_MOL_late = len(results[(results.line == 'MOL') & (results.points == 'CABLE TRANSMISSION SYSTEM') & (results[8] == 'Due')])
    try:
        CTS_MOL_late_percentage = "{:.0%}".format(CTS_MOL_late/ CTS_MOL_sched)
    except ZeroDivisionError:
        CTS_MOL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    CTS_OTHER_sched = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results.points == 'CABLE TRANSMISSION SYSTEM')])
    CTS_OTHER_comp = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results[8] == 'Yes') & (results.points == 'CABLE TRANSMISSION SYSTEM')])
    CTS_OTHER_late = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results[8] == 'Due') & (results.points == 'CABLE TRANSMISSION SYSTEM')])
    try:
        CTS_OTHER_late_percentage = "{:.0%}".format(CTS_OTHER_late/ CTS_OTHER_sched)
    except ZeroDivisionError:
        CTS_OTHER_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------    












    TEL_MBL_sched = len(results[(results.line == 'MBL') & (results.points == 'TELEPHONE (RAIL COMM)')])
    TEL_MBL_comp = len(results[(results.line == 'MBL') & (results.points == 'TELEPHONE (RAIL COMM)') & (results[8] == 'Yes')])
    TEL_MBL_late = len(results[(results.line == 'MBL') & (results.points == 'TELEPHONE (RAIL COMM)') & (results[8] == 'Due')])
    try:
        TEL_MBL_late_percentage = "{:.0%}".format(TEL_MBL_late/ TEL_MBL_sched)
    except ZeroDivisionError:
        TEL_MBL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    TEL_MRL_MPL_sched = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'TELEPHONE (RAIL COMM)'))])
    TEL_MRL_MPL_comp = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'TELEPHONE (RAIL COMM)')& (results[8] == 'Yes'))])
    TEL_MRL_MPL_late = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'TELEPHONE (RAIL COMM)')& (results[8] == 'Due'))])
    try:
        TEL_MRL_MPL_late_percentage = "{:.0%}".format(TEL_MRL_MPL_late/ TEL_MRL_MPL_sched)
    except ZeroDivisionError:
        TEL_MRL_MPL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    TEL_MGL_sched = len(results[(results.line == 'MGL') & (results.points == 'TELEPHONE (RAIL COMM)')])
    TEL_MGL_comp = len(results[(results.line == 'MGL') & (results.points == 'TELEPHONE (RAIL COMM)') & (results[8] == 'Yes')])
    TEL_MGL_late = len(results[(results.line == 'MGL') & (results.points == 'TELEPHONE (RAIL COMM)') & (results[8] == 'Due')])
    try:
        TEL_MGL_late_percentage = "{:.0%}".format(TEL_MGL_late/ TEL_MGL_sched)
    except ZeroDivisionError:
        TEL_MGL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    TEL_PGL_sched = len(results[(results.line == 'PGL') & (results.points == 'TELEPHONE (RAIL COMM)')])
    TEL_PGL_comp = len(results[(results.line == 'PGL') & (results.points == 'TELEPHONE (RAIL COMM)') & (results[8] == 'Yes')])
    TEL_PGL_late = len(results[(results.line == 'PGL') & (results.points == 'TELEPHONE (RAIL COMM)') & (results[8] == 'Due')])
    try:
        TEL_PGL_late_percentage = "{:.0%}".format(TEL_PGL_late/ TEL_PGL_sched)
    except ZeroDivisionError:
        TEL_PGL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    TEL_EXPO_sched = len(results[(results.line == 'EXPO') & (results.points == 'TELEPHONE (RAIL COMM)')])
    TEL_EXPO_comp = len(results[(results.line == 'EXPO') & (results.points == 'TELEPHONE (RAIL COMM)') & (results[8] == 'Yes')])
    TEL_EXPO_late = len(results[(results.line == 'EXPO') & (results.points == 'TELEPHONE (RAIL COMM)') & (results[8] == 'Due')])
    try:
        TEL_EXPO_late_percentage = "{:.0%}".format(TEL_EXPO_late/ TEL_EXPO_sched)
    except ZeroDivisionError:
        TEL_EXPO_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    TEL_MOL_sched = len(results[(results.line == 'MOL') & (results.points == 'TELEPHONE (RAIL COMM)')])
    TEL_MOL_comp = len(results[(results.line == 'MOL') & (results.points == 'TELEPHONE (RAIL COMM)') & (results[8] == 'Yes')])
    TEL_MOL_late = len(results[(results.line == 'MOL') & (results.points == 'TELEPHONE (RAIL COMM)') & (results[8] == 'Due')])
    try:
        TEL_MOL_late_percentage = "{:.0%}".format(TEL_MOL_late/ TEL_MOL_sched)
    except ZeroDivisionError:
        TEL_MOL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    TEL_OTHER_sched = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results.points == 'TELEPHONE (RAIL COMM)')])
    TEL_OTHER_comp = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results[8] == 'Yes') & (results.points == 'TELEPHONE (RAIL COMM)')])
    TEL_OTHER_late = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results[8] == 'Due') & (results.points == 'TELEPHONE (RAIL COMM)')])
    try:
        TEL_OTHER_late_percentage = "{:.0%}".format(TEL_OTHER_late/ TEL_OTHER_sched)
    except ZeroDivisionError:
        TEL_OTHER_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------  








    PA_MBL_sched = len(results[(results.line == 'MBL') & (results.points == 'PUBLIC ADDRESS SYSTEM')])
    PA_MBL_comp = len(results[(results.line == 'MBL') & (results.points == 'PUBLIC ADDRESS SYSTEM') & (results[8] == 'Yes')])
    PA_MBL_late = len(results[(results.line == 'MBL') & (results.points == 'PUBLIC ADDRESS SYSTEM') & (results[8] == 'Due')])
    try:
        PA_MBL_late_percentage = "{:.0%}".format(PA_MBL_late/ PA_MBL_sched)
    except ZeroDivisionError:
        PA_MBL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    PA_MRL_MPL_sched = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'PUBLIC ADDRESS SYSTEM'))])
    PA_MRL_MPL_comp = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'PUBLIC ADDRESS SYSTEM')& (results[8] == 'Yes'))])
    PA_MRL_MPL_late = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'PUBLIC ADDRESS SYSTEM')& (results[8] == 'Due'))])
    try:
        PA_MRL_MPL_late_percentage = "{:.0%}".format(PA_MRL_MPL_late/ PA_MRL_MPL_sched)
    except ZeroDivisionError:
        PA_MRL_MPL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    PA_MGL_sched = len(results[(results.line == 'MGL') & (results.points == 'PUBLIC ADDRESS SYSTEM')])
    PA_MGL_comp = len(results[(results.line == 'MGL') & (results.points == 'PUBLIC ADDRESS SYSTEM') & (results[8] == 'Yes')])
    PA_MGL_late = len(results[(results.line == 'MGL') & (results.points == 'PUBLIC ADDRESS SYSTEM') & (results[8] == 'Due')])
    try:
        PA_MGL_late_percentage = "{:.0%}".format(PA_MGL_late/ PA_MGL_sched)
    except ZeroDivisionError:
        PA_MGL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    PA_PGL_sched = len(results[(results.line == 'PGL') & (results.points == 'PUBLIC ADDRESS SYSTEM')])
    PA_PGL_comp = len(results[(results.line == 'PGL') & (results.points == 'PUBLIC ADDRESS SYSTEM') & (results[8] == 'Yes')])
    PA_PGL_late = len(results[(results.line == 'PGL') & (results.points == 'PUBLIC ADDRESS SYSTEM') & (results[8] == 'Due')])
    try:
        PA_PGL_late_percentage = "{:.0%}".format(PA_PGL_late/ PA_PGL_sched)
    except ZeroDivisionError:
        PA_PGL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    PA_EXPO_sched = len(results[(results.line == 'EXPO') & (results.points == 'PUBLIC ADDRESS SYSTEM')])
    PA_EXPO_comp = len(results[(results.line == 'EXPO') & (results.points == 'PUBLIC ADDRESS SYSTEM') & (results[8] == 'Yes')])
    PA_EXPO_late = len(results[(results.line == 'EXPO') & (results.points == 'PUBLIC ADDRESS SYSTEM') & (results[8] == 'Due')])
    try:
        PA_EXPO_late_percentage = "{:.0%}".format(PA_EXPO_late/ PA_EXPO_sched)
    except ZeroDivisionError:
        PA_EXPO_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    PA_MOL_sched = len(results[(results.line == 'MOL') & (results.points == 'PUBLIC ADDRESS SYSTEM')])
    PA_MOL_comp = len(results[(results.line == 'MOL') & (results.points == 'PUBLIC ADDRESS SYSTEM') & (results[8] == 'Yes')])
    PA_MOL_late = len(results[(results.line == 'MOL') & (results.points == 'PUBLIC ADDRESS SYSTEM') & (results[8] == 'Due')])
    try:
        PA_MOL_late_percentage = "{:.0%}".format(PA_MOL_late/ PA_MOL_sched)
    except ZeroDivisionError:
        PA_MOL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    PA_OTHER_sched = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results.points == 'PUBLIC ADDRESS SYSTEM')])
    PA_OTHER_comp = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results[8] == 'Yes') & (results.points == 'PUBLIC ADDRESS SYSTEM')])
    PA_OTHER_late = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results[8] == 'Due') & (results.points == 'PUBLIC ADDRESS SYSTEM')])
    try:
        PA_OTHER_late_percentage = "{:.0%}".format(PA_OTHER_late/ PA_OTHER_sched)
    except ZeroDivisionError:
        PA_OTHER_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #------------------------------------------------------------------------------------------------- 



    VMS_MBL_sched = len(results[(results.line == 'MBL') & (results.points == 'VARIABLE MESSAGE SYSTEM')])
    VMS_MBL_comp = len(results[(results.line == 'MBL') & (results.points == 'VARIABLE MESSAGE SYSTEM') & (results[8] == 'Yes')])
    VMS_MBL_late = len(results[(results.line == 'MBL') & (results.points == 'VARIABLE MESSAGE SYSTEM') & (results[8] == 'Due')])
    try:
        VMS_MBL_late_percentage = "{:.0%}".format(VMS_MBL_late/ VMS_MBL_sched)
    except ZeroDivisionError:
        VMS_MBL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    VMS_MRL_MPL_sched = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'VARIABLE MESSAGE SYSTEM'))])
    VMS_MRL_MPL_comp = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'VARIABLE MESSAGE SYSTEM')& (results[8] == 'Yes'))])
    VMS_MRL_MPL_late = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'VARIABLE MESSAGE SYSTEM')& (results[8] == 'Due'))])
    try:
        VMS_MRL_MPL_late_percentage = "{:.0%}".format(VMS_MRL_MPL_late/ VMS_MRL_MPL_sched)
    except ZeroDivisionError:
        VMS_MRL_MPL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    VMS_MGL_sched = len(results[(results.line == 'MGL') & (results.points == 'VARIABLE MESSAGE SYSTEM')])
    VMS_MGL_comp = len(results[(results.line == 'MGL') & (results.points == 'VARIABLE MESSAGE SYSTEM') & (results[8] == 'Yes')])
    VMS_MGL_late = len(results[(results.line == 'MGL') & (results.points == 'VARIABLE MESSAGE SYSTEM') & (results[8] == 'Due')])
    try:
        VMS_MGL_late_percentage = "{:.0%}".format(VMS_MGL_late/ VMS_MGL_sched)
    except ZeroDivisionError:
        VMS_MGL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    VMS_PGL_sched = len(results[(results.line == 'PGL') & (results.points == 'VARIABLE MESSAGE SYSTEM')])
    VMS_PGL_comp = len(results[(results.line == 'PGL') & (results.points == 'VARIABLE MESSAGE SYSTEM') & (results[8] == 'Yes')])
    VMS_PGL_late = len(results[(results.line == 'PGL') & (results.points == 'VARIABLE MESSAGE SYSTEM') & (results[8] == 'Due')])
    try:
        VMS_PGL_late_percentage = "{:.0%}".format(VMS_PGL_late/ VMS_PGL_sched)
    except ZeroDivisionError:
        VMS_PGL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    VMS_EXPO_sched = len(results[(results.line == 'EXPO') & (results.points == 'VARIABLE MESSAGE SYSTEM')])
    VMS_EXPO_comp = len(results[(results.line == 'EXPO') & (results.points == 'VARIABLE MESSAGE SYSTEM') & (results[8] == 'Yes')])
    VMS_EXPO_late = len(results[(results.line == 'EXPO') & (results.points == 'VARIABLE MESSAGE SYSTEM') & (results[8] == 'Due')])
    try:
        VMS_EXPO_late_percentage = "{:.0%}".format(VMS_EXPO_late/ VMS_EXPO_sched)
    except ZeroDivisionError:
        VMS_EXPO_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    VMS_MOL_sched = len(results[(results.line == 'MOL') & (results.points == 'VARIABLE MESSAGE SYSTEM')])
    VMS_MOL_comp = len(results[(results.line == 'MOL') & (results.points == 'VARIABLE MESSAGE SYSTEM') & (results[8] == 'Yes')])
    VMS_MOL_late = len(results[(results.line == 'MOL') & (results.points == 'VARIABLE MESSAGE SYSTEM') & (results[8] == 'Due')])
    try:
        VMS_MOL_late_percentage = "{:.0%}".format(VMS_MOL_late/ VMS_MOL_sched)
    except ZeroDivisionError:
        VMS_MOL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    VMS_OTHER_sched = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results.points == 'VARIABLE MESSAGE SYSTEM')])
    VMS_OTHER_comp = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results[8] == 'Yes') & (results.points == 'VARIABLE MESSAGE SYSTEM')])
    VMS_OTHER_late = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results[8] == 'Due') & (results.points == 'VARIABLE MESSAGE SYSTEM')])
    try:
        VMS_OTHER_late_percentage = "{:.0%}".format(VMS_OTHER_late/ VMS_OTHER_sched)
    except ZeroDivisionError:
        VMS_OTHER_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #------------------------------------------------------------------------------------------------- 






    RTU_MBL_sched = len(results[(results.line == 'MBL') & (results.points == 'RTU')])
    RTU_MBL_comp = len(results[(results.line == 'MBL') & (results.points == 'RTU') & (results[8] == 'Yes')])
    RTU_MBL_late = len(results[(results.line == 'MBL') & (results.points == 'RTU') & (results[8] == 'Due')])
    try:
        RTU_MBL_late_percentage = "{:.0%}".format(RTU_MBL_late/ RTU_MBL_sched)
    except ZeroDivisionError:
        RTU_MBL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    RTU_MRL_MPL_sched = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'RTU'))])
    RTU_MRL_MPL_comp = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'RTU')& (results[8] == 'Yes'))])
    RTU_MRL_MPL_late = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'RTU')& (results[8] == 'Due'))])
    try:
        RTU_MRL_MPL_late_percentage = "{:.0%}".format(RTU_MRL_MPL_late/ RTU_MRL_MPL_sched)
    except ZeroDivisionError:
        RTU_MRL_MPL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    RTU_MGL_sched = len(results[(results.line == 'MGL') & (results.points == 'RTU')])
    RTU_MGL_comp = len(results[(results.line == 'MGL') & (results.points == 'RTU') & (results[8] == 'Yes')])
    RTU_MGL_late = len(results[(results.line == 'MGL') & (results.points == 'RTU') & (results[8] == 'Due')])
    try:
        RTU_MGL_late_percentage = "{:.0%}".format(RTU_MGL_late/ RTU_MGL_sched)
    except ZeroDivisionError:
        RTU_MGL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    RTU_PGL_sched = len(results[(results.line == 'PGL') & (results.points == 'RTU')])
    RTU_PGL_comp = len(results[(results.line == 'PGL') & (results.points == 'RTU') & (results[8] == 'Yes')])
    RTU_PGL_late = len(results[(results.line == 'PGL') & (results.points == 'RTU') & (results[8] == 'Due')])
    try:
        RTU_PGL_late_percentage = "{:.0%}".format(RTU_PGL_late/ RTU_PGL_sched)
    except ZeroDivisionError:
        RTU_PGL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    RTU_EXPO_sched = len(results[(results.line == 'EXPO') & (results.points == 'RTU')])
    RTU_EXPO_comp = len(results[(results.line == 'EXPO') & (results.points == 'RTU') & (results[8] == 'Yes')])
    RTU_EXPO_late = len(results[(results.line == 'EXPO') & (results.points == 'RTU') & (results[8] == 'Due')])
    try:
        RTU_EXPO_late_percentage = "{:.0%}".format(RTU_EXPO_late/ RTU_EXPO_sched)
    except ZeroDivisionError:
        RTU_EXPO_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    RTU_MOL_sched = len(results[(results.line == 'MOL') & (results.points == 'RTU')])
    RTU_MOL_comp = len(results[(results.line == 'MOL') & (results.points == 'RTU') & (results[8] == 'Yes')])
    RTU_MOL_late = len(results[(results.line == 'MOL') & (results.points == 'RTU') & (results[8] == 'Due')])
    try:
        RTU_MOL_late_percentage = "{:.0%}".format(RTU_MOL_late/ RTU_MOL_sched)
    except ZeroDivisionError:
        RTU_MOL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    RTU_OTHER_sched = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results.points == 'RTU')])
    RTU_OTHER_comp = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results[8] == 'Yes') & (results.points == 'RTU')])
    RTU_OTHER_late = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results[8] == 'Due') & (results.points == 'RTU')])
    try:
        RTU_OTHER_late_percentage = "{:.0%}".format(RTU_OTHER_late/ RTU_OTHER_sched)
    except ZeroDivisionError:
        RTU_OTHER_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #------------------------------------------------------------------------------------------------- 



    PLC_MBL_sched = len(results[(results.line == 'MBL') & (results.points == 'PLC')])
    PLC_MBL_comp = len(results[(results.line == 'MBL') & (results.points == 'PLC') & (results[8] == 'Yes')])
    PLC_MBL_late = len(results[(results.line == 'MBL') & (results.points == 'PLC') & (results[8] == 'Due')])
    try:
        PLC_MBL_late_percentage = "{:.0%}".format(PLC_MBL_late/ PLC_MBL_sched)
    except ZeroDivisionError:
        PLC_MBL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    PLC_MRL_MPL_sched = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'PLC'))])
    PLC_MRL_MPL_comp = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'PLC')& (results[8] == 'Yes'))])
    PLC_MRL_MPL_late = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'PLC')& (results[8] == 'Due'))])
    try:
        PLC_MRL_MPL_late_percentage = "{:.0%}".format(PLC_MRL_MPL_late/ PLC_MRL_MPL_sched)
    except ZeroDivisionError:
        PLC_MRL_MPL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    PLC_MGL_sched = len(results[(results.line == 'MGL') & (results.points == 'PLC')])
    PLC_MGL_comp = len(results[(results.line == 'MGL') & (results.points == 'PLC') & (results[8] == 'Yes')])
    PLC_MGL_late = len(results[(results.line == 'MGL') & (results.points == 'PLC') & (results[8] == 'Due')])
    try:
        PLC_MGL_late_percentage = "{:.0%}".format(PLC_MGL_late/ PLC_MGL_sched)
    except ZeroDivisionError:
        PLC_MGL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    PLC_PGL_sched = len(results[(results.line == 'PGL') & (results.points == 'PLC')])
    PLC_PGL_comp = len(results[(results.line == 'PGL') & (results.points == 'PLC') & (results[8] == 'Yes')])
    PLC_PGL_late = len(results[(results.line == 'PGL') & (results.points == 'PLC') & (results[8] == 'Due')])
    try:
        PLC_PGL_late_percentage = "{:.0%}".format(PLC_PGL_late/ PLC_PGL_sched)
    except ZeroDivisionError:
        PLC_PGL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    PLC_EXPO_sched = len(results[(results.line == 'EXPO') & (results.points == 'PLC')])
    PLC_EXPO_comp = len(results[(results.line == 'EXPO') & (results.points == 'PLC') & (results[8] == 'Yes')])
    PLC_EXPO_late = len(results[(results.line == 'EXPO') & (results.points == 'PLC') & (results[8] == 'Due')])
    try:
        PLC_EXPO_late_percentage = "{:.0%}".format(PLC_EXPO_late/ PLC_EXPO_sched)
    except ZeroDivisionError:
        PLC_EXPO_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    PLC_MOL_sched = len(results[(results.line == 'MOL') & (results.points == 'PLC')])
    PLC_MOL_comp = len(results[(results.line == 'MOL') & (results.points == 'PLC') & (results[8] == 'Yes')])
    PLC_MOL_late = len(results[(results.line == 'MOL') & (results.points == 'PLC') & (results[8] == 'Due')])
    try:
        PLC_MOL_late_percentage = "{:.0%}".format(PLC_MOL_late/ PLC_MOL_sched)
    except ZeroDivisionError:
        PLC_MOL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    PLC_OTHER_sched = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results.points == 'PLC')])
    PLC_OTHER_comp = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results[8] == 'Yes') & (results.points == 'PLC')])
    PLC_OTHER_late = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results[8] == 'Due') & (results.points == 'PLC')])
    try:
        PLC_OTHER_late_percentage = "{:.0%}".format(PLC_OTHER_late/ PLC_OTHER_sched)
    except ZeroDivisionError:
        PLC_OTHER_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------


    FCS_MBL_sched = len(results[(results.line == 'MBL') & (results.points == 'FIRE CONTROL SYSTEM')])
    FCS_MBL_comp = len(results[(results.line == 'MBL') & (results.points == 'FIRE CONTROL SYSTEM') & (results[8] == 'Yes')])
    FCS_MBL_late = len(results[(results.line == 'MBL') & (results.points == 'FIRE CONTROL SYSTEM') & (results[8] == 'Due')])
    try:
        FCS_MBL_late_percentage = "{:.0%}".format(FCS_MBL_late/ FCS_MBL_sched)
    except ZeroDivisionError:
        FCS_MBL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    FCS_MRL_MPL_sched = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'FIRE CONTROL SYSTEM'))])
    FCS_MRL_MPL_comp = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'FIRE CONTROL SYSTEM')& (results[8] == 'Yes'))])
    FCS_MRL_MPL_late = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'FIRE CONTROL SYSTEM')& (results[8] == 'Due'))])
    try:
        FCS_MRL_MPL_late_percentage = "{:.0%}".format(FCS_MRL_MPL_late/ FCS_MRL_MPL_sched)
    except ZeroDivisionError:
        FCS_MRL_MPL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    FCS_MGL_sched = len(results[(results.line == 'MGL') & (results.points == 'FIRE CONTROL SYSTEM')])
    FCS_MGL_comp = len(results[(results.line == 'MGL') & (results.points == 'FIRE CONTROL SYSTEM') & (results[8] == 'Yes')])
    FCS_MGL_late = len(results[(results.line == 'MGL') & (results.points == 'FIRE CONTROL SYSTEM') & (results[8] == 'Due')])
    try:
        FCS_MGL_late_percentage = "{:.0%}".format(FCS_MGL_late/ FCS_MGL_sched)
    except ZeroDivisionError:
        FCS_MGL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    FCS_PGL_sched = len(results[(results.line == 'PGL') & (results.points == 'FIRE CONTROL SYSTEM')])
    FCS_PGL_comp = len(results[(results.line == 'PGL') & (results.points == 'FIRE CONTROL SYSTEM') & (results[8] == 'Yes')])
    FCS_PGL_late = len(results[(results.line == 'PGL') & (results.points == 'FIRE CONTROL SYSTEM') & (results[8] == 'Due')])
    try:
        FCS_PGL_late_percentage = "{:.0%}".format(FCS_PGL_late/ FCS_PGL_sched)
    except ZeroDivisionError:
        FCS_PGL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    FCS_EXPO_sched = len(results[(results.line == 'EXPO') & (results.points == 'FIRE CONTROL SYSTEM')])
    FCS_EXPO_comp = len(results[(results.line == 'EXPO') & (results.points == 'FIRE CONTROL SYSTEM') & (results[8] == 'Yes')])
    FCS_EXPO_late = len(results[(results.line == 'EXPO') & (results.points == 'FIRE CONTROL SYSTEM') & (results[8] == 'Due')])
    try:
        FCS_EXPO_late_percentage = "{:.0%}".format(FCS_EXPO_late/ FCS_EXPO_sched)
    except ZeroDivisionError:
        FCS_EXPO_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    FCS_MOL_sched = len(results[(results.line == 'MOL') & (results.points == 'FIRE CONTROL SYSTEM')])
    FCS_MOL_comp = len(results[(results.line == 'MOL') & (results.points == 'FIRE CONTROL SYSTEM') & (results[8] == 'Yes')])
    FCS_MOL_late = len(results[(results.line == 'MOL') & (results.points == 'FIRE CONTROL SYSTEM') & (results[8] == 'Due')])
    try:
        FCS_MOL_late_percentage = "{:.0%}".format(FCS_MOL_late/ FCS_MOL_sched)
    except ZeroDivisionError:
        FCS_MOL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    FCS_OTHER_sched = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results.points == 'FIRE CONTROL SYSTEM')])
    FCS_OTHER_comp = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results[8] == 'Yes') & (results.points == 'FIRE CONTROL SYSTEM')])
    FCS_OTHER_late = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results[8] == 'Due') & (results.points == 'FIRE CONTROL SYSTEM')])
    try:
        FCS_OTHER_late_percentage = "{:.0%}".format(FCS_OTHER_late/ FCS_OTHER_sched)
    except ZeroDivisionError:
        FCS_OTHER_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------




    TPIS_MBL_sched = len(results[(results.line == 'MBL') & (results.points == 'TPIS')])
    TPIS_MBL_comp = len(results[(results.line == 'MBL') & (results.points == 'TPIS') & (results[8] == 'Yes')])
    TPIS_MBL_late = len(results[(results.line == 'MBL') & (results.points == 'TPIS') & (results[8] == 'Due')])
    try:
        TPIS_MBL_late_percentage = "{:.0%}".format(TPIS_MBL_late/ TPIS_MBL_sched)
    except ZeroDivisionError:
        TPIS_MBL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    TPIS_MRL_MPL_sched = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'TPIS'))])
    TPIS_MRL_MPL_comp = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'TPIS')& (results[8] == 'Yes'))])
    TPIS_MRL_MPL_late = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'TPIS')& (results[8] == 'Due'))])
    try:
        TPIS_MRL_MPL_late_percentage = "{:.0%}".format(TPIS_MRL_MPL_late/ TPIS_MRL_MPL_sched)
    except ZeroDivisionError:
        TPIS_MRL_MPL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    TPIS_MGL_sched = len(results[(results.line == 'MGL') & (results.points == 'TPIS')])
    TPIS_MGL_comp = len(results[(results.line == 'MGL') & (results.points == 'TPIS') & (results[8] == 'Yes')])
    TPIS_MGL_late = len(results[(results.line == 'MGL') & (results.points == 'TPIS') & (results[8] == 'Due')])
    try:
        TPIS_MGL_late_percentage = "{:.0%}".format(TPIS_MGL_late/ TPIS_MGL_sched)
    except ZeroDivisionError:
        TPIS_MGL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    TPIS_PGL_sched = len(results[(results.line == 'PGL') & (results.points == 'TPIS')])
    TPIS_PGL_comp = len(results[(results.line == 'PGL') & (results.points == 'TPIS') & (results[8] == 'Yes')])
    TPIS_PGL_late = len(results[(results.line == 'PGL') & (results.points == 'TPIS') & (results[8] == 'Due')])
    try:
        TPIS_PGL_late_percentage = "{:.0%}".format(TPIS_PGL_late/ TPIS_PGL_sched)
    except ZeroDivisionError:
        TPIS_PGL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    TPIS_EXPO_sched = len(results[(results.line == 'EXPO') & (results.points == 'TPIS')])
    TPIS_EXPO_comp = len(results[(results.line == 'EXPO') & (results.points == 'TPIS') & (results[8] == 'Yes')])
    TPIS_EXPO_late = len(results[(results.line == 'EXPO') & (results.points == 'TPIS') & (results[8] == 'Due')])
    try:
        TPIS_EXPO_late_percentage = "{:.0%}".format(TPIS_EXPO_late/ TPIS_EXPO_sched)
    except ZeroDivisionError:
        TPIS_EXPO_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    TPIS_MOL_sched = len(results[(results.line == 'MOL') & (results.points == 'TPIS')])
    TPIS_MOL_comp = len(results[(results.line == 'MOL') & (results.points == 'TPIS') & (results[8] == 'Yes')])
    TPIS_MOL_late = len(results[(results.line == 'MOL') & (results.points == 'TPIS') & (results[8] == 'Due')])
    try:
        TPIS_MOL_late_percentage = "{:.0%}".format(TPIS_MOL_late/ TPIS_MOL_sched)
    except ZeroDivisionError:
        TPIS_MOL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    TPIS_OTHER_sched = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results.points == 'TPIS')])
    TPIS_OTHER_comp = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results[8] == 'Yes') & (results.points == 'TPIS')])
    TPIS_OTHER_late = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results[8] == 'Due') & (results.points == 'TPIS')])
    try:
        TPIS_OTHER_late_percentage = "{:.0%}".format(TPIS_OTHER_late/ TPIS_OTHER_sched)
    except ZeroDivisionError:
        TPIS_OTHER_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------







    BAT_RECT_MBL_sched = len(results[(results.line == 'MBL') & (results.points.str.contains('BATTERY|RECTIFIERS'))])
    BAT_RECT_MBL_comp = len(results[(results.line == 'MBL') & (results.points.str.contains('BATTERY|RECTIFIERS')) & (results[8] == 'Yes')])
    BAT_RECT_MBL_late = len(results[(results.line == 'MBL') & (results.points.str.contains('BATTERY|RECTIFIERS')) & (results[8] == 'Due')])
    try:
        BAT_RECT_MBL_late_percentage = "{:.0%}".format(BAT_RECT_MBL_late/ BAT_RECT_MBL_sched)
    except ZeroDivisionError:
        BAT_RECT_MBL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    BAT_RECT_MRL_MPL_sched = len(results[(results.line.str.contains('MRL|MPL') & (results.points.str.contains('BATTERY|RECTIFIERS')))])
    BAT_RECT_MRL_MPL_comp = len(results[(results.line.str.contains('MRL|MPL') & (results.points.str.contains('BATTERY|RECTIFIERS'))& (results[8] == 'Yes'))])
    BAT_RECT_MRL_MPL_late = len(results[(results.line.str.contains('MRL|MPL') & (results.points.str.contains('BATTERY|RECTIFIERS'))& (results[8] == 'Due'))])
    try:
        BAT_RECT_MRL_MPL_late_percentage = "{:.0%}".format(BAT_RECT_MRL_MPL_late/ BAT_RECT_MRL_MPL_sched)
    except ZeroDivisionError:
        BAT_RECT_MRL_MPL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    BAT_RECT_MGL_sched = len(results[(results.line == 'MGL') & (results.points.str.contains('BATTERY|RECTIFIERS'))])
    BAT_RECT_MGL_comp = len(results[(results.line == 'MGL') & (results.points.str.contains('BATTERY|RECTIFIERS')) & (results[8] == 'Yes')])
    BAT_RECT_MGL_late = len(results[(results.line == 'MGL') & (results.points.str.contains('BATTERY|RECTIFIERS')) & (results[8] == 'Due')])
    try:
        BAT_RECT_MGL_late_percentage = "{:.0%}".format(BAT_RECT_MGL_late/ BAT_RECT_MGL_sched)
    except ZeroDivisionError:
        BAT_RECT_MGL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    BAT_RECT_PGL_sched = len(results[(results.line == 'PGL') & (results.points.str.contains('BATTERY|RECTIFIERS'))])
    BAT_RECT_PGL_comp = len(results[(results.line == 'PGL') & (results.points.str.contains('BATTERY|RECTIFIERS')) & (results[8] == 'Yes')])
    BAT_RECT_PGL_late = len(results[(results.line == 'PGL') & (results.points.str.contains('BATTERY|RECTIFIERS')) & (results[8] == 'Due')])
    try:
        BAT_RECT_PGL_late_percentage = "{:.0%}".format(BAT_RECT_PGL_late/ BAT_RECT_PGL_sched)
    except ZeroDivisionError:
        BAT_RECT_PGL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    BAT_RECT_EXPO_sched = len(results[(results.line == 'EXPO') & (results.points.str.contains('BATTERY|RECTIFIERS'))])
    BAT_RECT_EXPO_comp = len(results[(results.line == 'EXPO') & (results.points.str.contains('BATTERY|RECTIFIERS')) & (results[8] == 'Yes')])
    BAT_RECT_EXPO_late = len(results[(results.line == 'EXPO') & (results.points.str.contains('BATTERY|RECTIFIERS')) & (results[8] == 'Due')])
    try:
        BAT_RECT_EXPO_late_percentage = "{:.0%}".format(BAT_RECT_EXPO_late/ BAT_RECT_EXPO_sched)
    except ZeroDivisionError:
        BAT_RECT_EXPO_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    BAT_RECT_MOL_sched = len(results[(results.line == 'MOL') & (results.points.str.contains('BATTERY|RECTIFIERS'))])
    BAT_RECT_MOL_comp = len(results[(results.line == 'MOL') & (results.points.str.contains('BATTERY|RECTIFIERS')) & (results[8] == 'Yes')])
    BAT_RECT_MOL_late = len(results[(results.line == 'MOL') & (results.points.str.contains('BATTERY|RECTIFIERS'))& (results[8] == 'Due')])
    try:
        BAT_RECT_MOL_late_percentage = "{:.0%}".format(BAT_RECT_MOL_late/ BAT_RECT_MOL_sched)
    except ZeroDivisionError:
        BAT_RECT_MOL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    BAT_RECT_OTHER_sched = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results.points.str.contains('BATTERY|RECTIFIERS'))])
    BAT_RECT_OTHER_comp = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results[8] == 'Yes') & (results.points.str.contains('BATTERY|RECTIFIERS'))])
    BAT_RECT_OTHER_late = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results[8] == 'Due') & (results.points.str.contains('BATTERY|RECTIFIERS'))])
    try:
        BAT_RECT_OTHER_late_percentage = "{:.0%}".format(BAT_RECT_OTHER_late/ BAT_RECT_OTHER_sched)
    except ZeroDivisionError:
        BAT_RECT_OTHER_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #------------------------------------------------------------------------------------------------- 


    GASANA_MBL_sched = len(results[(results.line == 'MBL') & (results.points == 'GAS ANALYZER')])
    GASANA_MBL_comp = len(results[(results.line == 'MBL') & (results.points == 'GAS ANALYZER') & (results[8] == 'Yes')])
    GASANA_MBL_late = len(results[(results.line == 'MBL') & (results.points == 'GAS ANALYZER') & (results[8] == 'Due')])
    try:
        GASANA_MBL_late_percentage = "{:.0%}".format(GASANA_MBL_late/ GASANA_MBL_sched)
    except ZeroDivisionError:
        GASANA_MBL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    GASANA_MRL_MPL_sched = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'GAS ANALYZER'))])
    GASANA_MRL_MPL_comp = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'GAS ANALYZER')& (results[8] == 'Yes'))])
    GASANA_MRL_MPL_late = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'GAS ANALYZER')& (results[8] == 'Due'))])
    try:
        GASANA_MRL_MPL_late_percentage = "{:.0%}".format(GASANA_MRL_MPL_late/ GASANA_MRL_MPL_sched)
    except ZeroDivisionError:
        GASANA_MRL_MPL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    GASANA_MGL_sched = len(results[(results.line == 'MGL') & (results.points == 'GAS ANALYZER')])
    GASANA_MGL_comp = len(results[(results.line == 'MGL') & (results.points == 'GAS ANALYZER') & (results[8] == 'Yes')])
    GASANA_MGL_late = len(results[(results.line == 'MGL') & (results.points == 'GAS ANALYZER') & (results[8] == 'Due')])
    try:
        GASANA_MGL_late_percentage = "{:.0%}".format(GASANA_MGL_late/ GASANA_MGL_sched)
    except ZeroDivisionError:
        GASANA_MGL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    GASANA_PGL_sched = len(results[(results.line == 'PGL') & (results.points == 'GAS ANALYZER')])
    GASANA_PGL_comp = len(results[(results.line == 'PGL') & (results.points == 'GAS ANALYZER') & (results[8] == 'Yes')])
    GASANA_PGL_late = len(results[(results.line == 'PGL') & (results.points == 'GAS ANALYZER') & (results[8] == 'Due')])
    try:
        GASANA_PGL_late_percentage = "{:.0%}".format(GASANA_PGL_late/ GASANA_PGL_sched)
    except ZeroDivisionError:
        GASANA_PGL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    GASANA_EXPO_sched = len(results[(results.line == 'EXPO') & (results.points == 'GAS ANALYZER')])
    GASANA_EXPO_comp = len(results[(results.line == 'EXPO') & (results.points == 'GAS ANALYZER') & (results[8] == 'Yes')])
    GASANA_EXPO_late = len(results[(results.line == 'EXPO') & (results.points == 'GAS ANALYZER') & (results[8] == 'Due')])
    try:
        GASANA_EXPO_late_percentage = "{:.0%}".format(GASANA_EXPO_late/ GASANA_EXPO_sched)
    except ZeroDivisionError:
        GASANA_EXPO_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    GASANA_MOL_sched = len(results[(results.line == 'MOL') & (results.points == 'GAS ANALYZER')])
    GASANA_MOL_comp = len(results[(results.line == 'MOL') & (results.points == 'GAS ANALYZER') & (results[8] == 'Yes')])
    GASANA_MOL_late = len(results[(results.line == 'MOL') & (results.points == 'GAS ANALYZER') & (results[8] == 'Due')])
    try:
        GASANA_MOL_late_percentage = "{:.0%}".format(GASANA_MOL_late/ GASANA_MOL_sched)
    except ZeroDivisionError:
        GASANA_MOL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    GASANA_OTHER_sched = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results.points == 'GAS ANALYZER')])
    GASANA_OTHER_comp = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results[8] == 'Yes') & (results.points == 'GAS ANALYZER')])
    GASANA_OTHER_late = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results[8] == 'Due') & (results.points == 'GAS ANALYZER')])
    try:
        GASANA_OTHER_late_percentage = "{:.0%}".format(GASANA_OTHER_late/ GASANA_OTHER_sched)
    except ZeroDivisionError:
        GASANA_OTHER_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------



    RAD_MBL_sched = len(results[(results.line == 'MBL') & (results.points == 'RADIO')])
    RAD_MBL_comp = len(results[(results.line == 'MBL') & (results.points == 'RADIO') & (results[8] == 'Yes')])
    RAD_MBL_late = len(results[(results.line == 'MBL') & (results.points == 'RADIO') & (results[8] == 'Due')])
    try:
        RAD_MBL_late_percentage = "{:.0%}".format(RAD_MBL_late/ RAD_MBL_sched)
    except ZeroDivisionError:
        RAD_MBL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    RAD_MRL_MPL_sched = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'RADIO'))])
    RAD_MRL_MPL_comp = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'RADIO')& (results[8] == 'Yes'))])
    RAD_MRL_MPL_late = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'RADIO')& (results[8] == 'Due'))])
    try:
        RAD_MRL_MPL_late_percentage = "{:.0%}".format(RAD_MRL_MPL_late/ RAD_MRL_MPL_sched)
    except ZeroDivisionError:
        RAD_MRL_MPL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    RAD_MGL_sched = len(results[(results.line == 'MGL') & (results.points == 'RADIO')])
    RAD_MGL_comp = len(results[(results.line == 'MGL') & (results.points == 'RADIO') & (results[8] == 'Yes')])
    RAD_MGL_late = len(results[(results.line == 'MGL') & (results.points == 'RADIO') & (results[8] == 'Due')])
    try:
        RAD_MGL_late_percentage = "{:.0%}".format(RAD_MGL_late/ RAD_MGL_sched)
    except ZeroDivisionError:
        RAD_MGL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    RAD_PGL_sched = len(results[(results.line == 'PGL') & (results.points == 'RADIO')])
    RAD_PGL_comp = len(results[(results.line == 'PGL') & (results.points == 'RADIO') & (results[8] == 'Yes')])
    RAD_PGL_late = len(results[(results.line == 'PGL') & (results.points == 'RADIO') & (results[8] == 'Due')])
    try:
        RAD_PGL_late_percentage = "{:.0%}".format(RAD_PGL_late/ RAD_PGL_sched)
    except ZeroDivisionError:
        RAD_PGL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    RAD_EXPO_sched = len(results[(results.line == 'EXPO') & (results.points == 'RADIO')])
    RAD_EXPO_comp = len(results[(results.line == 'EXPO') & (results.points == 'RADIO') & (results[8] == 'Yes')])
    RAD_EXPO_late = len(results[(results.line == 'EXPO') & (results.points == 'RADIO') & (results[8] == 'Due')])
    try:
        RAD_EXPO_late_percentage = "{:.0%}".format(RAD_EXPO_late/ RAD_EXPO_sched)
    except ZeroDivisionError:
        RAD_EXPO_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    RAD_MOL_sched = len(results[(results.line == 'MOL') & (results.points == 'RADIO')])
    RAD_MOL_comp = len(results[(results.line == 'MOL') & (results.points == 'RADIO') & (results[8] == 'Yes')])
    RAD_MOL_late = len(results[(results.line == 'MOL') & (results.points == 'RADIO') & (results[8] == 'Due')])
    try:
        RAD_MOL_late_percentage = "{:.0%}".format(RAD_MOL_late/ RAD_MOL_sched)
    except ZeroDivisionError:
        RAD_MOL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    RAD_OTHER_sched = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results.points == 'RADIO')])
    RAD_OTHER_comp = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results[8] == 'Yes') & (results.points == 'RADIO')])
    RAD_OTHER_late = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results[8] == 'Due') & (results.points == 'RADIO')])
    try:
        RAD_OTHER_late_percentage = "{:.0%}".format(RAD_OTHER_late/ RAD_OTHER_sched)
    except ZeroDivisionError:
        RAD_OTHER_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------










    TRAINCC_MBL_sched = len(results[(results.line == 'MBL') & (results.points == 'TRAIN CCTV')])
    TRAINCC_MBL_comp = len(results[(results.line == 'MBL') & (results.points == 'TRAIN CCTV') & (results[8] == 'Yes')])
    TRAINCC_MBL_late = len(results[(results.line == 'MBL') & (results.points == 'TRAIN CCTV') & (results[8] == 'Due')])
    try:
        TRAINCC_MBL_late_percentage = "{:.0%}".format(TRAINCC_MBL_late/ TRAINCC_MBL_sched)
    except ZeroDivisionError:
        TRAINCC_MBL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    TRAINCC_MRL_MPL_sched = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'TRAIN CCTV'))])
    TRAINCC_MRL_MPL_comp = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'TRAIN CCTV')& (results[8] == 'Yes'))])
    TRAINCC_MRL_MPL_late = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'TRAIN CCTV')& (results[8] == 'Due'))])
    try:
        TRAINCC_MRL_MPL_late_percentage = "{:.0%}".format(TRAINCC_MRL_MPL_late/ TRAINCC_MRL_MPL_sched)
    except ZeroDivisionError:
        TRAINCC_MRL_MPL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    TRAINCC_MGL_sched = len(results[(results.line == 'MGL') & (results.points == 'TRAIN CCTV')])
    TRAINCC_MGL_comp = len(results[(results.line == 'MGL') & (results.points == 'TRAIN CCTV') & (results[8] == 'Yes')])
    TRAINCC_MGL_late = len(results[(results.line == 'MGL') & (results.points == 'TRAIN CCTV') & (results[8] == 'Due')])
    try:
        TRAINCC_MGL_late_percentage = "{:.0%}".format(TRAINCC_MGL_late/ TRAINCC_MGL_sched)
    except ZeroDivisionError:
        TRAINCC_MGL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    TRAINCC_PGL_sched = len(results[(results.line == 'PGL') & (results.points == 'TRAIN CCTV')])
    TRAINCC_PGL_comp = len(results[(results.line == 'PGL') & (results.points == 'TRAIN CCTV') & (results[8] == 'Yes')])
    TRAINCC_PGL_late = len(results[(results.line == 'PGL') & (results.points == 'TRAIN CCTV') & (results[8] == 'Due')])
    try:
        TRAINCC_PGL_late_percentage = "{:.0%}".format(TRAINCC_PGL_late/ TRAINCC_PGL_sched)
    except ZeroDivisionError:
        TRAINCC_PGL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    TRAINCC_EXPO_sched = len(results[(results.line == 'EXPO') & (results.points == 'TRAIN CCTV')])
    TRAINCC_EXPO_comp = len(results[(results.line == 'EXPO') & (results.points == 'TRAIN CCTV') & (results[8] == 'Yes')])
    TRAINCC_EXPO_late = len(results[(results.line == 'EXPO') & (results.points == 'TRAIN CCTV') & (results[8] == 'Due')])
    try:
        TRAINCC_EXPO_late_percentage = "{:.0%}".format(TRAINCC_EXPO_late/ TRAINCC_EXPO_sched)
    except ZeroDivisionError:
        TRAINCC_EXPO_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    TRAINCC_MOL_sched = len(results[(results.line == 'MOL') & (results.points == 'TRAIN CCTV')])
    TRAINCC_MOL_comp = len(results[(results.line == 'MOL') & (results.points == 'TRAIN CCTV') & (results[8] == 'Yes')])
    TRAINCC_MOL_late = len(results[(results.line == 'MOL') & (results.points == 'TRAIN CCTV') & (results[8] == 'Due')])
    try:
        TRAINCC_MOL_late_percentage = "{:.0%}".format(TRAINCC_MOL_late/ TRAINCC_MOL_sched)
    except ZeroDivisionError:
        TRAINCC_MOL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    TRAINCC_OTHER_sched = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results.points == 'TRAIN CCTV')])
    TRAINCC_OTHER_comp = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results[8] == 'Yes') & (results.points == 'TRAIN CCTV')])
    TRAINCC_OTHER_late = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results[8] == 'Due') & (results.points == 'TRAIN CCTV')])
    try:
        TRAINCC_OTHER_late_percentage = "{:.0%}".format(TRAINCC_OTHER_late/ TRAINCC_OTHER_sched)
    except ZeroDivisionError:
        TRAINCC_OTHER_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------




    SEIS_MBL_sched = len(results[(results.line == 'MBL') & (results.points == 'SEISMIC')])
    SEIS_MBL_comp = len(results[(results.line == 'MBL') & (results.points == 'SEISMIC') & (results[8] == 'Yes')])
    SEIS_MBL_late = len(results[(results.line == 'MBL') & (results.points == 'SEISMIC') & (results[8] == 'Due')])
    try:
        SEIS_MBL_late_percentage = "{:.0%}".format(SEIS_MBL_late/ SEIS_MBL_sched)
    except ZeroDivisionError:
        SEIS_MBL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    SEIS_MRL_MPL_sched = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'SEISMIC'))])
    SEIS_MRL_MPL_comp = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'SEISMIC')& (results[8] == 'Yes'))])
    SEIS_MRL_MPL_late = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'SEISMIC')& (results[8] == 'Due'))])
    try:
        SEIS_MRL_MPL_late_percentage = "{:.0%}".format(SEIS_MRL_MPL_late/ SEIS_MRL_MPL_sched)
    except ZeroDivisionError:
        SEIS_MRL_MPL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    SEIS_MGL_sched = len(results[(results.line == 'MGL') & (results.points == 'SEISMIC')])
    SEIS_MGL_comp = len(results[(results.line == 'MGL') & (results.points == 'SEISMIC') & (results[8] == 'Yes')])
    SEIS_MGL_late = len(results[(results.line == 'MGL') & (results.points == 'SEISMIC') & (results[8] == 'Due')])
    try:
        SEIS_MGL_late_percentage = "{:.0%}".format(SEIS_MGL_late/ SEIS_MGL_sched)
    except ZeroDivisionError:
        SEIS_MGL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    SEIS_PGL_sched = len(results[(results.line == 'PGL') & (results.points == 'SEISMIC')])
    SEIS_PGL_comp = len(results[(results.line == 'PGL') & (results.points == 'SEISMIC') & (results[8] == 'Yes')])
    SEIS_PGL_late = len(results[(results.line == 'PGL') & (results.points == 'SEISMIC') & (results[8] == 'Due')])
    try:
        SEIS_PGL_late_percentage = "{:.0%}".format(SEIS_PGL_late/ SEIS_PGL_sched)
    except ZeroDivisionError:
        SEIS_PGL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    SEIS_EXPO_sched = len(results[(results.line == 'EXPO') & (results.points == 'SEISMIC')])
    SEIS_EXPO_comp = len(results[(results.line == 'EXPO') & (results.points == 'SEISMIC') & (results[8] == 'Yes')])
    SEIS_EXPO_late = len(results[(results.line == 'EXPO') & (results.points == 'SEISMIC') & (results[8] == 'Due')])
    try:
        SEIS_EXPO_late_percentage = "{:.0%}".format(SEIS_EXPO_late/ SEIS_EXPO_sched)
    except ZeroDivisionError:
        SEIS_EXPO_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    SEIS_MOL_sched = len(results[(results.line == 'MOL') & (results.points == 'SEISMIC')])
    SEIS_MOL_comp = len(results[(results.line == 'MOL') & (results.points == 'SEISMIC') & (results[8] == 'Yes')])
    SEIS_MOL_late = len(results[(results.line == 'MOL') & (results.points == 'SEISMIC') & (results[8] == 'Due')])
    try:
        SEIS_MOL_late_percentage = "{:.0%}".format(SEIS_MOL_late/ SEIS_MOL_sched)
    except ZeroDivisionError:
        SEIS_MOL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    SEIS_OTHER_sched = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results.points == 'SEISMIC')])
    SEIS_OTHER_comp = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results[8] == 'Yes') & (results.points == 'SEISMIC')])
    SEIS_OTHER_late = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results[8] == 'Due') & (results.points == 'SEISMIC')])
    try:
        SEIS_OTHER_late_percentage = "{:.0%}".format(SEIS_OTHER_late/ SEIS_OTHER_sched)
    except ZeroDivisionError:
        SEIS_OTHER_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------

    TRAINPRO_MBL_sched = len(results[(results.line == 'MBL') & (results.points == 'TRAIN PROTRAN')])
    TRAINPRO_MBL_comp = len(results[(results.line == 'MBL') & (results.points == 'TRAIN PROTRAN') & (results[8] == 'Yes')])
    TRAINPRO_MBL_late = len(results[(results.line == 'MBL') & (results.points == 'TRAIN PROTRAN') & (results[8] == 'Due')])
    try:
        TRAINPRO_MBL_late_percentage = "{:.0%}".format(TRAINPRO_MBL_late/ TRAINPRO_MBL_sched)
    except ZeroDivisionError:
        TRAINPRO_MBL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    TRAINPRO_MRL_MPL_sched = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'TRAIN PROTRAN'))])
    TRAINPRO_MRL_MPL_comp = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'TRAIN PROTRAN')& (results[8] == 'Yes'))])
    TRAINPRO_MRL_MPL_late = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'TRAIN PROTRAN')& (results[8] == 'Due'))])
    try:
        TRAINPRO_MRL_MPL_late_percentage = "{:.0%}".format(TRAINPRO_MRL_MPL_late/ TRAINPRO_MRL_MPL_sched)
    except ZeroDivisionError:
        TRAINPRO_MRL_MPL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    TRAINPRO_MGL_sched = len(results[(results.line == 'MGL') & (results.points == 'TRAIN PROTRAN')])
    TRAINPRO_MGL_comp = len(results[(results.line == 'MGL') & (results.points == 'TRAIN PROTRAN') & (results[8] == 'Yes')])
    TRAINPRO_MGL_late = len(results[(results.line == 'MGL') & (results.points == 'TRAIN PROTRAN') & (results[8] == 'Due')])
    try:
        TRAINPRO_MGL_late_percentage = "{:.0%}".format(TRAINPRO_MGL_late/ TRAINPRO_MGL_sched)
    except ZeroDivisionError:
        TRAINPRO_MGL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    TRAINPRO_PGL_sched = len(results[(results.line == 'PGL') & (results.points == 'TRAIN PROTRAN')])
    TRAINPRO_PGL_comp = len(results[(results.line == 'PGL') & (results.points == 'TRAIN PROTRAN') & (results[8] == 'Yes')])
    TRAINPRO_PGL_late = len(results[(results.line == 'PGL') & (results.points == 'TRAIN PROTRAN') & (results[8] == 'Due')])
    try:
        TRAINPRO_PGL_late_percentage = "{:.0%}".format(TRAINPRO_PGL_late/ TRAINPRO_PGL_sched)
    except ZeroDivisionError:
        TRAINPRO_PGL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    TRAINPRO_EXPO_sched = len(results[(results.line == 'EXPO') & (results.points == 'TRAIN PROTRAN')])
    TRAINPRO_EXPO_comp = len(results[(results.line == 'EXPO') & (results.points == 'TRAIN PROTRAN') & (results[8] == 'Yes')])
    TRAINPRO_EXPO_late = len(results[(results.line == 'EXPO') & (results.points == 'TRAIN PROTRAN') & (results[8] == 'Due')])
    try:
        TRAINPRO_EXPO_late_percentage = "{:.0%}".format(TRAINPRO_EXPO_late/ TRAINPRO_EXPO_sched)
    except ZeroDivisionError:
        TRAINPRO_EXPO_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    TRAINPRO_MOL_sched = len(results[(results.line == 'MOL') & (results.points == 'TRAIN PROTRAN')])
    TRAINPRO_MOL_comp = len(results[(results.line == 'MOL') & (results.points == 'TRAIN PROTRAN') & (results[8] == 'Yes')])
    TRAINPRO_MOL_late = len(results[(results.line == 'MOL') & (results.points == 'TRAIN PROTRAN') & (results[8] == 'Due')])
    try:
        TRAINPRO_MOL_late_percentage = "{:.0%}".format(TRAINPRO_MOL_late/ TRAINPRO_MOL_sched)
    except ZeroDivisionError:
        TRAINPRO_MOL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    TRAINPRO_OTHER_sched = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results.points == 'TRAIN PROTRAN')])
    TRAINPRO_OTHER_comp = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results[8] == 'Yes') & (results.points == 'TRAIN PROTRAN')])
    TRAINPRO_OTHER_late = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results[8] == 'Due') & (results.points == 'TRAIN PROTRAN')])
    try:
        TRAINPRO_OTHER_late_percentage = "{:.0%}".format(TRAINPRO_OTHER_late/ TRAINPRO_OTHER_sched)
    except ZeroDivisionError:
        TRAINPRO_OTHER_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------


    TRAINRAD_MBL_sched = len(results[(results.line == 'MBL') & (results.points == 'TRAIN RADIO')])
    TRAINRAD_MBL_comp = len(results[(results.line == 'MBL') & (results.points == 'TRAIN RADIO') & (results[8] == 'Yes')])
    TRAINRAD_MBL_late = len(results[(results.line == 'MBL') & (results.points == 'TRAIN RADIO') & (results[8] == 'Due')])
    try:
        TRAINRAD_MBL_late_percentage = "{:.0%}".format(TRAINRAD_MBL_late/ TRAINRAD_MBL_sched)
    except ZeroDivisionError:
        TRAINRAD_MBL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    TRAINRAD_MRL_MPL_sched = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'TRAIN RADIO'))])
    TRAINRAD_MRL_MPL_comp = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'TRAIN RADIO')& (results[8] == 'Yes'))])
    TRAINRAD_MRL_MPL_late = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'TRAIN RADIO')& (results[8] == 'Due'))])
    try:
        TRAINRAD_MRL_MPL_late_percentage = "{:.0%}".format(TRAINRAD_MRL_MPL_late/ TRAINRAD_MRL_MPL_sched)
    except ZeroDivisionError:
        TRAINRAD_MRL_MPL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    TRAINRAD_MGL_sched = len(results[(results.line == 'MGL') & (results.points == 'TRAIN RADIO')])
    TRAINRAD_MGL_comp = len(results[(results.line == 'MGL') & (results.points == 'TRAIN RADIO') & (results[8] == 'Yes')])
    TRAINRAD_MGL_late = len(results[(results.line == 'MGL') & (results.points == 'TRAIN RADIO') & (results[8] == 'Due')])
    try:
        TRAINRAD_MGL_late_percentage = "{:.0%}".format(TRAINRAD_MGL_late/ TRAINRAD_MGL_sched)
    except ZeroDivisionError:
        TRAINRAD_MGL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    TRAINRAD_PGL_sched = len(results[(results.line == 'PGL') & (results.points == 'TRAIN RADIO')])
    TRAINRAD_PGL_comp = len(results[(results.line == 'PGL') & (results.points == 'TRAIN RADIO') & (results[8] == 'Yes')])
    TRAINRAD_PGL_late = len(results[(results.line == 'PGL') & (results.points == 'TRAIN RADIO') & (results[8] == 'Due')])
    try:
        TRAINRAD_PGL_late_percentage = "{:.0%}".format(TRAINRAD_PGL_late/ TRAINRAD_PGL_sched)
    except ZeroDivisionError:
        TRAINRAD_PGL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    TRAINRAD_EXPO_sched = len(results[(results.line == 'EXPO') & (results.points == 'TRAIN RADIO')])
    TRAINRAD_EXPO_comp = len(results[(results.line == 'EXPO') & (results.points == 'TRAIN RADIO') & (results[8] == 'Yes')])
    TRAINRAD_EXPO_late = len(results[(results.line == 'EXPO') & (results.points == 'TRAIN RADIO') & (results[8] == 'Due')])
    try:
        TRAINRAD_EXPO_late_percentage = "{:.0%}".format(TRAINRAD_EXPO_late/ TRAINRAD_EXPO_sched)
    except ZeroDivisionError:
        TRAINRAD_EXPO_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    TRAINRAD_MOL_sched = len(results[(results.line == 'MOL') & (results.points == 'TRAIN RADIO')])
    TRAINRAD_MOL_comp = len(results[(results.line == 'MOL') & (results.points == 'TRAIN RADIO') & (results[8] == 'Yes')])
    TRAINRAD_MOL_late = len(results[(results.line == 'MOL') & (results.points == 'TRAIN RADIO') & (results[8] == 'Due')])
    try:
        TRAINRAD_MOL_late_percentage = "{:.0%}".format(TRAINRAD_MOL_late/ TRAINRAD_MOL_sched)
    except ZeroDivisionError:
        TRAINRAD_MOL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    TRAINRAD_OTHER_sched = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results.points == 'TRAIN RADIO')])
    TRAINRAD_OTHER_comp = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results[8] == 'Yes') & (results.points == 'TRAIN RADIO')])
    TRAINRAD_OTHER_late = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results[8] == 'Due') & (results.points == 'TRAIN RADIO')])
    try:
        TRAINRAD_OTHER_late_percentage = "{:.0%}".format(TRAINRAD_OTHER_late/ TRAINRAD_OTHER_sched)
    except ZeroDivisionError:
        TRAINRAD_OTHER_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------



    IDS_MBL_sched = len(results[(results.line == 'MBL') & (results.points == 'IDS')])
    IDS_MBL_comp = len(results[(results.line == 'MBL') & (results.points == 'IDS') & (results[8] == 'Yes')])
    IDS_MBL_late = len(results[(results.line == 'MBL') & (results.points == 'IDS') & (results[8] == 'Due')])
    try:
        IDS_MBL_late_percentage = "{:.0%}".format(IDS_MBL_late/ IDS_MBL_sched)
    except ZeroDivisionError:
        IDS_MBL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    IDS_MRL_MPL_sched = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'IDS'))])
    IDS_MRL_MPL_comp = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'IDS')& (results[8] == 'Yes'))])
    IDS_MRL_MPL_late = len(results[(results.line.str.contains('MRL|MPL') & (results.points == 'IDS')& (results[8] == 'Due'))])
    try:
        IDS_MRL_MPL_late_percentage = "{:.0%}".format(IDS_MRL_MPL_late/ IDS_MRL_MPL_sched)
    except ZeroDivisionError:
        IDS_MRL_MPL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    IDS_MGL_sched = len(results[(results.line == 'MGL') & (results.points == 'IDS')])
    IDS_MGL_comp = len(results[(results.line == 'MGL') & (results.points == 'IDS') & (results[8] == 'Yes')])
    IDS_MGL_late = len(results[(results.line == 'MGL') & (results.points == 'IDS') & (results[8] == 'Due')])
    try:
        IDS_MGL_late_percentage = "{:.0%}".format(IDS_MGL_late/ IDS_MGL_sched)
    except ZeroDivisionError:
        IDS_MGL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    IDS_PGL_sched = len(results[(results.line == 'PGL') & (results.points == 'IDS')])
    IDS_PGL_comp = len(results[(results.line == 'PGL') & (results.points == 'IDS') & (results[8] == 'Yes')])
    IDS_PGL_late = len(results[(results.line == 'PGL') & (results.points == 'IDS') & (results[8] == 'Due')])
    try:
        IDS_PGL_late_percentage = "{:.0%}".format(IDS_PGL_late/ IDS_PGL_sched)
    except ZeroDivisionError:
        IDS_PGL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    IDS_EXPO_sched = len(results[(results.line == 'EXPO') & (results.points == 'IDS')])
    IDS_EXPO_comp = len(results[(results.line == 'EXPO') & (results.points == 'IDS') & (results[8] == 'Yes')])
    IDS_EXPO_late = len(results[(results.line == 'EXPO') & (results.points == 'IDS') & (results[8] == 'Due')])
    try:
        IDS_EXPO_late_percentage = "{:.0%}".format(IDS_EXPO_late/ IDS_EXPO_sched)
    except ZeroDivisionError:
        IDS_EXPO_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    IDS_MOL_sched = len(results[(results.line == 'MOL') & (results.points == 'IDS')])
    IDS_MOL_comp = len(results[(results.line == 'MOL') & (results.points == 'IDS') & (results[8] == 'Yes')])
    IDS_MOL_late = len(results[(results.line == 'MOL') & (results.points == 'IDS') & (results[8] == 'Due')])
    try:
        IDS_MOL_late_percentage = "{:.0%}".format(IDS_MOL_late/ IDS_MOL_sched)
    except ZeroDivisionError:
        IDS_MOL_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    IDS_OTHER_sched = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results.points == 'IDS')])
    IDS_OTHER_comp = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results[8] == 'Yes') & (results.points == 'IDS')])
    IDS_OTHER_late = len(results[~results.line.str.contains('EXPO|MOL|MRL|MPL|MGL|PGL|CRENSHAW|MBL', na=False) & (results[8] == 'Due') & (results.points == 'IDS')])
    try:
        IDS_OTHER_late_percentage = "{:.0%}".format(IDS_OTHER_late/ IDS_OTHER_sched)
    except ZeroDivisionError:
        IDS_OTHER_late_percentage = "{:.0%}".format(0)
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------
    #-------------------------------------------------------------------------------------------------











    context = {'MGL_total': MGL_total,
            'MGL_completed':MGL_completed,
            'MGL_late':MGL_late,
            'MGL_late_percentage':MGL_late_percentage,
            'MGL_complete_percentage':MGL_complete_percentage,
            #--------------------------------------------------
            'MBL_complete_percentage':MBL_complete_percentage,
            'MBL_total': MBL_total,
            'MBL_completed':MBL_completed,
            'MBL_late':MBL_late,
            'MBL_late_percentage':MBL_late_percentage,
            #-------------------------------------------------
            'EXPO_total': EXPO_total,
            'EXPO_completed':EXPO_completed,
            'EXPO_late':EXPO_late,
            'EXPO_late_percentage':EXPO_late_percentage,
            'EXPO_complete_percentage':EXPO_complete_percentage,
            #----------------------------------------------------
            'MRL_MPL_total': MRL_MPL_total,
            'MRL_MPL_completed':MRL_MPL_completed,
            'MRL_MPL_late':MRL_MPL_late,
            'MRL_MPL_late_percentage':MRL_MPL_late_percentage,
            'MRL_MPL_complete_percentage':MRL_MPL_complete_percentage,
            #---------------------------------------------------------------
            'PGL_total': PGL_total,
            'PGL_completed':PGL_completed,
            'PGL_late':PGL_late,
            'PGL_late_percentage':PGL_late_percentage,
            'PGL_complete_percentage':PGL_complete_percentage,
            #---------------------------------------------------------------
            'CRENSHAW_total': CRENSHAW_total,
            'CRENSHAW_completed':CRENSHAW_completed,
            'CRENSHAW_late':CRENSHAW_late,
            'CRENSHAW_late_percentage':CRENSHAW_late_percentage,
            'CRENSHAW_complete_percentage':CRENSHAW_complete_percentage,
            #---------------------------------------------------------------
            'MOL_total':MOL_total,
            'MOL_completed':MOL_completed,
            'MOL_late':MOL_late,
            'MOL_late_percentage':MOL_late_percentage,
            'MOL_complete_percentage':MOL_complete_percentage,
            #---------------------------------------------------------------
            'OTHER_total':OTHER_total,
            'OTHER_completed':OTHER_completed,
            'OTHER_late':OTHER_late,
            'OTHER_late_percentage':OTHER_late_percentage,
            'OTHER_complete_percentage':OTHER_complete_percentage,
            #------------------------------------------------------------------------------------------------
            #-------------------------------------------------------------------------------------------------
            #-------------------------------------------------------------------------------------------------
            'CCTV_MBL_sched':CCTV_MBL_sched,
            'CCTV_MBL_late':CCTV_MBL_late,
            'CCTV_MBL_comp':CCTV_MBL_comp,
            'CCTV_MBL_late_percentage':CCTV_MBL_late_percentage,
            #---------------------------------------------------------------
            'CCTV_MRL_MPL_sched':CCTV_MRL_MPL_sched,
            'CCTV_MRL_MPL_late':CCTV_MRL_MPL_late,
            'CCTV_MRL_MPL_comp':CCTV_MRL_MPL_comp,
            'CCTV_MRL_MPL_late_percentage':CCTV_MRL_MPL_late_percentage,
            #---------------------------------------------------------------
            'CCTV_MGL_sched':CCTV_MGL_sched,
            'CCTV_MGL_late':CCTV_MGL_late,
            'CCTV_MGL_comp':CCTV_MGL_comp,
            'CCTV_MGL_late_percentage':CCTV_MGL_late_percentage,
            #---------------------------------------------------------------
            'CCTV_PGL_sched':CCTV_PGL_sched,
            'CCTV_PGL_late':CCTV_PGL_late,
            'CCTV_PGL_comp':CCTV_PGL_comp,
            'CCTV_PGL_late_percentage':CCTV_PGL_late_percentage,
            #---------------------------------------------------------------
            'CCTV_EXPO_sched':CCTV_EXPO_sched,
            'CCTV_EXPO_late':CCTV_EXPO_late,
            'CCTV_EXPO_comp':CCTV_EXPO_comp,
            'CCTV_EXPO_late_percentage':CCTV_EXPO_late_percentage,
            #---------------------------------------------------------------
            'CCTV_MOL_sched':CCTV_MOL_sched,
            'CCTV_MOL_late':CCTV_MOL_late,
            'CCTV_MOL_comp':CCTV_MOL_comp,
            'CCTV_MOL_late_percentage':CCTV_MOL_late_percentage,
            #---------------------------------------------------------------
            'CCTV_OTHER_sched':CCTV_OTHER_sched,
            'CCTV_OTHER_late':CCTV_OTHER_late,
            'CCTV_OTHER_comp':CCTV_OTHER_comp,
            'CCTV_OTHER_late_percentage':CCTV_OTHER_late_percentage,
            #------------------------------------------------------------------------------------------------
            #-------------------------------------------------------------------------------------------------
            #-------------------------------------------------------------------------------------------------
            'CTS_MBL_sched':CTS_MBL_sched,
            'CTS_MBL_late':CTS_MBL_late,
            'CTS_MBL_comp':CTS_MBL_comp,
            'CTS_MBL_late_percentage':CTS_MBL_late_percentage,
            #---------------------------------------------------------------
            'CTS_MRL_MPL_sched':CTS_MRL_MPL_sched,
            'CTS_MRL_MPL_late':CTS_MRL_MPL_late,
            'CTS_MRL_MPL_comp':CTS_MRL_MPL_comp,
            'CTS_MRL_MPL_late_percentage':CTS_MRL_MPL_late_percentage,
            #---------------------------------------------------------------
            'CTS_MGL_sched':CTS_MGL_sched,
            'CTS_MGL_late':CTS_MGL_late,
            'CTS_MGL_comp':CTS_MGL_comp,
            'CTS_MGL_late_percentage':CTS_MGL_late_percentage,
            #---------------------------------------------------------------
            'CTS_PGL_sched':CTS_PGL_sched,
            'CTS_PGL_late':CTS_PGL_late,
            'CTS_PGL_comp':CTS_PGL_comp,
            'CTS_PGL_late_percentage':CTS_PGL_late_percentage,
            #---------------------------------------------------------------
            'CTS_EXPO_sched':CTS_EXPO_sched,
            'CTS_EXPO_late':CTS_EXPO_late,
            'CTS_EXPO_comp':CTS_EXPO_comp,
            'CTS_EXPO_late_percentage':CTS_EXPO_late_percentage,
            #---------------------------------------------------------------
            'CTS_MOL_sched':CTS_MOL_sched,
            'CTS_MOL_late':CTS_MOL_late,
            'CTS_MOL_comp':CTS_MOL_comp,
            'CTS_MOL_late_percentage':CTS_MOL_late_percentage,
            #---------------------------------------------------------------
            'CTS_OTHER_sched':CTS_OTHER_sched,
            'CTS_OTHER_late':CTS_OTHER_late,
            'CTS_OTHER_comp':CTS_OTHER_comp,
            'CTS_OTHER_late_percentage':CTS_OTHER_late_percentage,
            #------------------------------------------------------------------------------------------------
            #-------------------------------------------------------------------------------------------------
            #-------------------------------------------------------------------------------------------------
            'TEL_MBL_sched':TEL_MBL_sched,
            'TEL_MBL_late':TEL_MBL_late,
            'TEL_MBL_comp':TEL_MBL_comp,
            'TEL_MBL_late_percentage':TEL_MBL_late_percentage,
            #---------------------------------------------------------------
            'TEL_MRL_MPL_sched':TEL_MRL_MPL_sched,
            'TEL_MRL_MPL_late':TEL_MRL_MPL_late,
            'TEL_MRL_MPL_comp':TEL_MRL_MPL_comp,
            'TEL_MRL_MPL_late_percentage':TEL_MRL_MPL_late_percentage,
            #---------------------------------------------------------------
            'TEL_MGL_sched':TEL_MGL_sched,
            'TEL_MGL_late':TEL_MGL_late,
            'TEL_MGL_comp':TEL_MGL_comp,
            'TEL_MGL_late_percentage':TEL_MGL_late_percentage,
            #---------------------------------------------------------------
            'TEL_PGL_sched':TEL_PGL_sched,
            'TEL_PGL_late':TEL_PGL_late,
            'TEL_PGL_comp':TEL_PGL_comp,
            'TEL_PGL_late_percentage':TEL_PGL_late_percentage,
            #---------------------------------------------------------------
            'TEL_EXPO_sched':TEL_EXPO_sched,
            'TEL_EXPO_late':TEL_EXPO_late,
            'TEL_EXPO_comp':TEL_EXPO_comp,
            'TEL_EXPO_late_percentage':TEL_EXPO_late_percentage,
            #---------------------------------------------------------------
            'TEL_MOL_sched':TEL_MOL_sched,
            'TEL_MOL_late':TEL_MOL_late,
            'TEL_MOL_comp':TEL_MOL_comp,
            'TEL_MOL_late_percentage':TEL_MOL_late_percentage,
            #---------------------------------------------------------------
            'TEL_OTHER_sched':TEL_OTHER_sched,
            'TEL_OTHER_late':TEL_OTHER_late,
            'TEL_OTHER_comp':TEL_OTHER_comp,
            'TEL_OTHER_late_percentage':TEL_OTHER_late_percentage,
            #------------------------------------------------------------------------------------------------
            #-------------------------------------------------------------------------------------------------
            #-------------------------------------------------------------------------------------------------
            'PA_MBL_sched':PA_MBL_sched,
            'PA_MBL_late':PA_MBL_late,
            'PA_MBL_comp':PA_MBL_comp,
            'PA_MBL_late_percentage':PA_MBL_late_percentage,
            #---------------------------------------------------------------
            'PA_MRL_MPL_sched':PA_MRL_MPL_sched,
            'PA_MRL_MPL_late':PA_MRL_MPL_late,
            'PA_MRL_MPL_comp':PA_MRL_MPL_comp,
            'PA_MRL_MPL_late_percentage':PA_MRL_MPL_late_percentage,
            #---------------------------------------------------------------
            'PA_MGL_sched':PA_MGL_sched,
            'PA_MGL_late':PA_MGL_late,
            'PA_MGL_comp':PA_MGL_comp,
            'PA_MGL_late_percentage':PA_MGL_late_percentage,
            #---------------------------------------------------------------
            'PA_PGL_sched':PA_PGL_sched,
            'PA_PGL_late':PA_PGL_late,
            'PA_PGL_comp':PA_PGL_comp,
            'PA_PGL_late_percentage':PA_PGL_late_percentage,
            #---------------------------------------------------------------
            'PA_EXPO_sched':PA_EXPO_sched,
            'PA_EXPO_late':PA_EXPO_late,
            'PA_EXPO_comp':PA_EXPO_comp,
            'PA_EXPO_late_percentage':PA_EXPO_late_percentage,
            #---------------------------------------------------------------
            'PA_MOL_sched':PA_MOL_sched,
            'PA_MOL_late':PA_MOL_late,
            'PA_MOL_comp':PA_MOL_comp,
            'PA_MOL_late_percentage':PA_MOL_late_percentage,
            #---------------------------------------------------------------
            'PA_OTHER_sched':PA_OTHER_sched,
            'PA_OTHER_late':PA_OTHER_late,
            'PA_OTHER_comp':PA_OTHER_comp,
            'PA_OTHER_late_percentage':PA_OTHER_late_percentage,
            #------------------------------------------------------------------------------------------------
            #-------------------------------------------------------------------------------------------------
            #-------------------------------------------------------------------------------------------------
            'VMS_MBL_sched':VMS_MBL_sched,
            'VMS_MBL_late':VMS_MBL_late,
            'VMS_MBL_comp':VMS_MBL_comp,
            'VMS_MBL_late_percentage':VMS_MBL_late_percentage,
            #---------------------------------------------------------------
            'VMS_MRL_MPL_sched':VMS_MRL_MPL_sched,
            'VMS_MRL_MPL_late':VMS_MRL_MPL_late,
            'VMS_MRL_MPL_comp':VMS_MRL_MPL_comp,
            'VMS_MRL_MPL_late_percentage':VMS_MRL_MPL_late_percentage,
            #---------------------------------------------------------------
            'VMS_MGL_sched':VMS_MGL_sched,
            'VMS_MGL_late':VMS_MGL_late,
            'VMS_MGL_comp':VMS_MGL_comp,
            'VMS_MGL_late_percentage':VMS_MGL_late_percentage,
            #---------------------------------------------------------------
            'VMS_PGL_sched':VMS_PGL_sched,
            'VMS_PGL_late':VMS_PGL_late,
            'VMS_PGL_comp':VMS_PGL_comp,
            'VMS_PGL_late_percentage':VMS_PGL_late_percentage,
            #---------------------------------------------------------------
            'VMS_EXPO_sched':VMS_EXPO_sched,
            'VMS_EXPO_late':VMS_EXPO_late,
            'VMS_EXPO_comp':VMS_EXPO_comp,
            'VMS_EXPO_late_percentage':VMS_EXPO_late_percentage,
            #---------------------------------------------------------------
            'VMS_MOL_sched':VMS_MOL_sched,
            'VMS_MOL_late':VMS_MOL_late,
            'VMS_MOL_comp':VMS_MOL_comp,
            'VMS_MOL_late_percentage':VMS_MOL_late_percentage,
            #---------------------------------------------------------------
            'VMS_OTHER_sched':VMS_OTHER_sched,
            'VMS_OTHER_late':VMS_OTHER_late,
            'VMS_OTHER_comp':VMS_OTHER_comp,
            'VMS_OTHER_late_percentage':VMS_OTHER_late_percentage,
            #------------------------------------------------------------------------------------------------
            #-------------------------------------------------------------------------------------------------
            #-------------------------------------------------------------------------------------------------
            'RTU_MBL_sched':RTU_MBL_sched,
            'RTU_MBL_late':RTU_MBL_late,
            'RTU_MBL_comp':RTU_MBL_comp,
            'RTU_MBL_late_percentage':RTU_MBL_late_percentage,
            #---------------------------------------------------------------
            'RTU_MRL_MPL_sched':RTU_MRL_MPL_sched,
            'RTU_MRL_MPL_late':RTU_MRL_MPL_late,
            'RTU_MRL_MPL_comp':RTU_MRL_MPL_comp,
            'RTU_MRL_MPL_late_percentage':RTU_MRL_MPL_late_percentage,
            #---------------------------------------------------------------
            'RTU_MGL_sched':RTU_MGL_sched,
            'RTU_MGL_late':RTU_MGL_late,
            'RTU_MGL_comp':RTU_MGL_comp,
            'RTU_MGL_late_percentage':RTU_MGL_late_percentage,
            #---------------------------------------------------------------
            'RTU_PGL_sched':RTU_PGL_sched,
            'RTU_PGL_late':RTU_PGL_late,
            'RTU_PGL_comp':RTU_PGL_comp,
            'RTU_PGL_late_percentage':RTU_PGL_late_percentage,
            #---------------------------------------------------------------
            'RTU_EXPO_sched':RTU_EXPO_sched,
            'RTU_EXPO_late':RTU_EXPO_late,
            'RTU_EXPO_comp':RTU_EXPO_comp,
            'RTU_EXPO_late_percentage':RTU_EXPO_late_percentage,
            #---------------------------------------------------------------
            'RTU_MOL_sched':RTU_MOL_sched,
            'RTU_MOL_late':RTU_MOL_late,
            'RTU_MOL_comp':RTU_MOL_comp,
            'RTU_MOL_late_percentage':RTU_MOL_late_percentage,
            #---------------------------------------------------------------
            'RTU_OTHER_sched':RTU_OTHER_sched,
            'RTU_OTHER_late':RTU_OTHER_late,
            'RTU_OTHER_comp':RTU_OTHER_comp,
            'RTU_OTHER_late_percentage':RTU_OTHER_late_percentage,
            #------------------------------------------------------------------------------------------------
            #-------------------------------------------------------------------------------------------------
            #-------------------------------------------------------------------------------------------------
            'PLC_MBL_sched':PLC_MBL_sched,
            'PLC_MBL_late':PLC_MBL_late,
            'PLC_MBL_comp':PLC_MBL_comp,
            'PLC_MBL_late_percentage':PLC_MBL_late_percentage,
            #---------------------------------------------------------------
            'PLC_MRL_MPL_sched':PLC_MRL_MPL_sched,
            'PLC_MRL_MPL_late':PLC_MRL_MPL_late,
            'PLC_MRL_MPL_comp':PLC_MRL_MPL_comp,
            'PLC_MRL_MPL_late_percentage':PLC_MRL_MPL_late_percentage,
            #---------------------------------------------------------------
            'PLC_MGL_sched':PLC_MGL_sched,
            'PLC_MGL_late':PLC_MGL_late,
            'PLC_MGL_comp':PLC_MGL_comp,
            'PLC_MGL_late_percentage':PLC_MGL_late_percentage,
            #---------------------------------------------------------------
            'PLC_PGL_sched':PLC_PGL_sched,
            'PLC_PGL_late':PLC_PGL_late,
            'PLC_PGL_comp':PLC_PGL_comp,
            'PLC_PGL_late_percentage':PLC_PGL_late_percentage,
            #---------------------------------------------------------------
            'PLC_EXPO_sched':PLC_EXPO_sched,
            'PLC_EXPO_late':PLC_EXPO_late,
            'PLC_EXPO_comp':PLC_EXPO_comp,
            'PLC_EXPO_late_percentage':PLC_EXPO_late_percentage,
            #---------------------------------------------------------------
            'PLC_MOL_sched':PLC_MOL_sched,
            'PLC_MOL_late':PLC_MOL_late,
            'PLC_MOL_comp':PLC_MOL_comp,
            'PLC_MOL_late_percentage':PLC_MOL_late_percentage,
            #---------------------------------------------------------------
            'PLC_OTHER_sched':PLC_OTHER_sched,
            'PLC_OTHER_late':PLC_OTHER_late,
            'PLC_OTHER_comp':PLC_OTHER_comp,
            'PLC_OTHER_late_percentage':PLC_OTHER_late_percentage,
            #------------------------------------------------------------------------------------------------
            #-------------------------------------------------------------------------------------------------
            #-------------------------------------------------------------------------------------------------
            'FCS_MBL_sched':FCS_MBL_sched,
            'FCS_MBL_late':FCS_MBL_late,
            'FCS_MBL_comp':FCS_MBL_comp,
            'FCS_MBL_late_percentage':FCS_MBL_late_percentage,
            #---------------------------------------------------------------
            'FCS_MRL_MPL_sched':FCS_MRL_MPL_sched,
            'FCS_MRL_MPL_late':FCS_MRL_MPL_late,
            'FCS_MRL_MPL_comp':FCS_MRL_MPL_comp,
            'FCS_MRL_MPL_late_percentage':FCS_MRL_MPL_late_percentage,
            #---------------------------------------------------------------
            'FCS_MGL_sched':FCS_MGL_sched,
            'FCS_MGL_late':FCS_MGL_late,
            'FCS_MGL_comp':FCS_MGL_comp,
            'FCS_MGL_late_percentage':FCS_MGL_late_percentage,
            #---------------------------------------------------------------
            'FCS_PGL_sched':FCS_PGL_sched,
            'FCS_PGL_late':FCS_PGL_late,
            'FCS_PGL_comp':FCS_PGL_comp,
            'FCS_PGL_late_percentage':FCS_PGL_late_percentage,
            #---------------------------------------------------------------
            'FCS_EXPO_sched':FCS_EXPO_sched,
            'FCS_EXPO_late':FCS_EXPO_late,
            'FCS_EXPO_comp':FCS_EXPO_comp,
            'FCS_EXPO_late_percentage':FCS_EXPO_late_percentage,
            #---------------------------------------------------------------
            'FCS_MOL_sched':FCS_MOL_sched,
            'FCS_MOL_late':FCS_MOL_late,
            'FCS_MOL_comp':FCS_MOL_comp,
            'FCS_MOL_late_percentage':FCS_MOL_late_percentage,
            #---------------------------------------------------------------
            'FCS_OTHER_sched':FCS_OTHER_sched,
            'FCS_OTHER_late':FCS_OTHER_late,
            'FCS_OTHER_comp':FCS_OTHER_comp,
            'FCS_OTHER_late_percentage':FCS_OTHER_late_percentage,
            #------------------------------------------------------------------------------------------------
            #-------------------------------------------------------------------------------------------------
            #-------------------------------------------------------------------------------------------------
            'TPIS_MBL_sched':TPIS_MBL_sched,
            'TPIS_MBL_late':TPIS_MBL_late,
            'TPIS_MBL_comp':TPIS_MBL_comp,
            'TPIS_MBL_late_percentage':TPIS_MBL_late_percentage,
            #---------------------------------------------------------------
            'TPIS_MRL_MPL_sched':TPIS_MRL_MPL_sched,
            'TPIS_MRL_MPL_late':TPIS_MRL_MPL_late,
            'TPIS_MRL_MPL_comp':TPIS_MRL_MPL_comp,
            'TPIS_MRL_MPL_late_percentage':TPIS_MRL_MPL_late_percentage,
            #---------------------------------------------------------------
            'TPIS_MGL_sched':TPIS_MGL_sched,
            'TPIS_MGL_late':TPIS_MGL_late,
            'TPIS_MGL_comp':TPIS_MGL_comp,
            'TPIS_MGL_late_percentage':TPIS_MGL_late_percentage,
            #---------------------------------------------------------------
            'TPIS_PGL_sched':TPIS_PGL_sched,
            'TPIS_PGL_late':TPIS_PGL_late,
            'TPIS_PGL_comp':TPIS_PGL_comp,
            'TPIS_PGL_late_percentage':TPIS_PGL_late_percentage,
            #---------------------------------------------------------------
            'TPIS_EXPO_sched':TPIS_EXPO_sched,
            'TPIS_EXPO_late':TPIS_EXPO_late,
            'TPIS_EXPO_comp':TPIS_EXPO_comp,
            'TPIS_EXPO_late_percentage':TPIS_EXPO_late_percentage,
            #---------------------------------------------------------------
            'TPIS_MOL_sched':TPIS_MOL_sched,
            'TPIS_MOL_late':TPIS_MOL_late,
            'TPIS_MOL_comp':TPIS_MOL_comp,
            'TPIS_MOL_late_percentage':TPIS_MOL_late_percentage,
            #---------------------------------------------------------------
            'TPIS_OTHER_sched':TPIS_OTHER_sched,
            'TPIS_OTHER_late':TPIS_OTHER_late,
            'TPIS_OTHER_comp':TPIS_OTHER_comp,
            'TPIS_OTHER_late_percentage':TPIS_OTHER_late_percentage,
            #------------------------------------------------------------------------------------------------
            #-------------------------------------------------------------------------------------------------
            
            'BAT_RECT_MBL_sched':BAT_RECT_MBL_sched,
            'BAT_RECT_MBL_late':BAT_RECT_MBL_late,
            'BAT_RECT_MBL_comp':BAT_RECT_MBL_comp,
            'BAT_RECT_MBL_late_percentage':BAT_RECT_MBL_late_percentage,
            #---------------------------------------------------------------
            'BAT_RECT_MRL_MPL_sched':BAT_RECT_MRL_MPL_sched,
            'BAT_RECT_MRL_MPL_late':BAT_RECT_MRL_MPL_late,
            'BAT_RECT_MRL_MPL_comp':BAT_RECT_MRL_MPL_comp,
            'BAT_RECT_MRL_MPL_late_percentage':BAT_RECT_MRL_MPL_late_percentage,
            #---------------------------------------------------------------
            'BAT_RECT_MGL_sched':BAT_RECT_MGL_sched,
            'BAT_RECT_MGL_late':BAT_RECT_MGL_late,
            'BAT_RECT_MGL_comp':BAT_RECT_MGL_comp,
            'BAT_RECT_MGL_late_percentage':BAT_RECT_MGL_late_percentage,
            #---------------------------------------------------------------
            'BAT_RECT_PGL_sched':BAT_RECT_PGL_sched,
            'BAT_RECT_PGL_late':BAT_RECT_PGL_late,
            'BAT_RECT_PGL_comp':BAT_RECT_PGL_comp,
            'BAT_RECT_PGL_late_percentage':BAT_RECT_PGL_late_percentage,
            #---------------------------------------------------------------
            'BAT_RECT_EXPO_sched':BAT_RECT_EXPO_sched,
            'BAT_RECT_EXPO_late':BAT_RECT_EXPO_late,
            'BAT_RECT_EXPO_comp':BAT_RECT_EXPO_comp,
            'BAT_RECT_EXPO_late_percentage':BAT_RECT_EXPO_late_percentage,
            #---------------------------------------------------------------
            'BAT_RECT_MOL_sched':BAT_RECT_MOL_sched,
            'BAT_RECT_MOL_late':BAT_RECT_MOL_late,
            'BAT_RECT_MOL_comp':BAT_RECT_MOL_comp,
            'BAT_RECT_MOL_late_percentage':BAT_RECT_MOL_late_percentage,
            #---------------------------------------------------------------
            'BAT_RECT_OTHER_sched':BAT_RECT_OTHER_sched,
            'BAT_RECT_OTHER_late':BAT_RECT_OTHER_late,
            'BAT_RECT_OTHER_comp':BAT_RECT_OTHER_comp,
            'BAT_RECT_OTHER_late_percentage':BAT_RECT_OTHER_late_percentage,
            #------------------------------------------------------------------------------------------------
            #-------------------------------------------------------------------------------------------------
            #-------------------------------------------------------------------------------------------------
            'GASANA_MBL_sched':GASANA_MBL_sched,
            'GASANA_MBL_late':GASANA_MBL_late,
            'GASANA_MBL_comp':GASANA_MBL_comp,
            'GASANA_MBL_late_percentage':GASANA_MBL_late_percentage,
            #---------------------------------------------------------------
            'GASANA_MRL_MPL_sched':GASANA_MRL_MPL_sched,
            'GASANA_MRL_MPL_late':GASANA_MRL_MPL_late,
            'GASANA_MRL_MPL_comp':GASANA_MRL_MPL_comp,
            'GASANA_MRL_MPL_late_percentage':GASANA_MRL_MPL_late_percentage,
            #---------------------------------------------------------------
            'GASANA_MGL_sched':GASANA_MGL_sched,
            'GASANA_MGL_late':GASANA_MGL_late,
            'GASANA_MGL_comp':GASANA_MGL_comp,
            'GASANA_MGL_late_percentage':GASANA_MGL_late_percentage,
            #---------------------------------------------------------------
            'GASANA_PGL_sched':GASANA_PGL_sched,
            'GASANA_PGL_late':GASANA_PGL_late,
            'GASANA_PGL_comp':GASANA_PGL_comp,
            'GASANA_PGL_late_percentage':GASANA_PGL_late_percentage,
            #---------------------------------------------------------------
            'GASANA_EXPO_sched':GASANA_EXPO_sched,
            'GASANA_EXPO_late':GASANA_EXPO_late,
            'GASANA_EXPO_comp':GASANA_EXPO_comp,
            'GASANA_EXPO_late_percentage':GASANA_EXPO_late_percentage,
            #---------------------------------------------------------------
            'GASANA_MOL_sched':GASANA_MOL_sched,
            'GASANA_MOL_late':GASANA_MOL_late,
            'GASANA_MOL_comp':GASANA_MOL_comp,
            'GASANA_MOL_late_percentage':GASANA_MOL_late_percentage,
            #---------------------------------------------------------------
            'GASANA_OTHER_sched':GASANA_OTHER_sched,
            'GASANA_OTHER_late':GASANA_OTHER_late,
            'GASANA_OTHER_comp':GASANA_OTHER_comp,
            'GASANA_OTHER_late_percentage':GASANA_OTHER_late_percentage,
            #------------------------------------------------------------------------------------------------
            #-------------------------------------------------------------------------------------------------
            #-------------------------------------------------------------------------------------------------
            'RAD_MBL_sched':RAD_MBL_sched,
            'RAD_MBL_late':RAD_MBL_late,
            'RAD_MBL_comp':RAD_MBL_comp,
            'RAD_MBL_late_percentage':RAD_MBL_late_percentage,
            #---------------------------------------------------------------
            'RAD_MRL_MPL_sched':RAD_MRL_MPL_sched,
            'RAD_MRL_MPL_late':RAD_MRL_MPL_late,
            'RAD_MRL_MPL_comp':RAD_MRL_MPL_comp,
            'RAD_MRL_MPL_late_percentage':RAD_MRL_MPL_late_percentage,
            #---------------------------------------------------------------
            'RAD_MGL_sched':RAD_MGL_sched,
            'RAD_MGL_late':RAD_MGL_late,
            'RAD_MGL_comp':RAD_MGL_comp,
            'RAD_MGL_late_percentage':RAD_MGL_late_percentage,
            #---------------------------------------------------------------
            'RAD_PGL_sched':RAD_PGL_sched,
            'RAD_PGL_late':RAD_PGL_late,
            'RAD_PGL_comp':RAD_PGL_comp,
            'RAD_PGL_late_percentage':RAD_PGL_late_percentage,
            #---------------------------------------------------------------
            'RAD_EXPO_sched':RAD_EXPO_sched,
            'RAD_EXPO_late':RAD_EXPO_late,
            'RAD_EXPO_comp':RAD_EXPO_comp,
            'RAD_EXPO_late_percentage':RAD_EXPO_late_percentage,
            #---------------------------------------------------------------
            'RAD_MOL_sched':RAD_MOL_sched,
            'RAD_MOL_late':RAD_MOL_late,
            'RAD_MOL_comp':RAD_MOL_comp,
            'RAD_MOL_late_percentage':RAD_MOL_late_percentage,
            #---------------------------------------------------------------
            'RAD_OTHER_sched':RAD_OTHER_sched,
            'RAD_OTHER_late':RAD_OTHER_late,
            'RAD_OTHER_comp':RAD_OTHER_comp,
            'RAD_OTHER_late_percentage':RAD_OTHER_late_percentage,
            #------------------------------------------------------------------------------------------------
            #-------------------------------------------------------------------------------------------------
            #-------------------------------------------------------------------------------------------------
            'SEIS_MBL_sched':SEIS_MBL_sched,
            'SEIS_MBL_late':SEIS_MBL_late,
            'SEIS_MBL_comp':SEIS_MBL_comp,
            'SEIS_MBL_late_percentage':SEIS_MBL_late_percentage,
            #---------------------------------------------------------------
            'SEIS_MRL_MPL_sched':SEIS_MRL_MPL_sched,
            'SEIS_MRL_MPL_late':SEIS_MRL_MPL_late,
            'SEIS_MRL_MPL_comp':SEIS_MRL_MPL_comp,
            'SEIS_MRL_MPL_late_percentage':SEIS_MRL_MPL_late_percentage,
            #---------------------------------------------------------------
            'SEIS_MGL_sched':SEIS_MGL_sched,
            'SEIS_MGL_late':SEIS_MGL_late,
            'SEIS_MGL_comp':SEIS_MGL_comp,
            'SEIS_MGL_late_percentage':SEIS_MGL_late_percentage,
            #---------------------------------------------------------------
            'SEIS_PGL_sched':SEIS_PGL_sched,
            'SEIS_PGL_late':SEIS_PGL_late,
            'SEIS_PGL_comp':SEIS_PGL_comp,
            'SEIS_PGL_late_percentage':SEIS_PGL_late_percentage,
            #---------------------------------------------------------------
            'SEIS_EXPO_sched':SEIS_EXPO_sched,
            'SEIS_EXPO_late':SEIS_EXPO_late,
            'SEIS_EXPO_comp':SEIS_EXPO_comp,
            'SEIS_EXPO_late_percentage':SEIS_EXPO_late_percentage,
            #---------------------------------------------------------------
            'SEIS_MOL_sched':SEIS_MOL_sched,
            'SEIS_MOL_late':SEIS_MOL_late,
            'SEIS_MOL_comp':SEIS_MOL_comp,
            'SEIS_MOL_late_percentage':SEIS_MOL_late_percentage,
            #---------------------------------------------------------------
            'SEIS_OTHER_sched':SEIS_OTHER_sched,
            'SEIS_OTHER_late':SEIS_OTHER_late,
            'SEIS_OTHER_comp':SEIS_OTHER_comp,
            'SEIS_OTHER_late_percentage':SEIS_OTHER_late_percentage,
            #------------------------------------------------------------------------------------------------
            #-------------------------------------------------------------------------------------------------
            #-------------------------------------------------------------------------------------------------
            'TRAINCC_MBL_sched':TRAINCC_MBL_sched,
            'TRAINCC_MBL_late':TRAINCC_MBL_late,
            'TRAINCC_MBL_comp':TRAINCC_MBL_comp,
            'TRAINCC_MBL_late_percentage':TRAINCC_MBL_late_percentage,
            #---------------------------------------------------------------
            'TRAINCC_MRL_MPL_sched':TRAINCC_MRL_MPL_sched,
            'TRAINCC_MRL_MPL_late':TRAINCC_MRL_MPL_late,
            'TRAINCC_MRL_MPL_comp':TRAINCC_MRL_MPL_comp,
            'TRAINCC_MRL_MPL_late_percentage':TRAINCC_MRL_MPL_late_percentage,
            #---------------------------------------------------------------
            'TRAINCC_MGL_sched':TRAINCC_MGL_sched,
            'TRAINCC_MGL_late':TRAINCC_MGL_late,
            'TRAINCC_MGL_comp':TRAINCC_MGL_comp,
            'TRAINCC_MGL_late_percentage':TRAINCC_MGL_late_percentage,
            #---------------------------------------------------------------
            'TRAINCC_PGL_sched':TRAINCC_PGL_sched,
            'TRAINCC_PGL_late':TRAINCC_PGL_late,
            'TRAINCC_PGL_comp':TRAINCC_PGL_comp,
            'TRAINCC_PGL_late_percentage':TRAINCC_PGL_late_percentage,
            #---------------------------------------------------------------
            'TRAINCC_EXPO_sched':TRAINCC_EXPO_sched,
            'TRAINCC_EXPO_late':TRAINCC_EXPO_late,
            'TRAINCC_EXPO_comp':TRAINCC_EXPO_comp,
            'TRAINCC_EXPO_late_percentage':TRAINCC_EXPO_late_percentage,
            #---------------------------------------------------------------
            'TRAINCC_MOL_sched':TRAINCC_MOL_sched,
            'TRAINCC_MOL_late':TRAINCC_MOL_late,
            'TRAINCC_MOL_comp':TRAINCC_MOL_comp,
            'TRAINCC_MOL_late_percentage':TRAINCC_MOL_late_percentage,
            #---------------------------------------------------------------
            'TRAINCC_OTHER_sched':TRAINCC_OTHER_sched,
            'TRAINCC_OTHER_late':TRAINCC_OTHER_late,
            'TRAINCC_OTHER_comp':TRAINCC_OTHER_comp,
            'TRAINCC_OTHER_late_percentage':TRAINCC_OTHER_late_percentage,
            #------------------------------------------------------------------------------------------------
            #-------------------------------------------------------------------------------------------------
            #-------------------------------------------------------------------------------------------------
            
            'TRAINPRO_MBL_sched':TRAINPRO_MBL_sched,
            'TRAINPRO_MBL_late':TRAINPRO_MBL_late,
            'TRAINPRO_MBL_comp':TRAINPRO_MBL_comp,
            'TRAINPRO_MBL_late_percentage':TRAINPRO_MBL_late_percentage,
            #---------------------------------------------------------------
            'TRAINPRO_MRL_MPL_sched':TRAINPRO_MRL_MPL_sched,
            'TRAINPRO_MRL_MPL_late':TRAINPRO_MRL_MPL_late,
            'TRAINPRO_MRL_MPL_comp':TRAINPRO_MRL_MPL_comp,
            'TRAINPRO_MRL_MPL_late_percentage':TRAINPRO_MRL_MPL_late_percentage,
            #---------------------------------------------------------------
            'TRAINPRO_MGL_sched':TRAINPRO_MGL_sched,
            'TRAINPRO_MGL_late':TRAINPRO_MGL_late,
            'TRAINPRO_MGL_comp':TRAINPRO_MGL_comp,
            'TRAINPRO_MGL_late_percentage':TRAINPRO_MGL_late_percentage,
            #---------------------------------------------------------------
            'TRAINPRO_PGL_sched':TRAINPRO_PGL_sched,
            'TRAINPRO_PGL_late':TRAINPRO_PGL_late,
            'TRAINPRO_PGL_comp':TRAINPRO_PGL_comp,
            'TRAINPRO_PGL_late_percentage':TRAINPRO_PGL_late_percentage,
            #---------------------------------------------------------------
            'TRAINPRO_EXPO_sched':TRAINPRO_EXPO_sched,
            'TRAINPRO_EXPO_late':TRAINPRO_EXPO_late,
            'TRAINPRO_EXPO_comp':TRAINPRO_EXPO_comp,
            'TRAINPRO_EXPO_late_percentage':TRAINPRO_EXPO_late_percentage,
            #---------------------------------------------------------------
            'TRAINPRO_MOL_sched':TRAINPRO_MOL_sched,
            'TRAINPRO_MOL_late':TRAINPRO_MOL_late,
            'TRAINPRO_MOL_comp':TRAINPRO_MOL_comp,
            'TRAINPRO_MOL_late_percentage':TRAINPRO_MOL_late_percentage,
            #---------------------------------------------------------------
            'TRAINPRO_OTHER_sched':TRAINPRO_OTHER_sched,
            'TRAINPRO_OTHER_late':TRAINPRO_OTHER_late,
            'TRAINPRO_OTHER_comp':TRAINPRO_OTHER_comp,
            'TRAINPRO_OTHER_late_percentage':TRAINPRO_OTHER_late_percentage,
            #------------------------------------------------------------------------------------------------
            #-------------------------------------------------------------------------------------------------
            #-------------------------------------------------------------------------------------------------
            'TRAINRAD_MBL_sched':TRAINRAD_MBL_sched,
            'TRAINRAD_MBL_late':TRAINRAD_MBL_late,
            'TRAINRAD_MBL_comp':TRAINRAD_MBL_comp,
            'TRAINRAD_MBL_late_percentage':TRAINRAD_MBL_late_percentage,
            #---------------------------------------------------------------
            'TRAINRAD_MRL_MPL_sched':TRAINRAD_MRL_MPL_sched,
            'TRAINRAD_MRL_MPL_late':TRAINRAD_MRL_MPL_late,
            'TRAINRAD_MRL_MPL_comp':TRAINRAD_MRL_MPL_comp,
            'TRAINRAD_MRL_MPL_late_percentage':TRAINRAD_MRL_MPL_late_percentage,
            #---------------------------------------------------------------
            'TRAINRAD_MGL_sched':TRAINRAD_MGL_sched,
            'TRAINRAD_MGL_late':TRAINRAD_MGL_late,
            'TRAINRAD_MGL_comp':TRAINRAD_MGL_comp,
            'TRAINRAD_MGL_late_percentage':TRAINRAD_MGL_late_percentage,
            #---------------------------------------------------------------
            'TRAINRAD_PGL_sched':TRAINRAD_PGL_sched,
            'TRAINRAD_PGL_late':TRAINRAD_PGL_late,
            'TRAINRAD_PGL_comp':TRAINRAD_PGL_comp,
            'TRAINRAD_PGL_late_percentage':TRAINRAD_PGL_late_percentage,
            #---------------------------------------------------------------
            'TRAINRAD_EXPO_sched':TRAINRAD_EXPO_sched,
            'TRAINRAD_EXPO_late':TRAINRAD_EXPO_late,
            'TRAINRAD_EXPO_comp':TRAINRAD_EXPO_comp,
            'TRAINRAD_EXPO_late_percentage':TRAINRAD_EXPO_late_percentage,
            #---------------------------------------------------------------
            'TRAINRAD_MOL_sched':TRAINRAD_MOL_sched,
            'TRAINRAD_MOL_late':TRAINRAD_MOL_late,
            'TRAINRAD_MOL_comp':TRAINRAD_MOL_comp,
            'TRAINRAD_MOL_late_percentage':TRAINRAD_MOL_late_percentage,
            #---------------------------------------------------------------
            'TRAINRAD_OTHER_sched':TRAINRAD_OTHER_sched,
            'TRAINRAD_OTHER_late':TRAINRAD_OTHER_late,
            'TRAINRAD_OTHER_comp':TRAINRAD_OTHER_comp,
            'TRAINRAD_OTHER_late_percentage':TRAINRAD_OTHER_late_percentage,
            #------------------------------------------------------------------------------------------------
            #-------------------------------------------------------------------------------------------------
            #-------------------------------------------------------------------------------------------------
            'IDS_MBL_sched':IDS_MBL_sched,
            'IDS_MBL_late':IDS_MBL_late,
            'IDS_MBL_comp':IDS_MBL_comp,
            'IDS_MBL_late_percentage':IDS_MBL_late_percentage,
            #---------------------------------------------------------------
            'IDS_MRL_MPL_sched':IDS_MRL_MPL_sched,
            'IDS_MRL_MPL_late':IDS_MRL_MPL_late,
            'IDS_MRL_MPL_comp':IDS_MRL_MPL_comp,
            'IDS_MRL_MPL_late_percentage':IDS_MRL_MPL_late_percentage,
            #---------------------------------------------------------------
            'IDS_MGL_sched':IDS_MGL_sched,
            'IDS_MGL_late':IDS_MGL_late,
            'IDS_MGL_comp':IDS_MGL_comp,
            'IDS_MGL_late_percentage':IDS_MGL_late_percentage,
            #---------------------------------------------------------------
            'IDS_PGL_sched':IDS_PGL_sched,
            'IDS_PGL_late':IDS_PGL_late,
            'IDS_PGL_comp':IDS_PGL_comp,
            'IDS_PGL_late_percentage':IDS_PGL_late_percentage,
            #---------------------------------------------------------------
            'IDS_EXPO_sched':IDS_EXPO_sched,
            'IDS_EXPO_late':IDS_EXPO_late,
            'IDS_EXPO_comp':IDS_EXPO_comp,
            'IDS_EXPO_late_percentage':IDS_EXPO_late_percentage,
            #---------------------------------------------------------------
            'IDS_MOL_sched':IDS_MOL_sched,
            'IDS_MOL_late':IDS_MOL_late,
            'IDS_MOL_comp':IDS_MOL_comp,
            'IDS_MOL_late_percentage':IDS_MOL_late_percentage,
            #---------------------------------------------------------------
            'IDS_OTHER_sched':IDS_OTHER_sched,
            'IDS_OTHER_late':IDS_OTHER_late,
            'IDS_OTHER_comp':IDS_OTHER_comp,
            'IDS_OTHER_late_percentage':IDS_OTHER_late_percentage
            
            }

    doc.render(context)
    doc.save('Template_Rendered.docx')
    root1.destroy()

label = tk.Label(root1, text = 'Please select excel file you would like to turn into data')
label.pack(pady=10)
my_btn = Button(root1, text = "Open File", command = open).pack(pady = 30)
root1.mainloop()


win = Tk()
win.title("Please exit when finished")
Label(win, text=" Tables created will be located within Template_Rendered.docx ", font=('Helvetica 14 bold')).pack(pady=40)
win_width = 750
win_height = 250
x = int(int(win.winfo_screenwidth()/2) - int(win_width/2))
y = int(int(win.winfo_screenheight()/2) - int(win_height/2))
win.geometry(f"{win_width}x{win_height}+{x}+{y}")
win.mainloop()