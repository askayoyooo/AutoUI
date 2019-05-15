# -*- coding: utf-8 -*-
import subprocess
import uiautomation as auto
import ctypes, sys, os
import time


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


def file_name(path, ext):
    # 获取某路径下扩展名为ext 的文件路径
    L = []
    for file in os.listdir(path):
        if os.path.splitext(file)[1] == ext:
            L.append(os.path.join(path, file))
    return L


def install_heaven_bench(tool_full_path):
    # install heaven benchmark
    print('installing heaven benchmark ')
    subprocess.Popen(tool_full_path)
    time.sleep(2)
    heavenWindow = auto.WindowControl(searchDepth=1, ClassName='TWizardForm')
    heavenWindow.SetTopmost(True)
    heavenWindow.ButtonControl(searchDepth=2, Name='Next >').Click()
    heavenWindow.RadioButtonControl(searchDepth=6, Name='I accept the agreement').Click()
    heavenWindow.ButtonControl(searchDepth=2, Name='Next >').Click()
    heavenWindow.ButtonControl(searchDepth=2, Name='Next >').Click()
    heavenWindow.ButtonControl(searchDepth=2, Name='Next >').Click()
    heavenWindow.ButtonControl(searchDepth=2, Name='Next >').Click()
    heavenWindow.ButtonControl(searchDepth=2, Name='Install').Click()
    time.sleep(5)
    heavenWindow.ButtonControl(searchDepth=2, Name='Finish').Click()
    print('heaven benchmark installed ')


def install_passmark_performance(tool_full_path):
    # install PassMark Performance 9
    print('installing PassMark Performance 9 ')
    subprocess.Popen(tool_full_path)
    time.sleep(2)
    passMarkPerf = auto.WindowControl(searchDepth=1, ClassName='TWizardForm')
    passMarkPerf.SetTopmost(True)
    passMarkPerf.RadioButtonControl(searchDepth=6, Name='I accept the agreement').Click()
    passMarkPerf.ButtonControl(searchDepth=2, Name='Next >').Click()
    try:
        passMarkPerf.ButtonControl(searchDepth=2, Name='Next >').Click()
        passMarkPerf.ButtonControl(searchDepth=2, Name='Next >').Click()
        passMarkPerf.ButtonControl(searchDepth=2, Name='Next >').Click()
        passMarkPerf.ButtonControl(searchDepth=2, Name='Install').Click()
        time.sleep(8)
        passMarkPerf.CheckBoxControl(searchDepth=5, Name='Launch PerformanceTest').Click()
        passMarkPerf.ButtonControl(searchDepth=2, Name='Finish').Click()
    except LookupError as e:
        # print('遇到错误：', e)
        passMarkPerf.ButtonControl(searchDepth=2, Name='Install').Click()
        time.sleep(8)
        passMarkPerf.CheckBoxControl(searchDepth=5, Name='Launch PerformanceTest').Click()
        passMarkPerf.ButtonControl(searchDepth=2, Name='Finish').Click()
    print('PassMark Performance installed ')


def install_passmark_monitor(tool_full_path):
    # install passmark monitor
    print('installing passmark monitor')
    subprocess.Popen(tool_full_path)
    time.sleep(2)
    passMarkMonitor = auto.WindowControl(searchDepth=1, ClassName='TWizardForm')
    passMarkMonitor.SetTopmost(True)
    passMarkMonitor.RadioButtonControl(searchDepth=6, Name='I accept the agreement').Click()
    passMarkMonitor.ButtonControl(searchDepth=2, Name='Next >').Click()
    try:
        passMarkMonitor.ButtonControl(searchDepth=2, Name='Next >').Click()
        passMarkMonitor.ButtonControl(searchDepth=2, Name='Next >').Click()
        passMarkMonitor.ButtonControl(searchDepth=2, Name='Next >').Click()
        passMarkMonitor.ButtonControl(searchDepth=2, Name='Install').Click()
        time.sleep(8)
        passMarkMonitor.ButtonControl(searchDepth=2, Name='Next >').Click()
        passMarkMonitor.CheckBoxControl(searchDepth=5, Name='Launch MonitorTest').Click()
        passMarkMonitor.ButtonControl(searchDepth=2, Name='Finish').Click()
    except LookupError as e:
        # print('遇到错误：', e)
        passMarkMonitor.ButtonControl(searchDepth=2, Name='Install').Click()
        time.sleep(8)
        passMarkMonitor.ButtonControl(searchDepth=2, Name='Next >').Click()
        passMarkMonitor.CheckBoxControl(searchDepth=5, Name='Launch MonitorTest').Click()
        passMarkMonitor.ButtonControl(searchDepth=2, Name='Finish').Click()
    print('passmark monitor installed')

def install_furmark(tool_full_path):
    # install Furmark
    print('installing furmark')
    subprocess.Popen(tool_full_path)
    time.sleep(2)
    furMark = auto.WindowControl(searchDepth=1, ClassName='TWizardForm')
    furMark.SetTopmost(True, waitTime=1)
    furMark.RadioButtonControl(searchDepth=6, Name='I accept the agreement').Click()
    furMark.ButtonControl(searchDepth=2, Name='Next >').Click()
    furMark.ButtonControl(searchDepth=2, Name='Next >').Click()
    try:
        furMark.ButtonControl(searchDepth=2, Name='Next >').Click()
        furMark.CheckBoxControl(searchDepth=7, Name='Create a &desktop shortcut').Click()
        furMark.CheckBoxControl(searchDepth=7, Name='Create a &Quick Launch icon').Click()
        furMark.ButtonControl(searchDepth=2, Name='Next >').Click()
        furMark.ButtonControl(searchDepth=2, Name='Install').Click()
        time.sleep(8)
        furMark.ButtonControl(searchDepth=2, Name='Next >').Click()
        furMark.CheckBoxControl(searchDepth=5, Name='Launch FurMark').Click()
        furMark.CheckBoxControl(searchDepth=5, Name='Launch FurMark release notes').Click()
        furMark.ButtonControl(searchDepth=2, Name='Finish').Click()
    except LookupError as e:
        # print('遇到错误：', e)
        folderExistWindow = auto.WindowControl(searchDepth=1, Name='Folder Exists')
        folderExistWindow.SetTopmost(True)
        time.sleep(2)
        try:
            folderExistWindow.ButtonControl(searchDepth=2, Name='是(Y)').Click()
        except LookupError as e:
            # print('遇到错误：', e)
            folderExistWindow.ButtonControl(searchDepth=2, Name='Yes').Click()
        furMark.ButtonControl(searchDepth=2, Name='Next >').Click()
        furMark.CheckBoxControl(searchDepth=7, Name='Create a &desktop shortcut').Click()
        furMark.CheckBoxControl(searchDepth=7, Name='Create a &Quick Launch icon').Click()
        furMark.ButtonControl(searchDepth=2, Name='Next >').Click()
        furMark.ButtonControl(searchDepth=2, Name='Install').Click()
        time.sleep(8)
        furMark.ButtonControl(searchDepth=2, Name='Next >').Click()
        furMark.CheckBoxControl(searchDepth=5, Name='Launch FurMark').Click()
        furMark.CheckBoxControl(searchDepth=5, Name='Launch FurMark release notes').Click()
        furMark.ButtonControl(searchDepth=2, Name='Finish').Click()
    print('furmark installed')


def install_aida(tool_full_path):
    print('installing aida64')
    subprocess.Popen(tool_full_path)
    time.sleep(2)
    aidaLanguage = auto.WindowControl(searchDepth=1, ClassName='TSelectLanguageForm')
    aidaLanguage.SetTopmost(True)
    aidaLanguage.ComboBoxControl(searchDepth=2).Click()
    aidaLanguage.ListItemControl(searchDepth=4, Name='English').Click()
    time.sleep(2)
    try:
        aidaLanguage.ButtonControl(searchDepth=2, Name='OK').Click()
        aida = auto.WindowControl(searchDepth=1, Name='Setup - AIDA64 Extreme')
        aida.SetTopmost(True)
        aida.ButtonControl(searchDepth=2, Name='Next >').Click()
        aida.RadioButtonControl(searchDepth=6, Name='I accept the agreement').Click()
        aida.ButtonControl(searchDepth=2, Name='Next >').Click()
        aida.ButtonControl(searchDepth=2, Name='Next >').Click()
        aida.ButtonControl(searchDepth=2, Name='Next >').Click()
        aida.ButtonControl(searchDepth=2, Name='Next >').Click()
        aida.ButtonControl(searchDepth=2, Name='Install').Click()
        time.sleep(8)
        aida.CheckBoxControl(searchDepth=5, Name='Launch AIDA64 Extreme').Click()
        aida.CheckBoxControl(searchDepth=5, Name='View AIDA64 Extreme Documentation').Click()
        aida.ButtonControl(searchDepth=2, Name='Finish').Click()
    except LookupError as e:
        #print('遇到错误：', e)
        aidaLanguage.ButtonControl(searchDepth=2, Name='确定').Click()
        aida = auto.WindowControl(searchDepth=1, Name='Setup - AIDA64 Extreme')
        aida.SetTopmost(True)
        aida.ButtonControl(searchDepth=2, Name='Next >').Click()
        aida.RadioButtonControl(searchDepth=6, Name='I accept the agreement').Click()
        aida.ButtonControl(searchDepth=2, Name='Next >').Click()
        aida.ButtonControl(searchDepth=2, Name='Next >').Click()
        aida.ButtonControl(searchDepth=2, Name='Next >').Click()
        aida.ButtonControl(searchDepth=2, Name='Next >').Click()
        aida.ButtonControl(searchDepth=2, Name='Install').Click()
        time.sleep(8)
        aida.CheckBoxControl(searchDepth=5, Name='Launch AIDA64 Extreme').Click()
        aida.CheckBoxControl(searchDepth=5, Name='View AIDA64 Extreme Documentation').Click()
        aida.ButtonControl(searchDepth=2, Name='Finish').Click()
    print('aida64 installed')


if is_admin():
    # 禁止鼠标键盘操作
    lock = ctypes.windll.LoadLibrary('user32.dll')
    lock.BlockInput(True)
    folder = 'fnc04'
    filenames = file_name(os.path.dirname(os.path.realpath(__file__))+'\\'+folder+'\\', '.exe')

    for file in filenames:
        if 'heaven' in file.lower():
            install_heaven_bench(file)
        elif 'furmark' in file.lower():
            install_furmark(file)
        elif 'monitor' in file.lower():
            install_passmark_monitor(file)
        elif 'performance' in file.lower():
            install_passmark_performance(file)
        elif 'aida' in file.lower():
            install_aida(file)


    lock.BlockInput(False)
else:
    # Re-run the program with admin rights
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)