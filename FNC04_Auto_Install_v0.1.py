# -*- coding: utf-8 -*-
import subprocess
import uiautomation as auto
import ctypes, sys
import time


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


# 禁止鼠标键盘操作
lock = ctypes.windll.LoadLibrary('user32.dll')
lock.BlockInput(True)

if is_admin():
    # install heaven benchmark
    print('installation will begin')
    subprocess.Popen('FNC040425\\Unigine_Heaven-4.0-[Guru3D.com].exe')
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
    # install PassMark Performance 9
    subprocess.Popen('FNC040425\\PassMark Performance 9 (Build 1031).exe')
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
        print('遇到错误：', e)
        passMarkPerf.ButtonControl(searchDepth=2, Name='Install').Click()
        time.sleep(8)
        passMarkPerf.CheckBoxControl(searchDepth=5, Name='Launch PerformanceTest').Click()
        passMarkPerf.ButtonControl(searchDepth=2, Name='Finish').Click()
    finally:
        # install passmark monitor
        subprocess.Popen('FNC040425\\PassMark Monitor_V3.2 (Build 1006).exe')
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
            print('遇到错误：', e)
            passMarkMonitor.ButtonControl(searchDepth=2, Name='Install').Click()
            time.sleep(8)
            passMarkMonitor.ButtonControl(searchDepth=2, Name='Next >').Click()
            passMarkMonitor.CheckBoxControl(searchDepth=5, Name='Launch MonitorTest').Click()
            passMarkMonitor.ButtonControl(searchDepth=2, Name='Finish').Click()
        finally:
            # install Furmark
            subprocess.Popen('FNC040425\\FurMark_1.20.5.0_Setup.exe')
            time.sleep(2)
            furMark = auto.WindowControl(searchDepth=1, ClassName='TWizardForm')
            furMark.SetTopmost(True,waitTime=1)
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
                print('遇到错误：', e)
                folderExistWindow = auto.WindowControl(searchDepth=1, Name='Folder Exists')
                folderExistWindow.SetTopmost(True)
                time.sleep(2)
                try:
                    folderExistWindow.ButtonControl(searchDepth=2, Name='是(Y)').Click()
                except LookupError as e:
                    print('遇到错误：', e)
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
            finally:
                # install AIDA64
                subprocess.Popen('FNC040425\\aida64extreme_5.99.4900.exe')
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
                    print('遇到错误：', e)
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
                finally:
                    lock.BlockInput(False)
else:
    # Re-run the program with admin rights
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)