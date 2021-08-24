#!/usr/bin/env python3
from pprint import pprint
from inspect import getmembers
import hashlib
import json
import os
import sys
import pythoncom

from .enums import get_dy_enum

from ..gpkgs.timeout import TimeOut
from ..gpkgs import message as msg
from ..gpkgs.prompt import prompt_boolean, prompt_multiple, prompt

from pywintypes import com_error
import pywintypes
import win32api
import win32com
import win32gui
import win32process
from win32com.client import Dispatch
import threading

import re
import time

def winEnumHandler(hwnd, ctx):
    if win32gui.IsWindowVisible( hwnd ):
        window_title=win32gui.GetWindowText(hwnd)
        if window_title != "":
            ctx[hwnd]=window_title

def focus_window(hwnd):
    # prevent _error pywintypes.error: (0, 'SetForegroundWindow', 'No error message is available')
    shell = win32com.client.Dispatch("WScript.Shell")
    shell.SendKeys('%')
    win32gui.SetForegroundWindow(hwnd)

def focus_workbook(dy_options):
    timer=TimeOut(3).start()
    while True:
        windows=dict()
        win32gui.EnumWindows( winEnumHandler, windows)
        found=False
        for hwnd in sorted(windows):
            window_title=windows[hwnd]
            if dy_options["window_title"] in window_title:
                focus_window(hwnd)
                found=True
                if dy_options["immediate"] is True:
                    shell = win32com.client.Dispatch("WScript.Shell")
                    time.sleep(.1)
                    shell.SendKeys("%{F11}")
                    time.sleep(.1)
                    shell.SendKeys("^g")
                    if dy_options["clear"] is True:
                        time.sleep(.1)
                        shell.SendKeys("^a")
                        shell.SendKeys("{DEL}")
                break
        if found is True:
            break
        elif timer.has_ended(pause=.05):
            msg.error("Can't focus window '{}'".format(dy_options["window_title"]), exit=1)

def macro(
    active_hwnd,
    clear,
    filenpa_workbook,
    macro_name,
    immediate=False,
    params=None,
    reset_macro=False,
    reset_macro_seconds=None,
):
    active_hwnd=win32gui.GetForegroundWindow()
    filen_workbook=os.path.basename(filenpa_workbook)
    if params is None:
        params=[]
    else:
        if isinstance(params, list) is False:
            params=[params]

    xl=None
    wb=None
    try:
        xl = win32com.client.GetActiveObject("Excel.Application")
        for tmp_wb in xl.Workbooks:
            if tmp_wb.Name == filen_workbook:
                if xl.Visible is False:
                    xl.Visible=True
                wb=tmp_wb
                break
    except com_error as e:
        pass

    if xl is None:
        xl = Dispatch("Excel.Application")
        xl.Visible = True

    if wb is None:
        wb = xl.Workbooks.Open(filenpa_workbook)

    wb.EnableAutoRecover=False
    dy_options=dict(
        clear=clear,
        immediate=immediate,
        window_title="{} - Excel".format(filen_workbook),
    )
    focus_workbook(dy_options)

    if reset_macro is True:
        dy_reset_options=dict(
            active_hwnd=active_hwnd,
            xl=xl, 
            has_ended=False,
            reset_macro_seconds=reset_macro_seconds,
        )

        th=threading.Thread(target=execute_reset_macro, args=(dy_reset_options,))
        th.start()

    cmd=[
        macro_name,
        *params
    ]

    try:
        xl.Run(*cmd)
    except com_error as e:
        manage_error(
            active_hwnd, 
            e, 
            "At '{}' when running macro '{}'".format(
                filen_workbook,
                macro_name,
            ),
        )
    else:
        msg.success("{} {}".format(filen_workbook, macro_name))
    finally:
        dy_reset_options["has_ended"]=True


def execute_reset_macro(dy_options):
    wait_seconds=dy_options["reset_macro_seconds"]
    if wait_seconds is None:
        wait_seconds=3
    pythoncom.CoInitialize()
    timer=TimeOut(wait_seconds).start()
    execute_reset=False
    while True:
    # Application.VBE.CommandBars(1).Controls("&Run").Controls("&Reset")
        if dy_options["has_ended"] is True:
            break 
        elif timer.has_ended(pause=.05):
            execute_reset=True
            break
    if execute_reset is True:
        try:
            # dy_options["xl"].VBE.CommandBars(1).Controls("&Run").Controls("&Reset").Execute()
            time.sleep(.1)
            shell = win32com.client.Dispatch("WScript.Shell")
            shell.SendKeys("{ENTER}")
            time.sleep(.1)
            win32com.client.GetActiveObject("Excel.Application").VBE.CommandBars(1).Controls("&Run").Controls("&Reset").Execute()
        except com_error as e:
            manage_error(
                dy_options["active_hwnd"], 
                e, 
                "issue when resetting macro",
            )

def manage_error(
    active_hwnd,
    e,
    custom_msg,
):
    time.sleep(.1)
    focus_window(active_hwnd)
    pprint(vars(e))
    e_msg=None
    error_code=None
    for i, elem in enumerate(e.excepinfo):
        if i == len(e.excepinfo) -1:
            try:
                error_code=elem
                e_msg=win32api.FormatMessage(elem).rstrip()
            except pywintypes.error as e:
                print(e)
                msg.error("error code '{}' not found with win32api.FormatMessage. Trying to extract error message anyway.".format(error_code))
        else:
            if isinstance(elem, str):
                if e_msg is None:
                    e_msg=elem
                else:
                    e_msg+=", {}".format(elem)
    
    msg.error("{} error '{}'".format(custom_msg, e_msg))
    sys.exit(1)

def export(
    active_hwnd,
    filenpa_workbook,
    direpa_srcs,
    overwrite,
):
    filen_workbook=os.path.basename(filenpa_workbook)
    xl=None
    wb=None
    try:
        xl = win32com.client.GetActiveObject("Excel.Application")
        for tmp_wb in xl.Workbooks:
            if tmp_wb.Name == filen_workbook:
                wb=tmp_wb
                break
    except com_error as e:
        pass

    if xl is None:
        xl = Dispatch("Excel.Application")
        xl.Visible = False

    if wb is None:
        wb = xl.Workbooks.Open(filenpa_workbook, 0, True)
    has_error=False

    if wb.VBProject.Protection == 1:
        has_error=True
        msg.error("Export Failed. The VBA in '{}' workbook is protected.".format(filen_workbook))
    else:
        dy_vbcomponents=get_dy_enum("VBComponents")
        for component in wb.VBProject.VBComponents:
            if component.Type in dy_vbcomponents:
                allowed=dy_vbcomponents[component.Type]["allowed"]
                if allowed is True:
                    ext=dy_vbcomponents[component.Type]["extension"]
                    module_name="{}{}".format(component.Name, ext)
                    to_export=True
                    filenpaXlsModule=os.path.join(direpa_srcs, module_name)
                    if os.path.exists(filenpaXlsModule) is True and overwrite is False:
                        msg.warning("Module '{}' already exists in '{}'".format(module_name, os.path.basename(direpa_srcs)))
                        to_export=prompt_boolean("Overwrite", "Y")

                    if to_export is True:
                        try:
                            component.Export(filenpaXlsModule)
                            msg.success("Module '{}' exported.".format(module_name))
                        except com_error as e:
                            manage_error(
                                active_hwnd, 
                                e, 
                                "At '{}' when exporting module '{}'".format(
                                    filen_workbook,
                                    module_name,
                                ),
                            )
                    else:
                        msg.info("Module '{}' ignored.".format(module_name))

    if xl.Visible is False:
        wb.Close()
        xl.Quit()

    if has_error is True:
        sys.exit(1)
    else:
        msg.success("Export modules from '{}' completed.".format(os.path.basename(filenpa_workbook)))

def _import(
    active_hwnd,
    filenpa_cache,
    filenpa_workbook,
    direpa_srcs,
    overwrite,
    reset_cache,
):
    filen_workbook=os.path.basename(filenpa_workbook)

    dy_cache=dict()
    if reset_cache is False:
        if os.path.exists(filenpa_cache):
            with open(filenpa_cache, "r") as f:
                dy_cache=json.load(f)

    if os.path.exists(direpa_srcs) is True:
        has_files=len([elem for elem in os.listdir(direpa_srcs) if os.path.splitext(elem)[1] in [".bas", ".frm", ".frx"]]) > 0
        if has_files is True:
            save_cache=False
            if filen_workbook not in dy_cache:
                save_cache=True
                dy_cache[filen_workbook]=dict()

            filens_update=[]
            all_filens=[]
            dy_vbcomponents=get_dy_enum("VBComponents")
            allowed_extensions=list(set([dy_vbcomponents[code]["extension"] for code in dy_vbcomponents if dy_vbcomponents[code]["allowed"] is True]))

            for elem in sorted(os.listdir(direpa_srcs)):
                update_file=False
                path_elem=os.path.join(direpa_srcs, elem)
                if os.path.isfile(path_elem):
                    filer, ext=os.path.splitext(elem)
                    if ext in allowed_extensions:
                        all_filens.append(elem)
                        md5=hashlib.md5(open(path_elem,'rb').read()).hexdigest()
                        if elem in dy_cache[filen_workbook]:
                            if md5 != dy_cache[filen_workbook][elem]:
                                update_file=True
                        else:
                            update_file=True

                        if update_file is True:
                            save_cache=True
                            filens_update.append(elem)
                            dy_cache[filen_workbook][elem]=md5

            for filen in sorted(dy_cache[filen_workbook]):
                if filen not in all_filens:
                    save_cache=True
                    del dy_cache[filen_workbook][filen]


            if len(filens_update) > 0:
                xl = Dispatch("Excel.Application")
                xl.Visible = True
                wb = xl.Workbooks.Open(filenpa_workbook)

                vbProj=wb.VBProject
                for component in vbProj.VBComponents:
                    if component.Type in dy_vbcomponents:
                        if dy_vbcomponents[component.Type]["allowed"]:
                            ext=dy_vbcomponents[component.Type]["extension"]
                            module_name="{}{}".format(component.Name, ext)
                            if module_name in filens_update:
                                to_import=overwrite
                                if overwrite is False:
                                    msg.warning("Module '{}' already exists in '{}'".format(module_name, os.path.basename(filen_workbook)))
                                    to_import=prompt_boolean("Overwrite", "Y")
                                if to_import is True:
                                    try:
                                        # objComponent=vbProj.VBComponents(component.Name)
                                        vbProj.VBComponents.Remove(component)
                                        msg.success("Module '{}' removed from '{}'.".format(module_name, filen_workbook))
                                    except com_error as e:
                                        pprint(vars(e))
                                        e_msg = win32api.FormatMessage(e.excepinfo[5]).rstrip()
                                        msg.error("At '{}' when deleting module '{}' error '{}'".format(
                                            filen_workbook,
                                            module_name,
                                            e_msg,
                                        ), exit=1)

                for module_name in filens_update:
                    filenpa_update=os.path.join(direpa_srcs, module_name)
                    vbProj.VBComponents.Import(filenpa_update)
                    msg.success("Module '{}' imported.".format(module_name))

                wb.Save()
            else:
                msg.info("Everything up-to-date")

            if save_cache is True:
                with open(filenpa_cache, "w") as f:
                    f.write(json.dumps(dy_cache, sort_keys=True, indent=4))
            
            msg.success("Import modules to '{}' completed.".format(os.path.basename(filenpa_workbook)))

        else:
            user_input=prompt_multiple([
                "create a new module .bas",
                "export {}".format(os.path.basename(filenpa_workbook)),
                ],
                values=["create", "export"],
                title="Choose action",
            )

            print()
            if user_input == "create":
                filenpa_module=os.path.join(direpa_srcs, "Module1.bas")
                msg.info("Creating a module to import in '{}'".format(filenpa_workbook))
                module_name=prompt("module name", default="Module1")
                with open(filenpa_module, "w") as f:
                    f.write("Attribute VB_Name = \"{}\"\n".format(module_name))
                    f.write("Public Sub MyModule()\n")
                    f.write("\t\n")
                    f.write("End Sub\n")
                print()
                msg.success("Module created '{}'".format(filenpa_module))
            elif user_input == "export":
                msg.info("Export modules from '{}'".format(filenpa_workbook))
                export(
                    active_hwnd,
                    filenpa_workbook,
                    direpa_srcs,
                    overwrite=True,
                )        
    else:
        msg.warning("Not found '{}'. Creating it with export".format(direpa_srcs))
        export(
            active_hwnd,
            filenpa_workbook,
            direpa_srcs,
            overwrite=False,
        )    
