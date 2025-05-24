#!/usr/bin/env python3

if __name__ == "__main__":
    import importlib
    import os
    import sys
    import win32gui
    direpa_script=os.path.dirname(os.path.realpath(__file__))
    direpa_script_parent=os.path.dirname(direpa_script)
    module_name=os.path.basename(direpa_script)
    sys.path.insert(0, direpa_script_parent)
    pkg = importlib.import_module(module_name)
    del sys.path[0]

    args=pkg.Nargs(
        metadata=dict(
            executable="vba-sync",
        ),
        options_file="config/options.yaml", 
    ).get_args()

    if args.pkill._here:
        os.system("TASKKILL /F /IM excel.exe")

    if args.no_recovery._here:
        direpa_recovery=os.path.join(os.path.expanduser("~"), r"AppData\Roaming\Microsoft\Excel")
        if os.path.exists(direpa_recovery):
             os.system('rmdir /S /Q "{}"'.format(direpa_recovery))
        os.makedirs(direpa_recovery, exist_ok=True)

    if args.export._here or args._["import"]._here or args.macro._here:
        active_hwnd=win32gui.GetForegroundWindow()
        direpa_srcs:str|None=None
        filenpa_workbook:str|None=args.workbook._value

        if filenpa_workbook is None:
            raise Exception("--workbook argument is required.")
        
        if args.export._here:
            direpa_srcs=args.export.srcs._value
        elif args._["import"]._here:
            direpa_srcs=args._["import"].srcs._value
        elif args.macro._here:
            direpa_srcs=args.macro.srcs._value

        if direpa_srcs is None:
            filer_workbook, ext=os.path.splitext(filenpa_workbook)
            direpa_srcs=os.path.join(os.path.dirname(filenpa_workbook), os.path.basename(filer_workbook).replace(" ", "_").lower())

        os.makedirs(direpa_srcs, exist_ok=True)
        filenpa_cache=os.path.join(direpa_srcs, "vba-sync-cache.json")

        if args.export._here:
            pkg.export(
                active_hwnd=active_hwnd,
                filenpa_workbook=filenpa_workbook,
                direpa_srcs=direpa_srcs,
                overwrite=args.export.overwrite._here,
            )
            sys.exit(0)
        elif args._["import"]._here:
            pkg._import(
                active_hwnd=active_hwnd,
                filenpa_cache=filenpa_cache,
                filenpa_workbook=filenpa_workbook,
                direpa_srcs=direpa_srcs,
                overwrite=args._["import"].overwrite._here,
                reset_cache=args._["import"].reset_cache._here,
            )

        if args.macro._here:
            pkg.macro(
                active_hwnd=active_hwnd,
                clear=args.macro.immediate.clear._here,
                filenpa_workbook=filenpa_workbook,
                macro_name=args.macro.macro._value,
                immediate=args.macro.immediate._here,
                params=args.macro.params._values,
                reset_macro=args.macro.reset_macro._here,
                reset_macro_seconds=args.macro.reset_macro._value,
            )
