#!/usr/bin/env python3

if __name__ == "__main__":
    import importlib
    import os
    import sys
    direpa_script=os.path.dirname(os.path.realpath(__file__))
    direpa_script_parent=os.path.dirname(direpa_script)
    module_name=os.path.basename(direpa_script)
    sys.path.insert(0, direpa_script_parent)
    pkg = importlib.import_module(module_name)
    del sys.path[0]

    def seed(pkg_major, direpas_configuration=dict(), fun_auto_migrate=None):
        fun_auto_migrate()
    etconf=pkg.Etconf(enable_dev_conf=False, tree=dict( files=dict({ "settings.json": dict() })), seed=seed)
    args, dy_app=pkg.Options(
        direpa_configuration=etconf.direpa_configuration,
        examples=r"""
            main.py --export --workbook ..\test\generate-interactions-report.xlsm
            main.py --import --workbook "..\test\generate-interactions report.xlsm" --overwrite
            main.py --import --workbook "..\test\generate-interactions report.xlsm" --overwrite --reset-cache
            main.py --workbook "..\test\generate-interactions report.xlsm" --macro RefreshFiles
            main.py --workbook "..\test\generate-interactions report.xlsm" --import --macro RefreshFiles --params myparam --overwrite
            main.py --workbook "..\test\generate-interactions report.xlsm" --macro RefreshFiles --params myparam --import --overwrite --immediate
        """, 
        filenpa_app="gpm.json", 
        filenpa_args="config/options.json"
    ).get_argsns_dy_app()

    if args.pkill.here:
        os.system("TASKKILL /F /IM excel.exe")

    if args.no_recovery.here:
        direpa_recovery=os.path.join(os.path.expanduser("~"), r"AppData\Roaming\Microsoft\Excel")
        if os.path.exists(direpa_recovery):
             os.system('rmdir /S /Q "{}"'.format(direpa_recovery))
        os.makedirs(direpa_recovery, exist_ok=True)

    direpa_tmp=os.path.join(os.path.expanduser("~"), "fty", "tmp", dy_app["name"][0], dy_app["name"])
    os.makedirs(direpa_tmp, exist_ok=True)
    filenpa_cache=os.path.join(direpa_tmp, "cache.json")

    if args.export.here or args._import.here or args.macro.here:
        direpa_srcs=args.srcs.value
        filenpa_workbook=args.workbook.value

        if direpa_srcs is None:
            filer_workbook, ext=os.path.splitext(filenpa_workbook)
            direpa_srcs=os.path.join(os.path.dirname(filenpa_workbook), os.path.basename(filer_workbook).replace(" ", "_").lower())

        if args.export.here:
            pkg.export(
                filenpa_workbook=filenpa_workbook,
                direpa_srcs=direpa_srcs,
                overwrite=args.overwrite.here,
            )
            sys.exit(0)
        elif args._import.here:
            # keep_open=False
            # if args.macro.here:
            #     keep_open=True
            pkg._import(
                filenpa_cache,
                filenpa_workbook=filenpa_workbook,
                direpa_srcs=direpa_srcs,
                overwrite=args.overwrite.here,
                reset_cache=args.reset_cache.here,
                # keep_open=keep_open,
            )

        if args.macro.here:
            pkg.macro(
                filenpa_workbook=filenpa_workbook,
                macro_name=args.macro.value,
                immediate=args.immediate.here,
                params=args.params.values,
            )
