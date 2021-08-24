#!/usr/bin/env python3
from pprint import pprint
import os
import sys

def get_dy_enum(enum_name):
    dy_enums={
        "VBComponents": {
            1: {
                "name": "vbext_ct_StdModule",
                "code": 1,
                "description": "Standard module",
                "extension": ".bas",
                "allowed": True
            },
            2: {
                "name": "vbext_ct_ClassModule",
                "code": 2,
                "description": "Class module",
                "extension": ".cls",
                "allowed": True
            },
            3: {
                "name": "vbext_ct_MSForm",
                "code": 3,
                "description": "Microsoft Form",
                "extension": ".frm",
                "allowed": True
            },
            11: {
                "name": "vbext_ct_ActiveXDesigner",
                "code": 11,
                "description": "ActiveX Designer",
                "extension": None,
                "allowed": False
            },
            100: {
                "name": "vbext_ct_Document",
                "code": 100,
                "description": "Document Module",
                "extension": ".cls",
                "allowed": False
            }
        }
    }

    return dy_enums[enum_name]