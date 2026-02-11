import argparse
import json
import os
import pandas as pd
class LoadJSONAction(argparse.Action):
    def __call__(self, parser, namespace, values, option_string=None):
        try:
            if values is None:
                setattr(namespace, self.dest, None)
                return
            with open(values, 'r') as f:
                setattr(namespace, self.dest, json.load(f))
        except FileNotFoundError:
            parser.error(f"File not found: {values}")
        except json.JSONDecodeError as e:
            parser.error(f"Error decoding JSON file {values}: {e}")

class ValidateDirectoryAction(argparse.Action):
    def __call__(self, parser, namespace, values, option_string=None):
        if not os.path.exists(values):
            print(f"Output directory {values} does not exist. Creating it.")
            os.makedirs(values)
        setattr(namespace, self.dest, values)

class LoadExcelAction(argparse.Action):
    def __call__(self, parser, namespace, values, option_string=None):
        try:
            # Read the Excel file
            path = str(values)
            lower = path.lower()
            if lower.endswith('.xls'):
                df = pd.read_excel(path, engine='xlrd')
            else:
                df = pd.read_excel(path, engine='openpyxl')
            setattr(namespace, self.dest, df)
        except FileNotFoundError:
            parser.error(f"File not found: {values}")
        except Exception as e:
            parser.error(f"Error reading Excel file {values}: {e}")
