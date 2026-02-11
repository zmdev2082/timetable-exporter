import os
import importlib

from .timetable_accessor import TimetableAccessor

for module in os.listdir(os.path.dirname(__file__)):
    if module.endswith('.py') and module != '__init__.py':
        module_name = module[:-3]  # Remove the .py extension
        module_import_path = f'timetable_exporter.user_extensions.{module_name}'
        # Import the module
        imported_module = importlib.import_module(module_import_path)
        
        # Iterate over all attributes in the module
        for name in dir(imported_module):
            # Skip special attributes
            if not name.startswith('__'):
                # Get the attribute from the module
                attr = getattr(imported_module, name)
                
                # Add the attribute to the TimetableAccessor class
                setattr(TimetableAccessor, name, attr)