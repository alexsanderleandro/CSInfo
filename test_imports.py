print("Testing imports...")
try:
    import csinfo
    print("- csinfo imported successfully")
    print(f"- csinfo path: {csinfo.__file__}")
    
    # Test submodules
    import csinfo._impl
    import csinfo._core
    import csinfo.network_discovery
    print("- All csinfo submodules imported successfully")
    
    # Test other required modules
    import win32api
    import win32con
    import win32security
    import pythoncom
    import pywintypes
    import wmi
    import psutil
    print("- All required modules imported successfully")
    
except Exception as e:
    print(f"Error: {str(e)}")
    import traceback
    traceback.print_exc()

input("Press Enter to exit...")
