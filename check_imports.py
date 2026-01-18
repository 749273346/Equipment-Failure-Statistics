
try:
    import ttkbootstrap as ttk
    print("ttkbootstrap imported successfully")
    try:
        from ttkbootstrap.constants import PRIMARY, SUCCESS, SECONDARY
        print("ttkbootstrap.constants imported PRIMARY, SUCCESS, SECONDARY")
    except ImportError as e:
        print(f"Failed to import constants: {e}")
        import ttkbootstrap.constants
        print(f"ttkbootstrap.constants content: {dir(ttkbootstrap.constants)}")

    try:
        from ttkbootstrap.widgets.scrolled import ScrolledText
        print("ttkbootstrap.widgets.scrolled imported ScrolledText")
    except ImportError as e:
        print(f"Failed to import ScrolledText from widgets: {e}")

except ImportError as e:
    print(f"Failed to import ttkbootstrap: {e}")

try:
    import win32com.client
    print("win32com.client imported successfully")
except ImportError as e:
    print(f"Failed to import win32com.client: {e}")

try:
    import matplotlib.pyplot as plt
    print("matplotlib.pyplot imported successfully")
except ImportError as e:
    print(f"Failed to import matplotlib: {e}")
