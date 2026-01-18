import ttkbootstrap as ttk
from auto_fill_defects import App

def main(): 
    # Create the main window using ttkbootstrap
    root = ttk.Window(themename="cosmo")
    app = App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
