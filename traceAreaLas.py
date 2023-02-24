# coding: utf-8
# Python 3.11.1
'''
    traceAreaLas 1.1

    Trace in Autocad the extent of the .las files.

    No install

    Requirements: Autocad
    External Modules: pyautocad, laspy

    Only for Windows
    Tested on Windows 10 and Autocad 2015, 2022

    :No copyright: (!) 2023 by Frédéric Coulon.
    :No license: Do with it what you want.
'''
from pyautocad import Autocad, APoint, aDouble
import laspy
from tkinter import Tk, filedialog, messagebox
import win32gui
import win32con

# Connect to Autocad
acad = Autocad(create_if_not_exists=True)

def gethandlewin(win):
    # get the HANDLE of open applications, argument is name of application
    rslts = []
    ret = None
    win32gui.EnumWindows(lambda h, liste: liste.append(h), rslts)
    for handle in rslts:
        res = win32gui.GetWindowText(handle)
        if win in res:
            ret = handle
    return ret

def traceAreaLas():
    doc = acad.ActiveDocument
    #  Start message
    acad.prompt('traceAreaLas connected\n')
    # Files explorer
    root = Tk()
    # Hides the root window
    root.withdraw()
    file_path = filedialog.askopenfilenames(initialdir='c:/',
                                        title='Select LAS files',
                                        filetypes=[('LAS files', '*.las')])
    # Show Autocad
    hwin = gethandlewin('Autodesk AutoCAD')
    if hwin:
        win32gui.ShowWindow(hwin, win32con.SW_SHOW)
        win32gui.SetForegroundWindow(hwin)
    # If files were found
    if file_path:
        acad.prompt('In Progress ...\n')
        # Counters
        cpt = 0
        cptI = 0
        # Insert units in Meter
        doc.SetVariable('insunits', 6)
        # UCS Controle
        ucs = doc.GetVariable('worlducs')
        if ucs != 1:
            doc.SendCommand('_ucs\n\n')
        # Loop over files
        for f in file_path:
            # Extract file path name
            nf = '/'.join(f.split('/')[-1:])[:-4]
            # Las reading
            las = laspy.read(f)
            # number of points
            ptCount = las.header.point_count
            if ptCount > 1000:
                # Extract coordinates in the las file (bounding box)
                x = las.X
                y = las.Y
                xscale = las.header.scales[0]
                xoffset = las.header.offsets[0]
                yscale = las.header.scales[1]
                yoffset = las.header.offsets[1]
                x1 = (min(x) * xscale) + xoffset
                y1 = (max(y) * yscale) + yoffset
                x2 = (max(x) * xscale) + xoffset
                y2 = (min(y) * yscale) + yoffset
                xm = int(x1/100) * 100
                ym = int(y1/100) * 100 + 100
                x1t = x1 - xm
                y1t = y1 - ym
                x2t = x2 - xm
                y2t = y2 - ym
                # Create a block in collection
                b1 = acad.doc.Blocks.Add(APoint(0, 0, 0), nf)
                # Add objects in the block
                b1.AddPolyline(aDouble(0, 0, 0,
                                        100, 0, 0,
                                        100, -100, 0,
                                        0, -100, 0,
                                        0, 0, 0))
                b1.AddPolyline(aDouble(x1t, y1t, 0,
                                        x2t, y1t, 0,
                                        x2t, y2t, 0,
                                        x1t, y2t, 0,
                                        x1t, y1t, 0,
                                        )).Color = 1
                b1.AddAttribute(5, 0,
                                "Coord.",
                                APoint(5, -40, 0),
                                "X_Y",
                                nf)
                b1.AddAttribute(5, 0,
                                "Counter_Point",
                                APoint(5, -50, 0),
                                "C_P",
                                str(ptCount)+" points")
                b1.AddAttribute(5, 0,
                                "Date",
                                APoint(5, -60, 0),
                                "DATE",
                                str(las.header.creation_date))
                
                # block insertion
                acad.model.InsertBlock(APoint(xm, ym), nf, 1, 1, 1, 0)
                cpt += 1
            else: # Add 1 to the ignored
                cptI += 1
        # eventual UCS recovery
        if ucs != 1:
            doc.SendCommand('_ucs\n_p\n')
        acad.app.ZoomExtents()
        # Final message
        acad.prompt(f'\n{str(cpt)} .las processed\n{str(cptI)} .las ignored\n')
        # acad.prompt(str(cptI) + ' .las ignored\n')
    else:
        messagebox.showerror(title='Error',
                    message='No file, or abandonment',)
# Autocad check
if acad:
    traceAreaLas()
else:
    messagebox.showerror(title='Error',
                    message='Autocad must be installed, or unknown error',)

# cd c:/Data/Python/traceAreaLas/
# pyinstaller --noconsole --onefile traceAreaLas.py