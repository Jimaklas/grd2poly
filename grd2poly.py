import win32com.client
from input import GRD_FILE

# # Generate modules of necessary typelibs (AutoCAD Civil 3D 2008)
# comtypes.client.GetModule("C:\\Program Files\\Common Files\\Autodesk Shared\\acax17enu.tlb")
# comtypes.client.GetModule("C:\\Program Files\\AutoCAD Civil 3D 2008\\AecXBase.tlb")
# comtypes.client.GetModule("C:\\Program Files\\AutoCAD Civil 3D 2008\\AecXUIBase.tlb")
# comtypes.client.GetModule("C:\\Program Files\\AutoCAD Civil 3D 2008\\Civil\\AeccXLand.tlb")
# comtypes.client.GetModule("C:\\Program Files\\AutoCAD Civil 3D 2008\\Civil\\AeccXUiLand.tlb")
# raise SystemExit

# Get running instance of the AutoCAD application
acadApp = win32com.client.Dispatch("AutoCAD.Application")
aeccApp = acadApp.GetInterfaceObject("AeccXUiLand.AeccApplication.5.0")

# Document object
doc = aeccApp.ActiveDocument
alignment, point_clicked = doc.Utility.GetEntity(None, None, Prompt="Select an alignment:")

command = "pl "
f = open(GRD_FILE, "r")
line = f.readline()
while 1:
    try:
        line = f.readline()
        section, station = line.strip().split()
        station = float(station)
        line = f.readline()
        while line[0] != "*":
            offset, h = line.strip().split()
            offset = float(offset)
            h = float(h)
            # draw the next polyline vertex
            print "Point at station %s (section %s) - offset %s" % (station, section, offset)
            x, y = alignment.PointLocation(station, offset)
            command = command + "%s,%s " % (x, y)
            line = f.readline()
    except ValueError:  # raised when trying to read past EOF (why not IOError? - need to think on it)
        break

doc.SendCommand(command + " ")
f.close()
# x, y = alignment.PointLocation(0.0, 10.0)
# print x, y
