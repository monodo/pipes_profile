import arcpy
import pythonaddins

class CoupeEle(object):
    """Implementation for coupes_geonis_addin.tool (Tool)"""
    def __init__(self):
        self.enabled = True
        self.cursor = 3
        self.shape = 'Rectangle'
    def onMouseDown(self, x, y, button, shift):
        pass
    def onMouseDownMap(self, x, y, button, shift):
        pass
    def onMouseUp(self, x, y, button, shift):
        pass
    def onMouseUpMap(self, x, y, button, shift):
        pass
    def onMouseMove(self, x, y, button, shift):
        pass
    def onMouseMoveMap(self, x, y, button, shift):
        pass
    def onDblClick(self):
        pass
    def onKeyDown(self, keycode, shift):
        pass
    def onKeyUp(self, keycode, shift):
        pass
    def deactivate(self):
        pass
    def onCircle(self, circle_geometry):
        pass
    def onLine(self, line_geometry):
        pass


    def onRectangle(self, rectangle_geometry):
        """Occurs when the rectangle is drawn and the mouse button is released.
        The rectangle is a extent object."""

        extent = rectangle_geometry
        # Create a fishnet with 10 rows and 10 columns.
        if arcpy.Exists(r'in_memory\fishnet'):
            arcpy.Delete_management(r'in_memory\fishnet')
        fishnet = arcpy.CreateFishnet_management(r'in_memory\fishnet',
                                '%f %f' %(extent.XMin, extent.YMin),
                                '%f %f' %(extent.XMin, extent.YMax),
                                0, 0, 10, 10,
                                '%f %f' %(extent.XMax, extent.YMax),'NO_LABELS',
                                '%f %f %f %f' %(extent.XMin, extent.YMin, extent.XMax, extent.YMax), 'POLYGON')
        arcpy.RefreshActiveView()
        return fishnet