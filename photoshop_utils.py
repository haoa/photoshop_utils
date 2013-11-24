# -*- coding: utf-8 -*-

__author__ = 'Alexander Haug'

import win32com.client

def walk_artlayers(layerOrCollection, caller=None, function=None):
    if layerOrCollection.TypeName == "Layers":
        for index in range(0, len(layerOrCollection)):
            layer = layerOrCollection[index]
            walk_artlayers(layer, caller, function)
    elif layerOrCollection.TypeName == "LayerSet":
        for layer in layerOrCollection.Layers:
            layer.Visible = False
            walk_artlayers(layer, caller, function)
    elif layerOrCollection.TypeName == "ArtLayer":
        try:
            layerOrCollection.Visible = False
            #if caller and function:
            getattr(caller, function)(layerOrCollection)
            #else:
            #    print "Silent error: caller is " + caller + " and function is " + function
        except Exception as e: # Will never happen as of the if else statement above. TODO: Decision necessary whether to raise error or print silent one.
            print "Original exception: " + e.message
            #raise Exception(
            #    "caller " + str(caller) + " does not contain function or functionname " + function + " was misspelled.")


def measure(bounds):
    return Bounds(bounds)


''' Begin custom error classes '''


class ReadOnlyError(Exception):
    def __init__(self, message):
        self.message = message

    def __str__(self):
        return repr(self.message)


'''constants'''


class NewDocumentMode():
    GRAYSCALE = 1
    RGB = 2
    CMYK = 3
    LAB = 4
    BITMAP = 5


class DocumentFill():
    BACKGROUNDCOLOR = 1
    TRANSPARENT = 2
    WHITE = 3


class BitsPerChannelType():
    EIGHT = 8
    ONE = 1
    SIXTEEN = 16
    THIRTYTWO = 32


class ColorProfileType():
    NO = 1
    WORKING = 2
    CUSTOM = 3


class ElementPlacement():
    INSIDE = 0
    PLACEATBEGINNING = 1
    PLACEATEND = 2
    PLACEBEFORE = 3
    PLACEAFTER = 4


class AnchorPosition():
    TOPLEFT = 1
    TOPCENTER = 2
    TOPRIGHT = 3
    MIDDLELEFT = 4
    MIDDLECENTER = 5
    MIDDLERIGHT = 6
    BOTTOMLEFT = 7
    BOTTOMCENTER = 8
    BOTTOMRIGHT = 9

class Units():
    PIXELS = 1
    INCHES = 2
    CM = 3
    MM = 4
    POINTS = 5
    PICAS = 6
    PERCENT = 7

class ExtensionType():
    NONE = 1
    LOWERCASE = 2
    UPPERCASE = 3

class DialogModes():
    ALL = 1
    ERROR = 2
    NONE = 3

class SaveOptions():
    SAVECHANGES = 1
    DONOTSAVECHANGES = 2
    PROMPTTOSAVECHANGES = 3

class PhotoshopSaveOptions():
    def __init__(self):
        self = win32com.client.Dispatch('Photoshop.PhotoshopSaveOptions')

    AlphaChannels = True

    Layers = True

    Annotations = True

    EmbedColorProfile = True

    SpotColors = True

'''collections'''


class Collection(list):
    def __init__(self, collection):
        self._collection = collection
        for item in collection:
            self.append(item)

    def __get_Length(self):
        return self.Length

    Length = property(__get_Length)

    def __get_Parent(self):
        return self.Parent

    Parent = property(__get_Parent)

    def __get_TypeName(self):
        return self._collection.TypeName

    TypeName = property(__get_TypeName)

    def __getitem__(self, index):
        return self._collection[index]

    def GetByName(self, name):
        return self._collection.GetByName(name)


class EditableCollection(Collection):
    def __init__(self, collection):
        super(EditableCollection, self).__init__(collection)

    def RemoveAll(self):
        self._collection.RemoveAll()

    def Add(self, name, comment, appearance, position, visibility):
        self._collection.Add(name, comment, appearance, position, visibility)

    def __setitem__(self, key, value):
        self._collection[key] = value


class Layers(EditableCollection):
    def __init__(self, layers):
        super(Layers, self).__init__(layers)


class LayerComps(EditableCollection):
    def __init__(self, layerComps):
        super(LayerComps, self).__init__(layerComps)


class ArtLayers(EditableCollection):
    def __init__(self, artLayers):
        super(ArtLayers, self).__init__(artLayers)

    def Add(self):
        return ArtLayer(self._collection.Add())


class HistoryStates(Collection):
    def __init__(self, layers):
        super(HistoryStates, self).__init__(layers)


class Documents(Collection):
    def __init__(self, collection):
        super(Documents, self).__init__(collection)

    def Add(self, width, height, resolution, name, mode, initialFill, pixelAspectRatio, bitsPerChannel,
            colorProfileName):
        self._collection.Add(width, height, resolution, name, mode, initialFill, pixelAspectRatio, bitsPerChannel,
                             colorProfileName)


''' Begin wrapper classes '''


class Bounds():
    def __init__(self, bounds):
        self._bounds = bounds

    def __get_X(self):
        return self._bounds[0]

    def __set_X(self, newX):
        raise ReadOnlyError("Bounds error: X property is read-only. Please use Translate function on Layer.")

    X = property(__get_X, __set_X)

    def __get_Y(self):
        return self._bounds[1]

    def __set_Y(self, newY):
        raise ReadOnlyError("Bounds error: Y property is read-only. Please use Translate function on Layer.")

    Y = property(__get_Y, __set_Y)

    def __get_Width(self):
        return self._bounds[2] - self._bounds[0]

    def __set_Width(self, newWidth):
        raise ReadOnlyError("Bounds error: Width property is read-only. Please use Resize function on Layer.")

    Width = property(__get_Width, __set_Width)

    def __get_Height(self):
        return self._bounds[3] - self._bounds[1]

    def __set_Height(self, newHeight):
        raise ReadOnlyError("Bounds error: Height property is read-only. Please use Resize function on Layer.")

    Height = property(__get_Height, __set_Height)


class ArtLayer(object):
    def __init__(self, artLayer):
        self._ArtLayer = artLayer

    def __get_name(self):
        return self._ArtLayer.Name

    def __set_name(self, newName):
        if newName is None:
            return False
        self._ArtLayer.Name = newName

    Name = property(__get_name, __set_name)

    def __get_Visible(self):
        return self._ArtLayer.Visible

    def __set_Visible(self, newVisible):
        if newVisible is None:
            return False
        self._ArtLayer.Visible = newVisible

    Visible = property(__get_Visible, __set_Visible)

    def __get_X(self):
        return self._ArtLayer.Bounds[0]

    X = property(__get_X)

    def __get_Y(self):
        return self._ArtLayer.Bounds[1]

    Y = property(__get_Y)

    def __get_Width(self):
        return int(self._ArtLayer.Bounds[2]) - int(self._ArtLayer.Bounds[0])

    Width = property(__get_Width)

    def __get_Height(self):
        return int(self._ArtLayer.Bounds[3]) - int(self._ArtLayer.Bounds[1])

    Height = property(__get_Height)

    def __get_TypeName(self):
        return self._ArtLayer.TypeName

    TypeName = property(__get_TypeName)

    def __get_Bounds(self):
            return self._ArtLayer.Bounds

    def __set_Bounds(self, newBounds):
        if newBounds is None:
            return False
        self._ArtLayer.Bounds = newBounds

    Bounds = property(__get_Bounds, __set_Bounds)

    '''functions'''

    def Translate(self, deltaX, deltaY):
        self._ArtLayer.Translate(deltaX, deltaY)

    def Resize(self, horizontal, vertical, anchorPosition=AnchorPosition.MIDDLECENTER):
        self._ArtLayer.Resize(horizontal, vertical, anchorPosition)

    def Duplicate(self, relativeObject=None, insertionLocation=None):
        self._ArtLayer.Duplicate(relativeObject, insertionLocation)


class LayerSet(object):
    def __init__(self, layerSet):
        self._LayerSet = layerSet

    def __get_name(self):
        return self._LayerSet.Name

    def __set_name(self, newName):
        if newName is None:
            return False
        self._LayerSet.Name = newName

    Name = property(__get_name, __set_name)


class ActiveLayer(ArtLayer):
    def __init__(self, layer):
        super(ActiveLayer, self).__init__(layer)


class Selection(object):
    def __init__(self, selection):
        self._Selection = selection

    def SelectAll(self):
        self._Selection.SelectAll()

    def Copy(self):
        self._Selection.Copy()


class LayerComp(object):
    def __init__(self, layerComp):
        self._layerComp = layerComp

    def __get_Name(self):
        return self._layerComp.Name

    def __set_Name(self, newName):
        if newName is None:
            return False
        self._layerComp.Name = newName

    Name = property(__get_Name, __set_Name)

    def Apply(self):
        self._layerComp.Apply()


class HistoryState(object):
    def __init__(self, historyState):
        self._historyState = historyState

    def __get_Name(self):
        return self.Name

    Name = property(__get_Name)

    def __get_TypeName(self):
        return self.TypeName

    Typename = property(__get_TypeName)

    def __get_Snapshot(self):
        return self.Snapshot

    Snapshot = property(__get_Snapshot)

    def __get_Parent(self):
        return Document(self.Parent)

    Parent = property(__get_Parent)


class Document(object):
    def __init__(self, document):
        self._Document = document
        #self.Layers = document.Layers
        #self.ActiveLayer = document.ActiveLayer
        #self.ActiveHistoryState = document.ActiveHistoryState
        #self.Selection = document.Selection

    def __get_activeLayer(self):
        activeLayer = ArtLayer(self._Document.ActiveLayer)
        if activeLayer.TypeName == "ArtLayer":
            return ArtLayer(activeLayer)
        elif activeLayer.TypeName == "LayerSet":
            return LayerSet(activeLayer)

    def __set_activeLayer(self, activeLayer):
        if activeLayer is None:
            return False
        self._Document.ActiveLayer = ArtLayer(activeLayer)._ArtLayer

    ActiveLayer = property(__get_activeLayer, __set_activeLayer)

    def __get_Layers(self):
        return Layers(self._Document.Layers)

    Layers = property(__get_Layers)

    def __get_ArtLayers(self):
        return ArtLayers(self._Document.ArtLayers)

    ArtLayers = property(__get_ArtLayers)

    def __get_selection(self):
        return Selection(self._Document.Selection)

    def __set_selection(self, selection):
        if selection is None:
            return False
        self._Document.Selection = Selection(selection)

    Selection = property(__get_selection, __set_selection)

    def __get_Name(self):
        return self._Document.Name

    def __set_Name(self, newName):
        if newName is None:
            return False
        self._Document.Name = newName

    Name = property(__get_Name, __set_Name)

    def __get_LayerComps(self):
        return self._Document.LayerComps

    def __set_LayerComps(self, newLayerComps):
        if newLayerComps is None:
            return False
        self._Document.LayerComps = newLayerComps

    LayerComps = property(__get_LayerComps, __set_LayerComps)

    def __get_HistoryStates(self):
        return HistoryStates(self._Document.HistoryStates)

    HistoryStates = property(__get_HistoryStates)

    def __get_ActiveHistoryState(self):
        return HistoryState(self._Document.ActiveHistoryState)

    def __set_ActiveHistoryState(self, newActiveHistoryState):
        if newActiveHistoryState is None:
            return False
        self._Document.ActiveHistoryState = newActiveHistoryState._historyState

    ActiveHistoryState = property(__get_ActiveHistoryState, __set_ActiveHistoryState)

    def __get_Width(self):
        return self._Width

    def __set_Width(self, newWidth):
        if newWidth is None:
            return False
        self._Width = newWidth

    Width = property(__get_Width, __set_Width)

    def __get_Height(self):
        return self._Height

    def __set_Height(self, newHeight):
        if newHeight is None:
            return False
        self._Height = newHeight

    Height = property(__get_Height, __set_Height)

    def __get_Resolution(self):
            return self._Document.Resolution

    def __set_Resolution(self, newResolution):
        if newResolution is None:
            return False
        self._Document.Resolution = newResolution

    Resolution = property(__get_Resolution, __set_Resolution)

    '''functions'''

    def Save(self):
        self._Document.Save()

    def SaveAs(self, saveIn, options, asCopy=False, extensionType=ExtensionType.LOWERCASE):
        self._Document.SaveAs(saveIn, options, asCopy, extensionType)

    def Paste(self, intoSelection=False):
        self._Document.Paste(intoSelection)

    def Export(self, exportIn, exportAs, options):
        self._Document.Export(ExportIn=exportIn, ExportAs=exportAs, Options=options)

    def MergeVisibleLayers(self):
        self._Document.MergeVisibleLayers()

    def Close(self, saveOptions):
        self._Document.Close(saveOptions)

    def goBackInHistory(self, steps):
        targetHistoryState = HistoryState(self.HistoryStates[len(self.HistoryStates) - steps - 1])
        self.ActiveHistoryState = targetHistoryState._historyState


class ActiveDocument(Document):
    def __init__(self, document):
        super(ActiveDocument, self).__init__(document)


class Preferences(object):
    def __init__(self, preferences):
        self._preferences = preferences

    def __get_RulerUnits(self):
            return self._preferences.RulerUnits

    def __set_RulerUnits(self, newRulerUnits):
        if newRulerUnits is None:
            return False
        self._preferences.RulerUnits = newRulerUnits

    RulerUnits = property(__get_RulerUnits, __set_RulerUnits)


class ScaleFluidObject(object):
    def __init__(self, layers, psApp):
        self._layers = layers
        self._psApp = psApp

    def to(self, factorOrValue):
        self._factorOrValue = factorOrValue
        return self

    def percent(self):
        self._factorOrValue = int(self._factorOrValue)
        self.resize(Units.PERCENT)

    def resize(self, units):
        starterRulerUnits = self._psApp.Preferences.RulerUnits
        self._psApp.Preferences.RulerUnits = units
        walk_artlayers(self._layers, self, "resizeCall")
        self._psApp.Preferences.RulerUnits = starterRulerUnits

    def resizeCall(self, layer):
        print "Resizing layer " + ArtLayer(layer).Name
        layer = ArtLayer(layer)
        layer.Resize(self._factorOrValue, self._factorOrValue, AnchorPosition.MIDDLECENTER)
        # TODO: Add behaviour (fluid functions)for having different width and height
        # TODO: Add behaviour (fluid function) for setting different Anchorpoint


class Photoshop_App:
    def __init__(self, initializeImmediatly=False):
        if initializeImmediatly:
            self.init()

    def __get_application(self):
        return self._Application

    Application = property(__get_application)

    def __get_activeDocument(self):
        return Document(self._Application.ActiveDocument)

    def __set_activeDocument(self, document):
        if document is None:
            return False
        self._Application.ActiveDocument = Document(document)

    ActiveDocument = property(__get_activeDocument, __set_activeDocument)

    def __get_Documents(self):
        return self._Application.Documents

    Documents = property(__get_Documents)

    def __get_Preferences(self):
            return Preferences(self._Application.Preferences)

    Preferences = property(__get_Preferences)

    def __get_DisplayDialogs(self):
            return self._Application.DisplayDialogs

    def __set_DisplayDialogs(self, newDisplayDialogs):
        if newDisplayDialogs is None:
            return False
        self._Application.DisplayDialogs = newDisplayDialogs

    DisplayDialogs = property(__get_DisplayDialogs, __set_DisplayDialogs)

    # separate init to be able to fire up Photoshop at will
    def init(self):
        self._Application = win32com.client.Dispatch("Photoshop.Application")
        #self._Application.Visible = False
        #self._Application.DisplayDialogs = DialogModes.NONE

    def new(self, width, height, resolution, name, newDocumentMode, documentFill, pixelAspectRatio=1,
            bitsPerChannel=BitsPerChannelType.EIGHT, colorProfileName=""):
        try:
            self._Application.Documents.Add(width, height, resolution, name, newDocumentMode, documentFill, pixelAspectRatio, bitsPerChannel, colorProfileName)
        except Exception as e:
            print "Original Photoshop Exception:\n" + e.message + "\n"
            print "Utils tip: Did you set the Units for the document right? Use Units constant on Photoshop_App.Preferences.Units to do so."
            print "I.e. setting values above 100 for width and height if units are set to percent, will cause Photoshop to throw an exception."

    def Load(self, pathToFile):
        if self._Application is None:
            return False
        self._Application.Load(pathToFile)

    def Open(self, pathToFile):
        if self._Application is None:
            return False
        self._Application.Open(pathToFile)

    def Quit(self):
        self._Application.Quit()

    def Export(self, exportIn, exportAs, options):
        self.ActiveDocument.Export(exportIn, exportAs, options)

    def getDocumentByName(self, name):
        documents = self._Application.Documents
        for document in documents:
            if Document(document).Name == name:
                return Document(document)
        return None

    def setDocumentActiveByName(self, name):
        documents = self._Application.Documents
        for index in range(0, len(documents)):
            document = Document(documents[index])
            if document.Name == name:
                self._Application.ActiveDocument = documents[index]
                break


    def makeAllLayersInvisibleBut(self, layer):
        walk_artlayers(self.ActiveDocument.Layers)
        self.ActiveDocument.ActiveLayer = ArtLayer(layer)

    def resizeAll(self, layers):
        return ScaleFluidObject(layers, self)

    def ExportToPNG(self, layerOrComposition, exportFilePath):
        if type(layerOrComposition) is LayerComp:
            LayerComp(layerOrComposition).Apply()
        elif type(layerOrComposition) is ArtLayer:
            self.makeAllLayersInvisibleBut(layerOrComposition)

        options = win32com.client.Dispatch('Photoshop.ExportOptionsSaveForWeb')
        options.Format = 13   # PNG
        options.PNG8 = False  # Sets it to PNG-24 bit
        self.ActiveDocument.Export(exportFilePath, 2, options)
