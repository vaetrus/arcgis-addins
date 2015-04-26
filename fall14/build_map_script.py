
# WARNING: ARCPY DOES NOT CREATE LAYOUT ELEMENTS
# THEREFORE THESE STEPS MUST BE TAKEN BEFORE THE SCRIPT IS RUN
# (1) CREATE A MAP
# (1a) INSERT TWO ADDITIONAL DATAFRAMES (FOR A TOTAL OF THREE) (OPTIONAL)
# (2) UNZIP BASEMAPDATA
# (2a) HAVE COMPLETED THE DIGITIZING BEFOREHAND (OPTIONAL)
# (3) HAVE FIELDTRIPPOINTS (OPTIONAL)
# (3a) HAVE CONVERTED EXCEL TO DBF BEFOREHAND (OPTIONAL)
# (4) ENSURE ALL FILES ARE IN THE SAME DIRECTORY AS THE SCRIPT

from __future__ import print_function
from sys import exit, path

def os_get_cwd():
    ''' None -> str
    '''
    from os import getcwd
    return getcwd()
    
def os_find_file(filename, filetype):
    ''' str, str -> bool
    Returns whether file is in workspace.
    '''
    from os.path import isfile
    return isfile(filename + filetype)

if __name__ != '__main__':
    sys_exit()
    
import logging
logger = logging.getLogger(__name__)
fr = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s')
sh = logging.StreamHandler()
sh.setFormatter(fr)
logger.addHandler(sh)

if getcwd() == r'H:\Desktop\assets':
    logger.setLevel(logging.DEBUG)
else:
    logger.setLevel(35)

try:
    logger.info('..importing arcpy..')
    import arcpy
except ImportError:
    logger.warn('..import not successful..')
    arcpy = ["C:/Program Files (x86)/ArcGIS/Desktop10.2/arcpy",
             "C:/Program Files (x86)/ArcGIS/Desktop10.2/arcpy/arcpy",
             "C:/Program Files (x86)/ArcGIS/Desktop10.2/bin",
             "C:/Program Files (x86)/ArcGIS/Desktop10.2/scripts",
             "C:/Program Files (x86)/ArcGIS/Desktop10.2/ArcToolbox",
             "C:/Program Files (x86)/ArcGIS/Desktop10.2/ArcToolbox/Scripts",
             "C:/Python27/ArcGIS10.2/Lib/site-packages"]
    logger.info('..appending to sys.path..')
    for i in arcpy:
        path.append(i)
    logger.info('..importing arcpy again..')
    try:
        import arcpy
    except:
        logger.info('..import not successful..')
        exit("Could not import arcpy 10.2.")
logger.info('..import successful..')

# Set workspace
arcpy.env.workspace = os_get_cwd()

# Check for map
cwd_files = arcpy.ListFiles()
for i in cwd_files:
    if i.endswith(".mxd"):
        mxd = arcpy.mapping.MapDocument(i)
temp_workspace = "/".join(arcpy.env.workspace.split("/")[:-1])
for i in temp_workspace:
    if i.endswith(".mxd"):
        mxd = arcpy.mapping.MapDocument(i)
try:
    mxd
except NameError:
    logger.error('..no map found..')
    exit("No map was found. Please create a new map and save it in the same folder.")

# Names to ensure correct files
names = [u'boundary.shp', u'contours.shp', u'places.shp', u'roads.shp',
    u'streams.shp', u'vegetation.shp', u'water.shp']

# Get reference to BasemapData.zip
basemapdata = [i for i in arcpy.ListFeatureClasses() if i in names]
'''
>>> basemapdata
[u'boundary.shp', u'contours.shp', u'places.shp', u'roads.shp',
    u'streams.shp', u'vegetation.shp', u'water.shp']
'''
if len(basemapdata) != 7:
    logger.error('..not enough starting files..')
    exit("There are files missing. Please extract basemapdata.zip fully.")

# Get first dataframe
first_frame = arcpy.mapping.ListDataFrames(mxd)[0]

# Name the frame for easier searching later
first_frame.name = "Main Map"

# Add all data to dataframe
# Note: this auto sorts as points > polylines > polygons > greater area polygons
for i in basemapdata:
    arcpy.mapping.AddLayer(first_frame, arcpy.mapping.Layer(i))

"""
>>> arcpy.mapping.ListLayers(first_frame)
[<map layer u'places'>, <map layer u'streams'>, <map layer u'roads'>, <map layer u'contours'>,
    <map layer u'water'>, <map layer u'vegetation'>, <map layer u'boundary'>]
"""

# Check for fieldtrippoints.xlsx
table_name = "fieldtrippoints"
logger.info('..checking for fieldtrippoints..')
# If excel version found
if "fieldtrippoints.dbf" not in arcpy.ListTables() and os_find_file(table_name, ".xlsx"):
    logger.info('..dbf not found, excel found, converting..')
    # Convert xlsx to dbf
    in_table = "fieldtrippoints.xlsx"
    out_table = "fieldtrippoints.dbf"
    try:
        logger.info('..converting xlsx to dbf..')
        arcpy.ExcelToTable_conversion(in_table, out_table)
    except ImportError:
        pass
# If no excel version found, table will be created
elif "fieldtrippoints.dbf" not in arcpy.ListTables() and False:
    logger.info('..fieldtrippoints not found, creating new table..')
    arcpy.CreateTable_management(arcpy.env.workspace, "fieldtrippoints.dbf")
    table = [i for i in arcpy.ListTables() if 'fieldtrippoints' in i][0]
    table = arcpy.mapping.TableView(table)
    fields = arcpy.ListFields(table.dataSource.split("\\")[-1])
    # arcpy.AlterField_management(table, [i for i in fields if i.name == 'Field1'][0], "Name")
    # arcpy.DeleteField_management(table, fields[1])
    ## Can't seem to rename or delete fields.
    arcpy.AddField_management(table, "Name", "TEXT")
    arcpy.AddField_management(table, "Northing", "FLOAT")
    arcpy.AddField_management(table, "Easting", "FLOAT")

    rows = arcpy.InsertCursor(table)
    station = []
    northing = []
    easting = []
    for i in range(7):
        row = rows.newRow()
        row.setValue("Name", station[i])
        row.setValue("Northing", northing[i])
        row.setValue("Easting", easting[i])
        rows.insertRow(row)

if "fieldtrippoints.dbf" in arcpy.ListTables():
    logger.info('..working with fieldtrippoints..')
    # Reference (new) table
    fieldtripdata = [i for i in arcpy.ListTables() if 'fieldtrippoints' in i][0]

    # Get table view to map
    fieldtrip_view = arcpy.mapping.TableView(fieldtripdata)

    # Add table to TOC
    arcpy.mapping.AddTableView(first_frame, fieldtrip_view)

    # Reference the fields, for XY display
    station_fields = arcpy.ListFields(fieldtripdata)
    '''
    >>> station_fields
    [u'OID', u'Name', u'Northing', u'Easting']
    '''

    # Check basemap projection
    if first_frame.spatialReference.name == "":
        lyr_pcs = arcpy.Describe(arcpy.ListFeatureClasses()[1]).spatialReference
    else:
        lyr_pcs = first_frame.spatialReference
    '''
    >>> lyr_pcs.name
    u'NAD_1983_UTM_Zone_17N'
    '''

    # Make first frame active (so that XY Layer will be created there)
    mxd.activeView = first_frame.name

    # Display XY Data
    logger.info('..creating xy layer from table..')
    arcpy.MakeXYEventLayer_management(fieldtripdata, station_fields[3].name, station_fields[2].name, "stations", lyr_pcs)

    # Save new layer and import
    logger.info('..creating stations shp..')
    arcpy.FeatureClassToShapefile_conversion("stations", arcpy.env.workspace)
    logger.info('..exporting xy layer to shp..')
    arcpy.mapping.AddLayer(first_frame, arcpy.mapping.Layer("stations.shp"))
else:
    logger.warn('..no fieldtrippoints..')

# Change layer names
for i in arcpy.mapping.ListLayers(first_frame):
    if i.name == 'places' or i.name == 'boundary':
        logger.info('..turning off visibility for {0}..'.format(i.name))
        i.visible = False
        continue
    if i.name == 'contours':
        i.name = 'Topography'
    elif i.name == 'roads':
        i.name = 'Transportation'
    else:
        i.name = i.name.capitalize()

# Zoom to full extent, change scale, reset orientation
first_frame.zoomToSelectedFeatures()
# first_frame.scale = 350000
# first_frame.elementPositionX, first_frame.elementPositionY = 1, 1
# first_frame.elementHeight, first_frame.elementWidth = 25, 20

logger.info('..completely done with first frame..')

#####
### Working on Second Dataframe
# Create reference to new frame
if len(arcpy.mapping.ListDataFrames(mxd)) > 1:
    second_frame = [i for i in arcpy.mapping.ListDataFrames(mxd) if i.name != first_frame.name][0]
    # Name the frame
    second_frame.name = "Physiography"
else:
    logger.info("..no second frame..")
    second_frame = first_frame

# Make next frame active
mxd.activeView = second_frame.name

# Copy places and boundary over to same dataframe
for i in [i for i in arcpy.mapping.ListLayers(first_frame) if not i.visible]:
    logger.info('..moving {0} over to next frame..'.format(i.name))
    arcpy.mapping.AddLayer(second_frame, i, "AUTO_ARRANGE")
    # Remove it from the first frame
    arcpy.mapping.RemoveLayer(first_frame, i)

# Get reference
places = [i for i in arcpy.mapping.ListLayers(second_frame) if "places" in i.name.lower()][0]

# Select towns
arcpy.SelectLayerByAttribute_management(places, "NEW_SELECTION", '"TYPE" =  1')

# Export to shapefile
logger.info('..creating towns shp..')
arcpy.FeatureClassToShapefile_conversion(places, arcpy.env.workspace)

# Import to TOC
towns = [j[:-4] for j in arcpy.ListFeatureClasses()]
towns = [i for i in towns if i not in [k.datasetName for k in arcpy.mapping.ListLayers(mxd)]][0] + ".shp"
arcpy.mapping.AddLayer(second_frame, arcpy.mapping.Layer(towns))

# Get reference for ease of query later
# Normally only references to df are needed
towns = [i for i in arcpy.mapping.ListLayers(second_frame) if towns in i.dataSource][0]
towns.name = "Towns"
towns.visible = False

# Can't georeference or digitize via script
# Check for physiography.shp
if os_find_file("physiography", ".shp"):
    logger.info('..working with physiography..')
    # Add physio to new frame
    map_lyr_names = [i for i in arcpy.mapping.ListLayers(second_frame)]
    physiography = [i for i in arcpy.ListFeatureClasses() if "physiography" in i][0]
    arcpy.mapping.AddLayer(second_frame, arcpy.mapping.Layer(physiography))
    physiography = [i for i in arcpy.mapping.ListLayers(second_frame) if i not in map_lyr_names][0]

    # Check if Type field has been created
    if not [True for i in arcpy.ListFields(physiography.dataSource) if i.name == "Type"][0]:
        arcpy.AddField_management(physiography, "Type", "TEXT", field_length = 50)

    # Check, then Calculate area of each bedrock
    if not [True for i in arcpy.ListFields(physiography.dataSource) if i.name == "Area"][0]:
        arcpy.AddField_management(physiography, "Area", "FLOAT", 12, 2)
        arcpy.CalculateField_management(physiography, "Area", "!SHAPE.AREA@SQUAREKILOMETERS!", "PYTHON_9.3")

    # Select larger of two entry (Lowlands)
    records = arcpy.SearchCursor(physiography) #returns gen obj!
    values = []
    for row in records:
        for field in arcpy.ListFields(physiography.dataSource):
            if field.name == "Area":
                values.append(row.getValue(field.name))
    types = ["Canadian Shield" if max(values) != i else "St. Lawrence Lowlands" for i in values]
    records = arcpy.UpdateCursor(physiography)
    field_type = [i for i in arcpy.ListFields(physiography.dataSource) if i.name == "Type"][0]
    field_area = [i for i in arcpy.ListFields(physiography.dataSource) if i.name == "Area"][0]
    logger.info('..setting new values for "Type" field..')
    for row in records:
        row.setValue(field_type.name, types[values.index(row.getValue(field_area.name))])
        records.updateRow(row)
    del records, row, field_type, field_area
    #arcpy.CalculateField_management(physiography, "Type", "1600", "PYTHON_9.3")
    #arcpy.CalculateField_management(physiography, "Type", "1600", "PYTHON_9.3")

    # Select by location, field == Lowlands
    arcpy.SelectLayerByAttribute_management(physiography, "NEW_SELECTION", '"Type" = \'Canadian Shield\'')
    arcpy.SelectLayerByLocation_management(towns, "COMPLETELY_WITHIN", physiography)
    logger.info('..creating shield_towns shp..')
    arcpy.FeatureClassToShapefile_conversion(towns, arcpy.env.workspace)

    # Import to TOC
    map_lyr_names = [i.dataSource.split("\\")[-1] for i in arcpy.mapping.ListLayers(mxd)]
    shield_towns = [i for i in arcpy.ListFeatureClasses()]
    shield_towns = [i for i in shield_towns if i not in map_lyr_names][0]
    arcpy.mapping.AddLayer(second_frame, arcpy.mapping.Layer(shield_towns))

     # Get reference
    shield_towns = [i for i in arcpy.mapping.ListLayers(second_frame) if i.name not in map_lyr_names][0]
    shield_towns.name = "Shield Towns"

    # Clear selections
    arcpy.SelectLayerByAttribute_management(physiography, "CLEAR_SELECTION")
    arcpy.SelectLayerByAttribute_management(towns, "CLEAR_SELECTION")

    # Select by location, field == Shield
    arcpy.SelectLayerByAttribute_management(physiography, "NEW_SELECTION", '"Type" = \'St. Lawrence Lowlands\'')
    arcpy.SelectLayerByLocation_management(towns, "COMPLETELY_WITHIN", physiography)
    logger.info('..creating lowland_towns shp..')
    arcpy.FeatureClassToShapefile_conversion(towns, arcpy.env.workspace)

    # Import to TOC
    map_lyr_names = [i.dataSource.split("\\")[-1] for i in arcpy.mapping.ListLayers(mxd)]
    lowland_towns = [i for i in arcpy.ListFeatureClasses()]
    lowland_towns = [i for i in lowland_towns if i not in map_lyr_names][0]
    arcpy.mapping.AddLayer(second_frame, arcpy.mapping.Layer(lowland_towns))

    # Get reference
    lowland_towns = [i for i in arcpy.mapping.ListLayers(second_frame) if i.name not in map_lyr_names][0]
    lowland_towns.name = "Lowland Towns"

    # Clear selections
    arcpy.SelectLayerByAttribute_management(physiography, "CLEAR_SELECTION")
    arcpy.SelectLayerByAttribute_management(towns, "CLEAR_SELECTION")

    # Remove towns, places, and boundary
    for i in [j for j in arcpy.mapping.ListLayers(second_frame) if j.visible == False]:
        logger.info('..removing {0} lyr..'.format(i.name))
        arcpy.mapping.RemoveLayer(second_frame, i)

# Change layer names
for i in arcpy.mapping.ListLayers(second_frame):
    i.name = i.name.capitalize()

# Zoom to full extent, change scale, reset orientation
second_frame.zoomToSelectedFeatures()
# second_frame.scale = 750000
# third_frame.elementPositionX, third_frame.elementPositionY = 22, 12
# third_frame.elementHeight, third_frame.elementWidth = 13, 9

logger.info('..completely done with second frame..')

#####
### Working on Third Dataframe
# Create reference to new frame
if len(arcpy.mapping.ListDataFrames(mxd)) > 2:
    third_frame = [i for i in arcpy.mapping.ListDataFrames(mxd) if i.name != first_frame.name if i.name != second_frame.name][0]

    # Name the frame
    third_frame.name = "Stream Density"
else:
    logger.info('..no third frame..')
    third_frame = first_frame

# Make next frame active
mxd.activeView = third_frame.name

# Copy places and boundary over to same dataframe
for i in [i for i in arcpy.mapping.ListLayers(first_frame) if "streams" in i.name.lower()]:
    # Extra condition if less than three frames
    if i not in arcpy.mapping.ListLayers(third_frame):
        arcpy.mapping.AddLayer(third_frame, i, "AUTO_ARRANGE")

# Copy streams over to Steam Density
# ListRasters
if arcpy.CheckExtension('Spatial'):
    arcpy.CheckOutExtension('Spatial')
    logger.info('..calculating line density..')
    streamden = arcpy.sa.LineDensity([i for i in arcpy.mapping.ListLayers(third_frame) if 'streams' in i.name.lower()][0], None, 50, 2000, "SQUARE_KILOMETERS")
    streamden.save("streamden")

    # Import to TOC
    arcpy.mapping.AddLayer(third_frame, arcpy.mapping.Layer(arcpy.ListRasters()[0]))

    ## Create graphs?

    arcpy.CheckInExtension('Spatial')
else:
    logger.warn('Extension not available.')

# Change layer names
for i in arcpy.mapping.ListLayers(third_frame):
    i.name = i.name.capitalize()

# Zoom to full extent, change scale, reset orientation
third_frame.zoomToSelectedFeatures()
# third_frame.scale = 700000
# third_frame.elementPositionX, third_frame.elementPositionY = 33, 12
# third_frame.elementHeight, third_frame.elementWidth = 13, 9

mxd.title = 'Semester_Project'

mxd.saveACopy(arcpy.env.workspace + '/' + mxd.filePath.split("\\")[-1].split(".")[0] + "_new.mxd")
# mxd.saveACopy(arcpy.env.workspace + '/' + "Semester_Project.mxd")

# arcpy.mapping.ExportToPDF(mxd, arcpy.env.workspace + "/" + mxd.title)

logger.info('..script end..')
