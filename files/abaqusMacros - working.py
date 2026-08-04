from abaqus import *
from abaqusConstants import *

    # --------------------------------
    # Import Pin & Box
    # --------------------------------

def A():

    import section
    import regionToolset
    import displayGroupMdbToolset as dgm
    import part
    import material
    import assembly
    import step
    import interaction
    import load
    import mesh
    import optimization
    import job
    import sketch
    import visualization
    import xyPlot
    import displayGroupOdbToolset as dgo
    import connectorBehavior
    import os
    from dxf2abq import importdxf
    workdir = os.getcwd()
    importdxf(fileName=os.path.join(workdir, 'PIN.DXF'))
    a = mdb.models['Model-1'].rootAssembly
    session.viewports['Viewport: 1'].setValues(displayedObject=a)
    from dxf2abq import importdxf
    importdxf(fileName=os.path.join(workdir, 'BOX.DXF'))


    # --------------------------------
    # Create Part Pin & Box
    # --------------------------------

    import connectorBehavior
    s1 = mdb.models['Model-1'].ConstrainedSketch(name='__profile__', 
        sheetSize=200.0)
    g, v, d, c = s1.geometry, s1.vertices, s1.dimensions, s1.constraints
    s1.sketchOptions.setValues(viewStyle=AXISYM)
    s1.setPrimaryObject(option=STANDALONE)
    s1.ConstructionLine(point1=(0.0, -100.0), point2=(0.0, 100.0))
    s1.FixedConstraint(entity=g[2])
    s1.retrieveSketch(sketch=mdb.models['Model-1'].sketches['PIN'])
    session.viewports['Viewport: 1'].view.fitView()
    p = mdb.models['Model-1'].Part(name='PIN', dimensionality=AXISYMMETRIC, 
        type=DEFORMABLE_BODY)
    p = mdb.models['Model-1'].parts['PIN']
    p.BaseShell(sketch=s1)
    s1.unsetPrimaryObject()
    p = mdb.models['Model-1'].parts['PIN']
    session.viewports['Viewport: 1'].setValues(displayedObject=p)
    del mdb.models['Model-1'].sketches['__profile__']
    s = mdb.models['Model-1'].ConstrainedSketch(name='__profile__', 
        sheetSize=200.0)
    g, v, d, c = s.geometry, s.vertices, s.dimensions, s.constraints
    s.sketchOptions.setValues(viewStyle=AXISYM)
    s.setPrimaryObject(option=STANDALONE)
    s.ConstructionLine(point1=(0.0, -100.0), point2=(0.0, 100.0))
    s.FixedConstraint(entity=g[2])
    s.retrieveSketch(sketch=mdb.models['Model-1'].sketches['BOX'])
    session.viewports['Viewport: 1'].view.fitView()
    p = mdb.models['Model-1'].Part(name='BOX', dimensionality=AXISYMMETRIC, 
        type=DEFORMABLE_BODY)
    p = mdb.models['Model-1'].parts['BOX']
    p.BaseShell(sketch=s)
    s.unsetPrimaryObject()
    p = mdb.models['Model-1'].parts['BOX']
    session.viewports['Viewport: 1'].setValues(displayedObject=p)
    del mdb.models['Model-1'].sketches['__profile__']
    # --------------------------------
    # Create_Sets_Surfaces
    # --------------------------------

    model_name = list(mdb.models.keys())[0]
    model = mdb.models[model_name]

    assembly = model.rootAssembly


    # --------------------------------
    # Create temporary dummy part
    # --------------------------------

    sketch = model.ConstrainedSketch(
        name='TempSketch',
        sheetSize=100.0
    )


    for i in range(6):

        y = float(i)

        sketch.Line(
            point1=(0.0, y),
            point2=(0.001, y)
        )


    temp_part = model.Part(
        name='DummyGeometry',
        dimensionality=TWO_D_PLANAR,
        type=DEFORMABLE_BODY
    )


    temp_part.BaseWire(
        sketch=sketch
    )


    # --------------------------------
    # Create temporary instance
    # --------------------------------

    assembly.Instance(
        name='DummyGeometry-1',
        part=temp_part,
        dependent=ON
    )


    dummy = assembly.instances['DummyGeometry-1']


    # --------------------------------
    # Names
    # --------------------------------

    surface_names = [
        'AXIAL',
        'PINT',
        'PEXT',
        'PinThreads',
        'BoxThreads',
        'BoxEnd'
    ]


    set_names = [
        'FixPin',
        'FixBox',
        'IntPressPenMaster',
        'IntPressPenSlave',
        'ExtPressPenSlave',
        'ExtPressPenMaster'
    ]


    # --------------------------------
    # Create surfaces and sets
    # --------------------------------

    for i in range(6):

        edge = dummy.edges[i:i+1]


        assembly.Set(
            name=set_names[i],
            edges=edge
        )


        assembly.Surface(
            name=surface_names[i],
            side1Edges=edge
        )


    # --------------------------------
    # Remove dummy geometry
    # --------------------------------

    del assembly.instances['DummyGeometry-1']

    del model.parts['DummyGeometry']

    del model.sketches['TempSketch']


    print('================================')
    print('Created 6 Sets:')
    for name in set_names:
        print(name)

    print('Created 6 Surfaces:')
    for name in surface_names:
        print(name)

    print('Dummy geometry removed')
    print('================================')


    # --------------------------------
    # Material Creation
    # --------------------------------

    if 'P-110' in model.materials.keys():
        del model.materials['P-110']

    mat = model.Material(name='P-110')

    mat.Elastic(
        table=((30000000.0, 0.3),)
    )

    mat.Plastic(
        table=(
            (100473.1, 0.0),
            (102947.7, 7.73615e-05),
            (105412.6, 0.000170728),
            (110505.5, 0.000578004),
            (112572.1, 0.000874124),
            (114651.3, 0.001328723),
            (116726.1, 0.001989506),
            (118865.3, 0.003127101),
            (121140.5, 0.005088999),
            (122548.7, 0.006745395),
            (125142.1, 0.010646338),
            (127979.0, 0.016141818),
            (130720.0, 0.022479825),
            (132786.7, 0.028164075),
            (134699.2, 0.034055993),
            (136000.5, 0.038682456),
            (137307.5, 0.044243551),
            (137807.9, 0.046495517),
            (138654.5, 0.050881691),
            (139517.6, 0.055992652),
            (140566.6, 0.063125487),
            (140883.6, 0.065334901),
            (141612.3, 0.072453384),
            (142250.6, 0.081627384),
            (142539.9, 0.090890538),
        )
    )

    print('Material P-110 created successfully.')



    # --------------------------------
    # Assign Material 
    # --------------------------------



    model = mdb.models['Model-1']

    # Create section if it doesn't exist
    sectionName = 'P-110 Section'

    if sectionName not in model.sections.keys():
        model.HomogeneousSolidSection(
            name=sectionName,
            material='P-110',
            thickness=None
        )

    # -------------------------
    # PIN
    # -------------------------
    part = model.parts['PIN']

    if 'PIN_Set' not in part.sets.keys():
        part.Set(
            faces=part.faces[:],
            name='PIN_Set'
        )

    if len(part.sectionAssignments):
        del part.sectionAssignments[:]

    part.SectionAssignment(
        region=part.sets['PIN_Set'],
        sectionName=sectionName,
        offset=0.0,
        offsetType=MIDDLE_SURFACE,
        offsetField='',
        thicknessAssignment=FROM_SECTION
    )

    # -------------------------
    # BOX
    # -------------------------
    part = model.parts['BOX']

    if 'BOX_Set' not in part.sets.keys():
        part.Set(
            faces=part.faces[:],
            name='BOX_Set'
        )

    if len(part.sectionAssignments):
        del part.sectionAssignments[:]

    part.SectionAssignment(
        region=part.sets['BOX_Set'],
        sectionName=sectionName,
        offset=0.0,
        offsetType=MIDDLE_SURFACE,
        offsetField='',
        thicknessAssignment=FROM_SECTION
    )

    print('Material P-110 assigned to PIN and BOX.')



    # --------------------------------
    # Assembly
    # --------------------------------


    a = mdb.models['Model-1'].rootAssembly
    a.DatumCsysByThreePoints(coordSysType=CYLINDRICAL, origin=(0.0, 0.0, 0.0), 
        point1=(1.0, 0.0, 0.0), point2=(0.0, 0.0, -1.0))
    p = mdb.models['Model-1'].parts['BOX']
    a.Instance(name='BOX-1', part=p, dependent=OFF)
    p = mdb.models['Model-1'].parts['PIN']
    a.Instance(name='PIN-1', part=p, dependent=OFF)


# --------------------------------
# Load Schedule from Excel
# --------------------------------

import os
import win32com.client
from collections import defaultdict


# Current Abaqus working directory
workdir = os.getcwd()


# --------------------------------
# Find Excel load schedule
# --------------------------------

excelFiles = []

for file in os.listdir(workdir):
    if file.lower().endswith(('.xls', '.xlsx', '.xlsm')):
        excelFiles.append(file)


if len(excelFiles) == 0:
    raise Exception(
        'No Excel load schedule found in working directory.'
    )


if len(excelFiles) > 1:
    raise Exception(
        'More than one Excel file found. Keep only one load schedule file.'
    )


excelFile = os.path.join(workdir, excelFiles[0])


print('Loading load schedule:')
print(excelFile)


# --------------------------------
# Open Excel
# --------------------------------

xl = win32com.client.Dispatch("Excel.Application")

xl.Visible = False

wb = xl.Workbooks.Open(excelFile)

ws = wb.Worksheets(1)


# --------------------------------
# Read load schedule
# Row 1 = title
# Row 2 = headers
# Data starts row 3
# --------------------------------

StepsList = []
Loads = []
PInt = []
PExt = []
Steps = []

tracker = defaultdict(int)


# Find last used row in column A
lastRow = ws.Cells(
    ws.Rows.Count,
    1
).End(-4162).Row     # xlUp


for i in range(3, lastRow + 1):

    # Column A
    loadPoint = int(
        ws.Cells(i,1).Value
    )

    # Column B
    pressure = float(
        ws.Cells(i,2).Value
    )

    # Column C
    axial = float(
        ws.Cells(i,3).Value
    )


    StepsList.append(loadPoint)


    # Abaqus axial sign convention
    Loads.append(-axial)


    # Pressure split
    if pressure >= 0:

        PInt.append(pressure)
        PExt.append(0.0)

    else:

        PInt.append(0.0)
        PExt.append(-pressure)



# --------------------------------
# Create step names
# --------------------------------

for stepNumber in StepsList:

    tracker[stepNumber] += 1

    Steps.append(
        'LP' + str(stepNumber) + '_' + str(tracker[stepNumber])
    )


# --------------------------------
# Close Excel
# --------------------------------

wb.Close(False)

xl.Quit()


print('--------------------------------')
print('Load schedule imported')
print('Number of load points:', len(Steps))
print('--------------------------------')


