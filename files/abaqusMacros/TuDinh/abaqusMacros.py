from abaqus import *
from abaqusConstants import *


def Create_Sets_Surfaces():

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

Create_Sets_Surfaces()

def A_Create_Part():
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


