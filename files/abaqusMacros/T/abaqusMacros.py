# Do not delete the following import lines
from abaqus import *
from abaqusConstants import *
import __main__
#from dxf2abq import importdxf 

def X_EnterSteps():
    import visualization
    import xyPlot
    import displayGroupOdbToolset as dgo
    import win32com.client                                    #Allows imports from "Active" Excel document
    from dxf2abq import importdxf
    
    ## Get the Model Name
    model = getInput('Enter the model name', default = mdb.models.keys()[0])
    a = mdb.models[model].rootAssembly

     ##Imports from excel
    xl = win32com.client.Dispatch("Excel.Application")        #Type xl. before using Excel procedures

    if xl.ActiveSheet == None:                                 #Tests to see if an Excel sheet is open
        reply = getWarningReply(message = 'Need to have "Load" Excel sheet open', buttons = ('OK'))
        
    CellNum = xl.ActiveSheet.UsedRange.Rows.Count            #index of the last used cell in a column
    Pin_Sketches = []
    Box_Sketches = []

    for i in range(3,CellNum -1):                            #Create load and step arrays in abaqus after importing from excel
        Pin_Sketches.append(xl.Cells(i,1).value)  
        Box_Sketches.append(xl.Cells(i,2).value) 
    
    importdxf(fileName='H:/2012/QX-HT/FEA Macros/Sketches/' + str(Box_Sketches[i-3]) + '.DXF')

    ##Enter the Steps
    vp = session.viewports[session.currentViewportName]

def A_CreateModels():
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
    import win32com.client                                    #Allows imports from "Active" Excel document
    from dxf2abq import importdxf    
    
    ##Imports from excel
    xl = win32com.client.Dispatch("Excel.Application")        #Type xl. before using Excel procedures
    CellNum = xl.ActiveSheet.UsedRange.Rows.Count            #index of the last used cell in a column
    
    print CellNum
    Models = []

    for i in range(2,7):                            #Create load and step arrays in abaqus after importing from excel
        Models.append(xl.Cells(i,2).value)  
        mdb.Model(name= str(xl.Cells(i,2).value), objectToCopy=mdb.models['Template'])

def B_ImportSketches():
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
    import win32com.client                                    #Allows imports from "Active" Excel document
    from dxf2abq import importdxf    
    
    ##Imports from excel
    xl = win32com.client.Dispatch("Excel.Application")        #Type xl. before using Excel procedures
    CellNum = xl.ActiveSheet.UsedRange.Rows.Count            #index of the last used cell in a column
    
    print CellNum
    Models = []
    Pin_Sketch_Name = []
    Box_Sketch_Name = []
    
    for i in range(2,7):                            #Create load and step arrays in abaqus after importing from excel
        print i

        p1 = mdb.models[str(xl.Cells(i,2).value)].parts['Box']
        session.viewports['Viewport: 1'].setValues(displayedObject=p1)
        session.viewports['Viewport: 1'].setValues(displayedObject=None)
        from dxf2abq import importdxf
        importdxf(
        fileName='C:\FEA\FEA Demonstration\Sketches/' + str(xl.Cells(i,3).value) + '.DXF')
        importdxf(
        fileName='C:\FEA\FEA Demonstration\Sketches/' + str(xl.Cells(i,4).value) + '.DXF')

def E_AddSketchestoParts():
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
    import win32com.client                                    #Allows imports from "Active" Excel document
    
    ##Imports from excel
    xl = win32com.client.Dispatch("Excel.Application")        #Type xl. before using Excel procedures
    CellNum = xl.ActiveSheet.UsedRange.Rows.Count            #index of the last used cell in a column
    
    print CellNum
    Models = []
    Pin_Sketch_Name = []
    Box_Sketch_Name = []
    
    Part_Type = getInput('Enter "Pin" or "Box"', default = 'Pin')
    if Part_Type == 'Box':
        Part_Index=4
    elif Part_Type=='Pin':
        Part_Index=3
    else:
        Part_Index=5
    
    for i in range(2,7):
        p1 = mdb.models[str(xl.Cells(i,2).value)].parts[Part_Type]
        print 'Model Name: ' + str(xl.Cells(i,2).value)
        session.viewports['Viewport: 1'].setValues(displayedObject=p1)
        s = mdb.models[str(xl.Cells(i,2).value)].ConstrainedSketch(name='__profile__', sheetSize=200.0, gridSpacing=5.0)
        g, v, d, c = s.geometry, s.vertices, s.dimensions, s.constraints
        s.sketchOptions.setValues(viewStyle=AXISYM)
        s.setPrimaryObject(option=SUPERIMPOSE)
        p = mdb.models[str(xl.Cells(i,2).value)].parts[Part_Type]
        p.projectReferencesOntoSketch(sketch=s, filter=COPLANAR_EDGES)
        s.ConstructionLine(point1=(0.0, -100.0), point2=(0.0, 100.0))
        print 'C'
        s.retrieveSketch(
            sketch=mdb.models[str(xl.Cells(i,2).value)].sketches[str(xl.Cells(i,Part_Index).value)])
        session.viewports['Viewport: 1'].view.fitView()
        print 'D'
        p = mdb.models[str(xl.Cells(i,2).value)].parts[Part_Type]
        print 'D'
        p.Shell(sketch=s)
        print 'E'
        s.unsetPrimaryObject()
        del mdb.models[str(xl.Cells(i,2).value)].sketches['__profile__']

def C_Rotate():
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
    import win32com.client                                    #Allows imports from "Active" Excel document
    
    ##Imports from excel
    xl = win32com.client.Dispatch("Excel.Application")        #Type xl. before using Excel procedures
    CellNum = xl.ActiveSheet.UsedRange.Rows.Count            #index of the last used cell in a column

    Models = []
    for j in range(3,5):
        for i in range(2,7):    
            session.viewports['Viewport: 1'].setValues(displayedObject=None)
            s1 = mdb.models[str(xl.Cells(i,2).value)].ConstrainedSketch(name='__edit__', 
                    objectToCopy=mdb.models[str(xl.Cells(i,2).value)].sketches[str(xl.Cells(i,j).value)])
            g, v, d, c = s1.geometry, s1.vertices, s1.dimensions, s1.constraints
            s1.setPrimaryObject(option=STANDALONE)
            s1.rotate(centerPoint=(7.9520666967, 5.4862252367), angle=-90.0, objectList=(
                g.values()))
            s1.move(vector=(20, 20), objectList=(g.values())) # Move the geometry away from the centerline
            mdb.models[str(xl.Cells(i,2).value)].sketches.changeKey(fromName='__edit__', 
                toName=str(xl.Cells(i,j).value))
            s1.unsetPrimaryObject()

def Z_GetBoundingBox():
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
    import win32com.client                                    #Allows imports from "Active" Excel document
    
    ##Imports from excel
    xl = win32com.client.Dispatch("Excel.Application")        #Type xl. before using Excel procedures
    CellNum = xl.ActiveSheet.UsedRange.Rows.Count            #index of the last used cell in a column
    Type = ['Pin', 'Box']
    for j in Type:
        for i in range(2,22):    
            p = mdb.models[str(xl.Cells(i,1).value)].parts[j]
            tuple = p.edges.getBoundingBox()
            tuple_high= tuple['high']
            print str(xl.Cells(i,1).value)+ ' ' + j + ' ' + ' ' + str(tuple_high[0]) + ' ' + str(tuple_high[1])

def F_Move():
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
    import win32com.client                                    #Allows imports from "Active" Excel document
    
    ##Imports from excel
    xl = win32com.client.Dispatch("Excel.Application")        #Type xl. before using Excel procedures
    CellNum = xl.ActiveSheet.UsedRange.Rows.Count            #index of the last used cell in a column
    Type = ['Pin', 'Box']
    for j in Type:
        for i in range(2,7):    
            print 'i: ' + str(i)
            p = mdb.models[str(xl.Cells(i,2).value)].parts[j]
            tuple = p.edges.getBoundingBox()
            tuple_high= tuple['high']
            tuple_low = tuple['low']
            print str(xl.Cells(i,2).value)+ ' ' + j + ' ' + ' ' + str(tuple_high[0]) + ' ' + str(tuple_high[1])

            
            #p1 = mdb.models['QX_XT 5,500_23,00 FTL 0,6'].parts['Box']
            #session.viewports['Viewport: 1'].setValues(displayedObject=p1)
            s = p.features['Shell planar-1'].sketch
            mdb.models[str(xl.Cells(i,2).value)].ConstrainedSketch(name='__edit__', 
                objectToCopy=s)
            s1 = mdb.models[str(xl.Cells(i,2).value)].sketches['__edit__']
            g, v, d, c = s1.geometry, s1.vertices, s1.dimensions, s1.constraints
            s1.setPrimaryObject(option=SUPERIMPOSE)
            p.projectReferencesOntoSketch(sketch=s1, 
                upToFeature=p.features['Shell planar-1'], filter=COPLANAR_EDGES)
            g2 = list(g.values())
            if j == 'Pin':
                s1.move(vector=(xl.Cells(i,6).value-tuple_high[0], xl.Cells(i,7).value-tuple_high[1]), objectList=(g2[1:]))
                print str(xl.Cells(i,2).value)+ ' ' + str(j) + 'moved'
            if j == 'Box':
                s1.move(vector=(xl.Cells(i,5).value-tuple_low[0], -tuple_low[1]), objectList=(g2[1:]))
                print str(xl.Cells(i,2).value)+ ' ' + str(j) + 'moved'                
            s1.unsetPrimaryObject()
            p = mdb.models[str(xl.Cells(i,2).value)].parts[j]
            p.features['Shell planar-1'].setValues(sketch=s1)
            del mdb.models[str(xl.Cells(i,2).value)].sketches['__edit__']
            p = mdb.models[str(xl.Cells(i,2).value)].parts[j]
            p.regenerate()
    
def Z_CreateSet():
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
    a = mdb.models['QX_XT 5,500_23,00 FTL 0,2'].rootAssembly
    f1 = a.instances['Box-1'].faces
    faces1 = f1.getSequenceFromMask(mask=('[#1 ]', ), )
    a.Set(faces=faces1, name='Box OD')

def Z_CreateBoxThreadSurface():
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
    import win32com.client                                    #Allows imports from "Active" Excel document
    
    ##Imports from excel
    xl = win32com.client.Dispatch("Excel.Application")        #Type xl. before using Excel procedures
    


    for i in range(2,22):
        
        a = mdb.models[str(xl.Cells(i,1).value)].rootAssembly
        a.regenerate
        a.features['Pin-1'].suppress()
        s1 = a.instances['Box-1'].edges.getByBoundingBox( xMin=xl.Cells(i,18).value , xMax =xl.Cells(i,17).value, yMin = xl.Cells(i,16).value, yMax= xl.Cells(i,15).value)
        side1Edges1 = s1.getSequenceFromMask(mask=('[#ffffffe1 #ffffffff:2 #3f ]', ), )
        a.Surface(side1Edges=side1Edges1, name='Box Thread')
        a.features['Pin-1'].resume()

def X_Format_Viewport_Font():
    # import selection
    # import regionToolset
    # import displayGroupMdbToolset as dgm
    # import part
    # import material
    # import assembly
    # import step
    # import interaction
    # import load
    # import mesh
    # import optimization
    # import job
    # import sketch
    import visualization
    import xyPlot
    import displayGroupOdbToolset as dgo
    # import connectorBehavior

    inp = getInput('What size font do you want to use?', '14')
    
    for key in session.viewports.keys():
        session.viewports[key].viewportAnnotationOptions.setValues(
        triadFont='-*-verdana-bold-r-normal-*-*-' + inp + '0-*-*-p-*-*-*',
        legendFont='-*-verdana-bold-r-normal-*-*-' + inp + '0-*-*-p-*-*-*',
        titleFont='-*-verdana-bold-r-normal-*-*-' + inp + '0-*-*-p-*-*-*',
        stateFont='-*-verdana-bold-r-normal-*-*-' + inp + '0-*-*-p-*-*-*')

def X_PrintAllSteps():
    # import section
    # import regionToolset
    # import displayGroupMdbToolset as dgm
    # import part
    # import material
    # import assembly
    # import step
    # import interaction
    # import load
    # import mesh
    # import optimization
    # import job
    # import sketch
    import visualization
    import xyPlot
    import displayGroupOdbToolset as dgo
    import time

    #Count the Steps
    lastStep = len(session.odbs[session.odbs.keys()[0]].steps)        #last step of the current output database: 'C:/FEA/FEAautomation/Job-1.odb'
    
    for x in range(0,lastStep):
        for key in session.viewports.keys():
            vp = session.viewports[key]
            vp.makeCurrent
            step_name = session.odbs[session.odbs.keys()[0]].steps.keys()[x]         #Name/key of the current step
            lastFrame = session.odbs[session.odbs.keys()[0]].steps[step_name].frames[-1].incrementNumber    #index of the last frame of the current step
            vp.odbDisplay.setFrame(step=x, frame=lastFrame)
            od = vp.odbDisplay

            Var = od.primaryVariable[0]                #This is the symbol representation of the current load type
            var2 = od.primaryVariable[5]
            
            session.pngOptions.setValues(imageSize = (2272,1704))
            session.printOptions.setValues(rendition=COLOR, vpDecorations=OFF, vpBackground=OFF)
            
            #Saves a file with a unique name for each step of a load type
            session.printToFile(fileName= 'Stress' + '_' + Var + '_' + var2 + '(' + str(x) + ')', format = PNG, canvasObjects= (vp,))    

def D_FilletDominator():
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
    import win32com.client                                    #Allows imports from "Active" Excel document
    
    ###############################################################################################################################################################################
    #                    Imports from Excel
    ###############################################################################################################################################################################    
    xl = win32com.client.Dispatch("Excel.Application")        #Type xl. before using Excel procedures
    CellNum = xl.ActiveSheet.UsedRange.Rows.Count            #index of the last used cell in a column

    for i in range(2,20):    




    ###############################################################################################################################################################################
    #                    Inputs Portion
    ###############################################################################################################################################################################    
        Model_Name = str(xl.Cells(i,2).value)
        Sketch_Name = str(xl.Cells(i,4).value)
        RequestedFilletRadius = 0.004
        
    ###############################################################################################################################################################################
    #                    Check if Sketcch is Closed Portion
    ###############################################################################################################################################################################    
    #
    #
    #
    #
    #


        
    ###############################################################################################################################################################################
    #                    Geometery Portion
    ###############################################################################################################################################################################
        Geometry=['Geometry']
        Geometry_Type=['Geometry_Type']
        X1=['X1']
        Y1=['Y1']
        X2=['X2']
        Y2=['Y2']
        X3=['X3']
        Y3=['Y3']
        #List Count starts at 0
        Combined_List = [ Geometry , Geometry_Type , X1 , Y1 , X2 , Y2 , X3 , Y3]
        
        sketch =mdb.models[Model_Name].sketches[Sketch_Name]
        geometry_list = list(sketch.geometry.keys())
        
        for i in geometry_list: 
            #print 'Geometry Key is: ' + str(i)
            current_geometry= sketch.geometry[i]
            
            Geometry.append(current_geometry.id) # Append the geometry's name to the list
            Geometry_Type.append(current_geometry.curveType) # Append the geometry's type to the list
            #print 'Curve Type: ' + str(current_geometry.curveType)
            vert = current_geometry.getVertices()

            #print vert
            type = str(current_geometry.curveType)
            for k in range(3):
                if type == 'LINE':
                    if k == 0:     # Append the Vertices's name to the corresponding list        
                        X1.append(vert[k].coords[0])
                        Y1.append(vert[k].coords[1])
                        #print 'X1: ' + str(X1)
                    elif k == 1:
                        X2.append(vert[k].coords[0])
                        Y2.append(vert[k].coords[1])                
                    elif k == 2:
                        X3.append('This is a line')
                        Y3.append('This is a line')
                    else:
                        print 'else'                                
                elif type == 'ARC':
                    if k == 0:     # Append the Vertices's name to the corresponding list        
                        X1.append(vert[k].coords[0])
                        Y1.append(vert[k].coords[1])
                    elif k == 1:
                        X2.append(vert[k].coords[0])
                        Y2.append(vert[k].coords[1])                
                    elif k == 2:
                        X3.append(vert[k].coords[0])
                        Y3.append(vert[k].coords[1])
                    else:
                        print 'else'                      
                else:
                    print 'This is not a line or an arc'
        
        myOutputFile = open('Geometry_data.txt','w+')
        m = 0
        #print 'range(len(Geometry)-1): ' +str(range(len(Geometry)-1))
        #print 'range(len(Combined_List)-1): ' +str(range(len(Combined_List)))
        
        for Count1 in range(len(Geometry)):
            new_line = ''
            n = 0

            for Count2 in range(len(Combined_List)):
                appendage = str(Combined_List[n][m]) + '~'
                new_line = new_line + appendage
                #print 'n: ' + str(n)
                n = n +1
            new_line = new_line + '\n'
            myOutputFile.write(new_line)
            m = m +1    




    ###############################################################################################################################################################################
    #                    Vertice Portion
    ###############################################################################################################################################################################    
        Vertice=['Vertice']
        X=['X']
        Y=['Y']
        Geometry1=['Geometry1']
        Geometry1_Type=['Geometry1_Type']
        X11=['X11']
        Y11=['Y11']
        X12=['X12']
        Y12=['Y12']
        X13=['X13']
        Y13=['Y13']
        X1_mid=['X1_mid']
        Y1_mid=['Y1_mid']
        Geometry1_Angle=['Geometry1_Angle']
        Geometry2=['Geometry2']
        Geometry2_Type=['Geometry2_Type']    
        X21=['X21']
        Y21=['Y21']
        X22=['X22']
        Y22=['Y22']
        X23=['X23']
        Y23=['Y23']
        X2_mid=['X2_mid']
        Y2_mid=['Y2_mid']
        Geometry2_Angle=['Geometry2_Angle']
        AngleBetween=['AngleBetween']
        MaxFilletRadius=['MaxFilletRadius']
        FilletRadius=['FilletRadius']

        #List Count starts at 0
        Combined_List = [ Vertice , X , Y , Geometry1,Geometry1_Type,X11,Y11,X12,Y12,X13,Y13,X1_mid,Y1_mid,Geometry1_Angle,Geometry2, Geometry2_Type, X21,Y21,X22,Y22,X23,Y23,X2_mid,Y2_mid,Geometry2_Angle,AngleBetween,MaxFilletRadius,FilletRadius]
        
        
        sketch =mdb.models[Model_Name].sketches[Sketch_Name]
        

        vert_keys = sketch.vertices.keys()

        for i in vert_keys: 
            #print 'Vertice Key is: ' + str(i)
            Vertice.append(i)
            X.append(sketch.vertices[i].coords[0])
            Y.append(sketch.vertices[i].coords[1])
            
            Geometry1.append('-')
            Geometry1_Type.append('-')
            X11.append('-')
            Y11.append('-')
            X12.append('-')
            Y12.append('-')
            X13.append('-')
            Y13.append('-')

            X1_mid.append('-')
            Y1_mid.append('-')
            Geometry1_Angle.append('-')
            Geometry2.append('-')
            Geometry2_Type.append('-')
            X21.append('-')
            Y21.append('-')
            X22.append('-')
            Y22.append('-')
            X23.append('-')
            Y23.append('-')
            X2_mid.append('-')
            Y2_mid.append('-')
            Geometry2_Angle.append('-')
            AngleBetween.append('-')
            MaxFilletRadius.append('-')
            FilletRadius.append('-')

        #print Combined_List    

        myOutputFile = open('Vertice_Data.txt','w+')
        m = 0
        
        for Count1 in range(len(Vertice)):
            new_line = ''
            n = 0

            for Count2 in range(len(Combined_List)):
                appendage = str(Combined_List[n][m]) + '~'
                new_line = new_line + appendage
                #print 'n: ' + str(n)
                n = n +1
            new_line = new_line + '\n'
            myOutputFile.write(new_line)
            m = m +1        

    ###############################################################################################################################################################################
    #                    Remove Duplicate Vertice Portion
    ###############################################################################################################################################################################    








            
    ###############################################################################################################################################################################
    #                    Combining Tables Portion
    ###############################################################################################################################################################################         
        for Count1 in range(len(Vertice)):
            # Search of the Geometry for the Corresponding Vertice. Start Search at Begining of List
            for Count2 in range(len(Geometry)):
                if X[Count1] ==X1[Count2] and Y[Count1] ==Y1[Count2]:
                    Geometry1.pop(Count1)
                    Geometry1.insert(Count1, Geometry[Count2])
                    Geometry1_Type.pop(Count1)
                    Geometry1_Type.insert(Count1, Geometry_Type[Count2])
                    X11.pop(Count1)
                    X11.insert(Count1, X1[Count2])
                    Y11.pop(Count1)
                    Y11.insert(Count1, Y1[Count2])
                    X12.pop(Count1)
                    X12.insert(Count1, X2[Count2])
                    Y12.pop(Count1)
                    Y12.insert(Count1, Y2[Count2])
                    X13.pop(Count1)
                    X13.insert(Count1, X3[Count2])
                    Y13.pop(Count1)
                    Y13.insert(Count1, Y3[Count2])
                    #print 'Y3: ' + str(Y3[Count2])
                    break
                if X[Count1] ==X2[Count2] and Y[Count1] ==Y2[Count2]:
                    Geometry1.pop(Count1)
                    Geometry1.insert(Count1, Geometry[Count2])
                    Geometry1_Type.pop(Count1)
                    Geometry1_Type.insert(Count1, Geometry_Type[Count2])
                    X11.pop(Count1)
                    X11.insert(Count1, X1[Count2])
                    Y11.pop(Count1)
                    Y11.insert(Count1, Y1[Count2])
                    X12.pop(Count1)
                    X12.insert(Count1, X2[Count2])
                    Y12.pop(Count1)
                    Y12.insert(Count1, Y2[Count2])
                    X13.pop(Count1)
                    X13.insert(Count1, X3[Count2])
                    Y13.pop(Count1)
                    Y13.insert(Count1, Y3[Count2])
                    break
                if X[Count1] ==X3[Count2] and Y[Count1] ==Y3[Count2]:
                    Geometry1.pop(Count1)
                    Geometry1.insert(Count1, Geometry[Count2])
                    Geometry1_Type.pop(Count1)
                    Geometry1_Type.insert(Count1, Geometry_Type[Count2])
                    X11.pop(Count1)
                    X11.insert(Count1, X1[Count2])
                    Y11.pop(Count1)
                    Y11.insert(Count1, Y1[Count2])
                    X12.pop(Count1)
                    X12.insert(Count1, X2[Count2])
                    Y12.pop(Count1)
                    Y12.insert(Count1, Y2[Count2])
                    X13.pop(Count1)
                    X13.insert(Count1, X3[Count2])
                    Y13.pop(Count1)
                    Y13.insert(Count1, Y3[Count2])
                    break
            # Search of the Geometry for the Corresponding Vertice. Start Search at end of List
            for Count2 in range(len(Geometry)):
                Count3 = len(Geometry)-1 - Count2 # Reverse the Order of previous for loop. Start at end of list
                if X[Count1] ==X3[Count3] and Y[Count1] ==Y3[Count3]:
                    Geometry2.pop(Count1)
                    Geometry2.insert(Count1, Geometry[Count3])
                    Geometry2_Type.pop(Count1)
                    Geometry2_Type.insert(Count1, Geometry_Type[Count3])
                    X21.pop(Count1)
                    X21.insert(Count1, X1[Count3])
                    Y21.pop(Count1)
                    Y21.insert(Count1, Y1[Count3])
                    X22.pop(Count1)
                    X22.insert(Count1, X2[Count3])
                    Y22.pop(Count1)
                    Y22.insert(Count1, Y2[Count3])
                    X23.pop(Count1)
                    X23.insert(Count1, X3[Count3])
                    Y23.pop(Count1)
                    Y23.insert(Count1, Y3[Count3])
                    break
                if X[Count1] ==X2[Count3] and Y[Count1] ==Y2[Count3]:
                    Geometry2.pop(Count1)
                    Geometry2.insert(Count1, Geometry[Count3])
                    Geometry2_Type.pop(Count1)
                    Geometry2_Type.insert(Count1, Geometry_Type[Count3])
                    X21.pop(Count1)
                    X21.insert(Count1, X1[Count3])
                    Y21.pop(Count1)
                    Y21.insert(Count1, Y1[Count3])
                    X22.pop(Count1)
                    X22.insert(Count1, X2[Count3])
                    Y22.pop(Count1)
                    Y22.insert(Count1, Y2[Count3])
                    X23.pop(Count1)
                    X23.insert(Count1, X3[Count3])
                    Y23.pop(Count1)
                    Y23.insert(Count1, Y3[Count3])
                    break
                if X[Count1] ==X1[Count3] and Y[Count1] ==Y1[Count3]:
                    Geometry2.pop(Count1)
                    Geometry2.insert(Count1, Geometry[Count3])
                    Geometry2_Type.pop(Count1)
                    Geometry2_Type.insert(Count1, Geometry_Type[Count3])
                    X21.pop(Count1)
                    X21.insert(Count1, X1[Count3])
                    Y21.pop(Count1)
                    Y21.insert(Count1, Y1[Count3])
                    X22.pop(Count1)
                    X22.insert(Count1, X2[Count3])
                    Y22.pop(Count1)
                    Y22.insert(Count1, Y2[Count3])
                    X23.pop(Count1)
                    X23.insert(Count1, X3[Count3])
                    Y23.pop(Count1)
                    Y23.insert(Count1, Y3[Count3])
                    break

    ###############################################################################################################################################################################
    #                    Get angles of Line Geometry
    ############################################################################################################################################################################### 

        for Count4 in range(len(Vertice)):
            #print 'Geometry1_Type: ' + str(Geometry1_Type[Count4]) + ' Count: ' + str(Count4)
            if str(Geometry1_Type[Count4]) =='LINE':
                #print 'Geometry1_Type: ' + str(Geometry1_Type[Count4]) + ' Count: ' + str(Count4)
                if X11[Count4]== X12[Count4]:
                    Geometry1_Angle.pop(Count4)
                    Geometry1_Angle.insert(Count4, 90)
                if Y11[Count4]== Y12[Count4]:
                    Geometry1_Angle.pop(Count4)
                    Geometry1_Angle.insert(Count4, 0)
                if X11[Count4]<> X12[Count4] and Y11[Count4]<> Y12[Count4]:
                    Geometry1_Angle.pop(Count4)
                    Angle = abs(degrees(atan( (Y11[Count4]- Y12[Count4])/(X11[Count4]- X12[Count4]))))
                    Geometry1_Angle.insert(Count4, Angle)                   
            if str(Geometry2_Type[Count4]) =='LINE':
                if X21[Count4]== X22[Count4]:
                    Geometry2_Angle.pop(Count4)
                    Geometry2_Angle.insert(Count4, 90)
                if Y21[Count4]== Y22[Count4]:
                    Geometry2_Angle.pop(Count4)
                    Geometry2_Angle.insert(Count4, 0)
                if X21[Count4]<> X22[Count4] and Y21[Count4]<> Y22[Count4]:
                    Geometry2_Angle.pop(Count4)
                    Angle = abs(degrees(atan( (Y21[Count4]- Y22[Count4])/(X21[Count4]- X22[Count4]))))
                    Geometry2_Angle.insert(Count4, Angle)
    ###############################################################################################################################################################################
    #                    Get angles of Arc Geometry
    ############################################################################################################################################################################### 
        print 'Step: Get angles of Arc Geometry'
        for Count4 in range(len(Vertice)-1):
            Count4 = Count4+1
            if str(Geometry1_Type[Count4]) =='ARC':
                #print 'Geometry1_Type: ' + str(Geometry1_Type[Count4]) + ' Count: ' + str(Count4)
                A = X11[Count4]
                B = Y11[Count4]
                C = X12[Count4]
                D = Y12[Count4]
                E =  mdb.models[Model_Name].sketches[Sketch_Name].geometry[Geometry1[Count4]].pointOn[0]
                F =  mdb.models[Model_Name].sketches[Sketch_Name].geometry[Geometry1[Count4]].pointOn[1]

                X_ = X[Count4]
                Y_ = Y[Count4]

                K = (.5)*((A**2+B**2)*(E-C) + (C**2+D**2)*(A-E) + (E**2+F**2)*(C-A)) / (B*(E-C)+D*(A-E)+F*(C-A))
                H = (.5)*((A**2+B**2)*(F-D) + (C**2+D**2)*(B-F) + (E**2+F**2)*(D-B)) / (A*(F-D)+C*(B-F)+E*(D-B)) 
                R = ((A-H)**2 + (B-K)**2 )**.5

                if Y_-K == 0:
                    Angle = 0
                elif X_-H == 0:
                    Angle = 90
                else:
                    Angle = abs(degrees(atan((X_-H)/(Y_-K))))
                Geometry1_Angle.pop(Count4)
                Geometry1_Angle.insert(Count4, Angle)
            if str(Geometry2_Type[Count4]) =='ARC':
                #print 'Geometry2_Type: ' + str(Geometry2_Type[Count4]) + ' Count: ' + str(Count4)
                A = X21[Count4]
                B = Y21[Count4]
                C = X22[Count4]
                D = Y22[Count4]

                E =  mdb.models[Model_Name].sketches[Sketch_Name].geometry[Geometry2[Count4]].pointOn[0]
                F =  mdb.models[Model_Name].sketches[Sketch_Name].geometry[Geometry2[Count4]].pointOn[1]
                
                X_ = X[Count4]
                Y_ = Y[Count4]

                K = (.5)*((A**2+B**2)*(E-C) + (C**2+D**2)*(A-E) + (E**2+F**2)*(C-A)) / (B*(E-C)+D*(A-E)+F*(C-A))
                H = (.5)*((A**2+B**2)*(F-D) + (C**2+D**2)*(B-F) + (E**2+F**2)*(D-B)) / (A*(F-D)+C*(B-F)+E*(D-B))            
                R = ((A-H)**2 + (B-K)**2 )**.5

                if Y_-K == 0:
                    Angle = 0
                elif X_-H == 0:
                    Angle = 90
                else:
                    Angle = abs(degrees(atan((X_-H)/(Y_-K))))
                #print 'A: ' + str(A)
                #print 'B: ' + str(B)
                #print 'C: ' + str(C)
                #print 'D: ' + str(D)
                #print 'E: ' + str(E)
                #print 'F: ' + str(F)
                #print 'X_: ' + str(X_    )        
                #print 'Y_: ' + str(Y_)
                #print 'K: ' + str(K)
                #print 'H: ' + str(H    )        
                #print 'R: ' + str(R)
                #print 'Angle: ' + str(Angle)                
                Geometry2_Angle.pop(Count4)
                Geometry2_Angle.insert(Count4, Angle)              
    ###############################################################################################################################################################################
    #                    Get midpoint of Line Geometry
    ############################################################################################################################################################################### 
        print 'Step: Get midpoint of Line Geometry'
        for Count5 in range(len(Vertice)):
            if str(Geometry1_Type[Count5]) =='LINE':
                X_mid = (X11[Count5] +X12[Count5])/2
                X1_mid.pop(Count5)
                X1_mid.insert(Count5, X_mid)
                Y_mid = (Y11[Count5] +Y12[Count5])/2            
                Y1_mid.pop(Count5)
                Y1_mid.insert(Count5, Y_mid)
            if str(Geometry2_Type[Count5]) =='LINE':
                X_mid = (X21[Count5] +X22[Count5])/2        
                X2_mid.pop(Count5)
                X2_mid.insert(Count5, X_mid)
                Y_mid = (Y21[Count5] +Y22[Count5])/2                
                Y2_mid.pop(Count5)
                Y2_mid.insert(Count5, Y_mid)
            if str(Geometry1_Type[Count5]) =='ARC':
                X_mid = mdb.models[Model_Name].sketches[Sketch_Name].geometry[Geometry1[Count5]].pointOn[0]
                X1_mid.pop(Count5)
                X1_mid.insert(Count5, X_mid)
                Y_mid = mdb.models[Model_Name].sketches[Sketch_Name].geometry[Geometry1[Count5]].pointOn[1]        
                Y1_mid.pop(Count5)
                Y1_mid.insert(Count5, Y_mid)
            if str(Geometry2_Type[Count5]) =='ARC':
                X_mid = mdb.models[Model_Name].sketches[Sketch_Name].geometry[Geometry2[Count5]].pointOn[0]    
                X2_mid.pop(Count5)
                X2_mid.insert(Count5, X_mid)
                Y_mid = mdb.models[Model_Name].sketches[Sketch_Name].geometry[Geometry2[Count5]].pointOn[1]                
                Y2_mid.pop(Count5)
                Y2_mid.insert(Count5, Y_mid)
    ###############################################################################################################################################################################
    #                    Get Angle between of Arc and Line Geometry
    ############################################################################################################################################################################### 
        print 'Step: Get angle between Arc and Line Geometry'
        for Count6 in range(len(Vertice)-1):
            Count6 = Count6 + 1
            #print 'Count6: ' + str(Count6)
            Anglebtwn = abs((Geometry1_Angle[Count6] -Geometry2_Angle[Count6]))
            if Anglebtwn > 180:
                #print 'Anglebtwn'
                Anglebtwn = 360 - Anglebtwn
            AngleBetween.pop(Count6)
            AngleBetween.insert(Count6, Anglebtwn)
            
    ###############################################################################################################################################################################
    #                    Get Max Fillet Radius
    ############################################################################################################################################################################### 
        print 'Get Max Fillet Radius'
        for Count6 in range(len(Vertice)-1):
            Count6 = Count6+1
            L1 = ((X11[Count6] - X12[Count6])**2 +(Y11[Count6] - Y12[Count6])**2)**.5
            L2 = ((X21[Count6] - X22[Count6])**2 +(Y21[Count6] - Y22[Count6])**2)**.5
            if L1 > L2:
                L = L2
            elif L1< L2:
                L = L1
            else:
                L = L1
            Theta = radians(AngleBetween[Count6])
            MaxRadius = (L * sin(Theta/2)) /(sin( radians(90)- Theta/2))
            MaxFilletRadius .pop(Count6)
            MaxFilletRadius .insert(Count6, MaxRadius)
            if MaxFilletRadius < RequestedFilletRadius:
                FilletRadius.pop(Count6)
                FilletRadius.insert(Count6, .5*MaxFilletRadius)
            elif MaxFilletRadius > RequestedFilletRadius:
                FilletRadius.pop(Count6)
                FilletRadius.insert(Count6, RequestedFilletRadius)
            else:
                FilletRadius.pop(Count6)
                FilletRadius.insert(Count6, RequestedFilletRadius)
                
    ###############################################################################################################################################################################
    #                    Print Combined Tables Portion
    ############################################################################################################################################################################### 
        print 'Step: Print combined tables portion'    
        myOutputFile = open('Combined_Data.txt','w+')
        m = 0
        
        for Count1 in range(len(Vertice)):
            new_line = ''
            n = 0

            for Count2 in range(len(Combined_List)):
                appendage = str(Combined_List[n][m]) + '~'
                new_line = new_line + appendage
                #print 'n: ' + str(n)
                n = n +1
            new_line = new_line + '\n'
            myOutputFile.write(new_line)
            m = m +1
    ###############################################################################################################################################################################
    #                    Create Fillet
    ############################################################################################################################################################################### 
        
        s1 = mdb.models[Model_Name].ConstrainedSketch(name='__edit__', 
            objectToCopy=mdb.models[Model_Name].sketches[Sketch_Name])
        g, v, d, c = s1.geometry, s1.vertices, s1.dimensions, s1.constraints

        s1.setPrimaryObject(option=STANDALONE)
        for Count7 in range(len(Vertice)-1):
            Count7 = Count7+1
            if 30 < AngleBetween[Count7]  and AngleBetween[Count7] < 135 and Geometry1[Count7] <> Geometry2[Count7]:
                g1= Geometry1[Count7]
                g2= Geometry2[Count7]
                #print 'g1: ' + str(g1) + ' g2: ' + str(g2)
                s1.FilletByRadius(radius=RequestedFilletRadius, curve1=g[g1], nearPoint1=(X1_mid[Count7] , Y1_mid[Count7]), curve2=g[g2], nearPoint2=(X2_mid[Count7] , Y2_mid[Count7]))
        mdb.models[Model_Name].sketches.changeKey(fromName='__edit__', toName= Sketch_Name)
        s1.unsetPrimaryObject()
     
def G_Define_Surfaces():
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
    import win32com.client                                    #Allows imports from "Active" Excel document
    
    ##Imports from excel
    xl = win32com.client.Dispatch("Excel.Application")        #Type xl. before using Excel procedures

    # Loop through all your models. 
    for i in range(2,7):                    
        # Assign a variable to your models root assembly
        a = mdb.models[str(xl.Cells(i,2).value)].rootAssembly
        session.viewports['Viewport: 1'].setValues(displayedObject=a)

        #PIN THREADED SECTION SURFACE
        #assign a variable to all of your pin edges
        s1 = a.instances['Pin-1'].edges
        #get all of the edges that are with in the bounding box
        side1Edges1 = s1.getByBoundingBox(xMax = xl.Cells(i,24).value, xMin = xl.Cells(i,25).value, yMax = xl.Cells(i,26).value, yMin = xl.Cells(i,27).value-.2)
        #set the surface to the bounding box
        a.Surface(side1Edges=side1Edges1, name='Pin Thread')
        
        #PIN SEAL SECTION SURFACE
        #assign a variable to all of your pin edges
        s1 = a.instances['Pin-1'].edges
        #get all of the edges that are with in the bounding box
        side1Edges1 = s1.getByBoundingBox(xMax = xl.Cells(i,20).value, xMin = xl.Cells(i,21).value, yMax = xl.Cells(i,22).value, yMin = xl.Cells(i,23).value)
        #set the surface to the bounding box
        a.Surface(side1Edges=side1Edges1, name='Pin Seal')
        
        #PIN BOUNDARY CONDITION SECTION SET
        #assign a variable to all of your pin edges
        s1 = a.instances['Pin-1'].edges
        #get all of the edges that are with in the bounding box
        edges1 = s1.getByBoundingBox(xMax = xl.Cells(i,16).value, xMin = xl.Cells(i,17).value, yMax = xl.Cells(i,18).value, yMin = xl.Cells(i,19).value)
        #set the surface to the bounding box
        a.Set(edges=edges1, name='Pin_BC')
        
        #BOX THREADED SECTION
        #assign a variable to all of your pin edges
        s1 = a.instances['Box-1'].edges
        #get all of the edges that are with in the bounding box
        side1Edges1 = s1.getByBoundingBox(xMax = xl.Cells(i,12).value, xMin = xl.Cells(i,13).value, yMax = xl.Cells(i,14).value, yMin = xl.Cells(i,15).value)
        #set the surface to the bounding box
        a.Surface(side1Edges=side1Edges1, name='Box Thread')

        #BOX BOUNDARY CONDITION SECTION SET
        #assign a variable to all of your pin edges
        s1 = a.instances['Box-1'].edges
        #get all of the edges that are with in the bounding box
        #box yMin is negative because is on the other side of the axis
        edges1 = s1.getByBoundingBox(xMax = xl.Cells(i,8).value, xMin = xl.Cells(i,9).value, yMax = xl.Cells(i,10).value, yMin = -xl.Cells(i,11).value)
        #set the surface to the bounding box
        a.Set(edges=edges1, name='Box_BC')

        
        a = mdb.models[str(xl.Cells(i,2).value)].rootAssembly

def H_SEEDandMESH():
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
    import win32com.client                                    #Allows imports from "Active" Excel document
    
    ##Imports from excel
    xl = win32com.client.Dispatch("Excel.Application")        #Type xl. before using Excel procedures

    #Set Mesh Size
    large_mesh = 0.05    
    pin_mesh = .004
    box_mesh = .005    
    
    # Loop through all your models. 
    for i in range(2,7):                    
        # Assign a variable to your models root assembly
        a = mdb.models[str(xl.Cells(i,2).value)].rootAssembly

        # session.viewports['Viewport: 1'].setValues(displayedObject=a)
        # session.viewports['Viewport: 1'].view.setValues(session.views['User-1'])
        # session.viewports['Viewport: 1'].assemblyDisplay.setValues(mesh=ON)
        # session.viewports['Viewport: 1'].assemblyDisplay.meshOptions.setValues(
            # meshTechnique=ON)


        e1 = a.instances['Box-1'].edges
        e2 = a.instances['Pin-1'].edges
        # Pick entire part and seed with the large mesh
        pickedEdges = e1.getByBoundingBox(xMax = 20, xMin = 0, yMax = 20, yMin = -1)+\
            e2.getByBoundingBox(xMax = 20, xMin = 0, yMax = 20, yMin = -1)
        a.seedEdgeBySize(edges=pickedEdges, size=large_mesh, deviationFactor=0.1, 
            constraint=FINER)

        # Pick Box Threads and Seal and seed with mesh
        e1 = a.instances['Box-1'].edges
        pickedEdges = e1.getByBoundingBox(xMax = xl.Cells(i,12).value, xMin = xl.Cells(i,13).value, yMax = xl.Cells(i,14).value, yMin = xl.Cells(i,15).value)
        a.seedEdgeBySize(edges=pickedEdges, size=box_mesh, deviationFactor=0.1, 
            constraint=FINER)
        
        
        # Pick Pin Threads and seed with mesh
        e1 = a.instances['Pin-1'].edges
        pickedEdges = e1.getByBoundingBox(xMax = xl.Cells(i,24).value, xMin = xl.Cells(i,25).value, yMax = xl.Cells(i,26).value, yMin = xl.Cells(i,27).value-.2)
        a.seedEdgeBySize(edges=pickedEdges, size=pin_mesh, deviationFactor=0.1, 
            constraint=FINER)

        # Pick Pin Threads and seed with mesh
        e1 = a.instances['Pin-1'].edges
        pickedEdges = e1.getByBoundingBox(xMax = xl.Cells(i,20).value, xMin = xl.Cells(i,21).value, yMax = xl.Cells(i,22).value, yMin = xl.Cells(i,23).value)
        a.seedEdgeBySize(edges=pickedEdges, size=pin_mesh, deviationFactor=0.1, 
            constraint=FINER)            
        # Set mesh controls to quad dominated
        f1 = a.instances['Box-1'].faces
        faces1 = f1.getByBoundingBox(xMax = 20, xMin = 0, yMax = 20, yMin = -1)
        f2 = a.instances['Pin-1'].faces
        faces2 = f2.getByBoundingBox(xMax = 20, xMin = 0, yMax = 20, yMin = -1)
        pickedRegions = faces1+faces2
    
        a.setMeshControls(regions=pickedRegions, elemShape=QUAD)
    
        a = mdb.models[str(xl.Cells(i,2).value)].rootAssembly
    
        partInstances =(a.instances['Box-1'], a.instances['Pin-1'], )
        a.generateMesh(regions=partInstances)


def FF_Material():
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
    import win32com.client                                    #Allows imports from "Active" Excel document
    
    ##Imports from excel
    xl = win32com.client.Dispatch("Excel.Application")        #Type xl. before using Excel procedures
    CellNum = xl.ActiveSheet.UsedRange.Rows.Count            #index of the last used cell in a column
    Type = ['Pin', 'Box']
    for j in Type:
        for i in range(2,7):    
            
            p = mdb.models[str(xl.Cells(i,2).value)].parts[j]

            f = p.faces
            #This gets all the faces in a 20 x 20 area
            faces = f.getByBoundingBox(xMax = 20, xMin = 0, yMax = 20, yMin = -1)
            region=regionToolset.Region(faces=faces)
            p.sectionAssignments[0].setValues(region=region)

def I_Create_Jobs():
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
    import win32com.client                                    #Allows imports from "Active" Excel document
    
    ##Imports from excel
    xl = win32com.client.Dispatch("Excel.Application")        #Type xl. before using Excel procedures

    for i in range(2,7):    

        mdb.Job(name=str(xl.Cells(i,2).value), model=str(xl.Cells(i,2).value), 
            description=str(xl.Cells(i,2).value), type=ANALYSIS, atTime=None, 
            waitMinutes=0, waitHours=0, queue=None, memory=90, 
            memoryUnits=PERCENTAGE, getMemoryFromAnalysis=True, 
            explicitPrecision=SINGLE, nodalOutputPrecision=SINGLE, echoPrint=OFF, 
            modelPrint=OFF, contactPrint=OFF, historyPrint=OFF, userSubroutine='', 
            scratch='', multiprocessingMode=DEFAULT, numCpus=1)

def J_Print_Viewports():
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
    import win32com.client                                    #Allows imports from "Active" Excel document
    
    ##Imports from excel
    xl = win32com.client.Dispatch("Excel.Application")        #Type xl. before using Excel procedures


        
    vp_list =  list(session.viewports.keys())
    vp1 = vp_list[0]
    vp2 = vp_list[1]
    vp3 = vp_list[2]
    vp4 = vp_list[3]
    print 'Viewport List: ' 
    print vp_list    
    
    von_mises_max = 125000
    hoop_max = 125000
    peeq_max = 0.005
    cpress_max = 150000
    
    working_directory_path = 'C:/FEA/QX-HT Full Analysis/'


    for i in range(2,20):  
    
        model_name = str(xl.Cells(i,2).value)
        odb_path = str(working_directory_path) + str(model_name) + '.odb'
        print 'ODB Path: ' + str(odb_path)        
        o1 = session.openOdb(name=str(odb_path))        
        odb = session.odbs[odb_path]   
    
        #set the odb for each viewport
        for i in vp_list:
            print  str(i) + ' odb updated'
            session.viewports[str(i)].makeCurrent()
            session.viewports[str(i)].setValues(displayedObject=odb)
            print  str(i) + ' odb updated'
            

        #set top left viewport for von mises
        session.viewports[str(vp1)].makeCurrent()
        session.viewports[str(vp1)].odbDisplay.setPrimaryVariable(
            variableLabel='S', outputPosition=INTEGRATION_POINT, refinement=(
            INVARIANT, 'Mises'), )
        session.viewports[str(vp1)].odbDisplay.contourOptions.setValues(
            maxAutoCompute=OFF, maxValue=von_mises_max, minValue=2836.03)
        session.viewports[str(vp1)].view.setValues(session.views['User-1'])
        session.viewports[str(vp1)].odbDisplay.display.setValues(plotState=(
            CONTOURS_ON_DEF, ))
            
        #set top right viewport for hoop        
        session.viewports[str(vp2)].makeCurrent()
        session.viewports[str(vp2)].odbDisplay.setPrimaryVariable(
            variableLabel='S', outputPosition=INTEGRATION_POINT, refinement=(
            COMPONENT, 'S33'), )
        session.viewports[str(vp2)].odbDisplay.contourOptions.setValues(
            maxAutoCompute=OFF, maxValue=hoop_max, minValue=-hoop_max)
        session.viewports[str(vp2)].view.setValues(session.views['User-1'])
        session.viewports[str(vp2)].odbDisplay.display.setValues(plotState=(
            CONTOURS_ON_DEF, ))
            
        #set bottom left viewport for peeq        
        session.viewports[str(vp3)].makeCurrent()
        session.viewports[str(vp3)].odbDisplay.setPrimaryVariable(
            variableLabel='PEEQ', outputPosition=INTEGRATION_POINT, )
        session.viewports[str(vp3)].odbDisplay.contourOptions.setValues(
            maxAutoCompute=OFF, maxValue=peeq_max, minValue=0)
        session.viewports[str(vp3)].view.setValues(session.views['User-1'])
        session.viewports[str(vp3)].odbDisplay.display.setValues(plotState=(
            CONTOURS_ON_DEF, ))
            
        #set bottom right viewport for cpress
        session.viewports[str(vp4)].makeCurrent()
        session.viewports[str(vp4)].odbDisplay.commonOptions.setValues(
            renderStyle=SHADED, )
        leaf = dgo.LeafFromPartInstance(partInstanceName=('BOX-1', ))
        session.viewports[str(vp4)].odbDisplay.displayGroup.remove(leaf=leaf)


        session.viewports[str(vp4)].odbDisplay.contourOptions.setValues(
            tickmarkCurveColor='#000000', maxAutoCompute=OFF, maxValue=cpress_max, 
            minValue=24340.2)
        session.viewports[str(vp4)].odbDisplay.contourOptions.setValues(
            tickmarkPlots=ON)
        session.viewports[str(vp4)].odbDisplay.setPrimaryVariable(
            variableLabel='CPRESS', outputPosition=ELEMENT_NODAL, )

        session.viewports[str(vp4)].odbDisplay.display.setValues(plotState=(
            CONTOURS_ON_DEF, ))
        session.viewports[str(vp4)].view.setValues(session.views['User-1'])
        #session.viewports[str(vp4)].view.setValues(nearPlane=9.13127, farPlane=13.5706, width=1.37142, height=0.566195, viewOffsetX=xl.Cells(i,21).value, viewOffsetY=xl.Cells(i,23).value)
		
        #Print all viewports
        session.printToFile(fileName=str(model_name) + ' FEA', format=PNG, canvasObjects=(
            session.viewports[str(vp1)], session.viewports[str(vp2)], 
            session.viewports[str(vp3)], session.viewports[str(vp4)]))

        #Close odb
        #session.odbs[str(odb_path)].close()




