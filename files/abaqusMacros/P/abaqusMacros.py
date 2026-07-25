########################################################################################################################################################

#Scripts for creating and assigning material, step creation, setting the view, adjusting font, applying loads/importing loading data, 
#creating interactions and interaction properties, and creating pressure penetration

######################################
###						           ###
###    Written by: Ryan Hodgins    ###
###						           ###
######################################

########################################################################################################################################################
from abaqus import *
from abaqusConstants import *
import __main__
import step
import visualization
import xyPlot
import displayGroupMdbToolset as dgm
import displayGroupOdbToolset as dgo
import win32com.client		#Allows imports from "Active" Excel document
from collections import defaultdict
xl = win32com.client.Dispatch("Excel.Application")
#import sys
#sys.path.append(r'C:\SIMULIA\Abaqus\6.12-3\tools\SMApy\Lib\lib-tk')
#import Tkinter
#import tkMessageBox as tk

def B__Create_Material():
	#Request user inputs for Connection, size, and weight
	conntype = getInput('Connection Type: ')

	not_decimal = True

	while not_decimal:
		found_slash = False
		size = getInput('Outside Diameter(in): ')
		charcount = len(size)

		for i in range(0, charcount):
			if size[i] == '/':
				getWarningReply('Please enter size as a decimal', "CANCEL") #Cancel button will not show up
				found_slash = True
				break
			elif size[i] == '.':
			#Convert size input from string to list, then back to string because strings are immutable whereas lists are not
				sizechars = list(size)
				sizechars[i] = 'P'
				size = "".join(sizechars)

		if not found_slash:
			not_decimal = False

	weight = getInput('Weight (lb): ')

	#Material Creation
	model = mdb.models.keys()[0]
	outside_range = True  #While loop to repeat input request until user enters an acceptable value.
	while outside_range:
		material = getInput('Material?\n\n1 for P-110\n\n2 for I-80\n\n3 for J-55\n\n4 for Q-125')
		if material == '1':
			mdb.models[model].Material(name='P-110')
			mdb.models[model].materials['P-110'].Elastic(table=((30000000.0, 0.3), ))
			mdb.models[model].materials['P-110'].Plastic(table=((100473.1, 0.0), (
				102947.7, 7.73615e-05), (105412.6, 0.000170728), (110505.5, 0.000578004), (
				112572.1, 0.000874124), (114651.3, 0.001328723), (116726.1, 0.001989506), (
				118865.3, 0.003127101), (121140.5, 0.005088999), (122548.7, 0.006745395), (
				125142.1, 0.010646338), (127979.0, 0.016141818), (130720.0, 0.022479825), (
				132786.7, 0.028164075), (134699.2, 0.034055993), (136000.5, 0.038682456), (
				137307.5, 0.044243551), (137807.9, 0.046495517), (138654.5, 0.050881691), (
				139517.6, 0.055992652), (140566.6, 0.063125487), (140883.6, 0.065334901), (
				141612.3, 0.072453384), (142250.6, 0.081627384), (142539.9, 0.090890538)))
			outside_range = False
		elif material == '2':
			mdb.models[model].Material(name='I-80')
			mdb.models[model].materials['I-80'].Elastic(table=((30000000.0, 0.3), ))
			mdb.models[model].materials['I-80'].Plastic(table=((85250.0, 0.0), (
				89144.0, 0.014), (93885.0, 0.021), (99883.0, 0.032), (104217.0, 0.043), (
				107533.0, 0.055), (109976.0, 0.066), (111707.0, 0.077), (113103.0, 0.088), 
				(114256.0, 0.099), (115479.0, 0.109), (117130.0, 0.127)))
			outside_range = False
		elif material == '3':
			mdb.models[model].Material(name='J-55')
			mdb.models[model].materials['J-55'].Elastic(table=((30000000.0, 0.3), ))
			mdb.models[model].materials['J-55'].Plastic(table=((56641.0, 0.0), (
				60012.0, 9e-05), (60774.0, 0.00027), (59035.0, 0.00052), (58797.0, 
				0.00073), (58559.0, 0.00094), (57869.0, 0.00116), (57881.0, 0.00136), (
				57091.0, 0.00158), (57103.0, 0.00178), (56613.0, 0.002), (56624.0, 0.0022), 
				(56636.0, 0.0024), (56146.0, 0.00261), (56157.0, 0.00281), (56168.0, 
				0.00301), (56179.0, 0.00321), (56347.0, 0.00619)))
			outside_range = False
		elif material == '4':
			mdb.models[model].Material(name='Q-125')
			mdb.models[model].materials['Q-125'].Elastic(table=((30000000.0, 0.3), ))
			mdb.models[model].materials['Q-125'].Plastic(table=((115000, 0.0), (
				120540.3, 0.0000666173414860709), (130673.2, 0.000309607706748417), (140844.1, 0.00128600906443848), (140945.7, 
				0.00131240336335226), (142745.0, 0.00175671825897578), (143779.2, 0.00221748786501674), (144667.4, 0.00270142433948302), (
				145369.0, 0.00317584531744502), (145877, 0.00366511361124562), (146357.0, 0.00414291338333622), (146894.2, 0.00461213003632251), 
				(147356.1, 0.00509943167424629), (147683.9, 0.00558445330372281), (148184.0, 0.00607598238036932), (148180.3, 
				0.00655715280801751), (148600.5, 0.00706076688788915), (148802.6, 0.0075644598302633), (149114.6, 0.00801091189874437), (148725.9, 0.00849122233568163)))
			outside_range = False
		elif material == None:
			outside_range = False
			return
		else:
			getWarningReply('Input is outside range. Please enter 1, 2, 3, or 4.', "OKAY") #Okay button will not show up

	#Renames model to common format
	model2 = conntype+'_'+size+'_'+weight+'lb_'+'description'
	model2 = model.upper()    #Capitalizes model
	mdb.models.changeKey(fromName=model, toName=model2)

###Macro for changing view
def A__Set_View():
	session.View(name='User-1', nearPlane=12.28, farPlane=36.841, width=3.6572, 
		height=5.3095, projection=PERSPECTIVE, cameraPosition=(3.4831, 4.2139, 
		24.561), cameraUpVector=(1, -4.3711e-008, 0), cameraTarget=(3.4831, 4.2139, 
		0), viewOffsetX=0, viewOffsetY=0, autoFit=ON)

def D__Create_Steps():
	import step
	input = False
	model = mdb.models.keys()[0]
	a = mdb.models[model].rootAssembly
	while input == False:
		answer = getInput('Do you want to create multiple Load Steps or only a MakeUp step?\n\nEnter "1" for multiple Load Steps\n\nEnter "2" for only MakeUp')
		print answer
		if answer == '1':
			try:
				##Imports from excel
				xl = win32com.client.Dispatch("Excel.Application")		#Type xl. before using Excel procedures
				if xl.ActiveSheet == None: 								#Tests to see if an Excel sheet is open
					getWarningReply('Need to have Excel sheet for desired Test open. Please open and rerun macro.', 'OK')
					return
					
				#Create arrays of load point values and step names from active excel sheet
				CellNum = xl.ActiveSheet.UsedRange.Rows.Count			#index of the last used cell in a column
				Steps = []		#Create empty lists
				StepsList = []
				tracker = defaultdict(int)
				for i in range(3,CellNum + 1):
					StepsList.append(int(xl.Cells(i,1).value))				#Format for cell lookup is (ROW,COLUMN) 		
				for cell in StepsList:									#Creates naming convention for steps
					tracker[cell]+=1
					Steps.append('LP'+str(cell)+'_'+str(tracker[cell]))
				# Deletes previously created steps if applicable
				input = False
				if len(mdb.models[model].steps.keys()) > 1:
					while input == False:
						answer = getInput('Steps have been previously created. Would you like to delete them before you create new ones?\n\nEnter "1" for Yes\n\nEnter "2" for No')
						if answer == '1':
							for elem in mdb.models[model].steps.keys()[1:len(mdb.models[model].steps.keys())+1]:
								del mdb.models[model].steps[elem]
							input = True
						elif answer == '2':
							input = True
							pass
						else:
							getWarningReply('Invalid answer. Please enter 1 or 2.', 'OK')
							pass
				#Create MakeUp Step
				mdb.models[model].StaticStep(name='MakeUp', previous='Initial', 		
					stabilizationMagnitude=0.0002, stabilizationMethod=DISSIPATED_ENERGY_FRACTION, 
					continueDampingFactors=False, adaptiveDampingRatio=0.05, initialInc=0.001, 
					minInc=1e-08, maxInc=0.05, maxNumInc=1000, matrixSolver=DIRECT, 
					matrixStorage=UNSYMMETRIC, nlgeom=ON)
				#Create LP Steps
				for StepName in Steps:
					PreviousStep = mdb.models[mdb.models.keys()[0]].steps.keys()[-1]
					mdb.models[model].StaticStep(name=StepName, previous=PreviousStep, 
						stabilizationMagnitude=0.0002, stabilizationMethod=DISSIPATED_ENERGY_FRACTION,
						continueDampingFactors=False, adaptiveDampingRatio=0.05, initialInc=0.1, 
						minInc=1e-08, maxInc=0.5, maxNumInc=1000, matrixSolver=DIRECT, 
						matrixStorage=UNSYMMETRIC, nlgeom=ON)
				input = True

			except:
				getWarningReply('Check active Excel sheet. You may have the wrong sheet active or it may be incorrectly formatted.', 'OK')
				input = True
				return
		elif answer == '2':
			mdb.models[model].StaticStep(name='MakeUp', previous='Initial', 		
				stabilizationMagnitude=0.0002, stabilizationMethod=DISSIPATED_ENERGY_FRACTION, 
				continueDampingFactors=False, adaptiveDampingRatio=0.05, initialInc=0.001, 
				minInc=1e-08, maxInc=0.05, maxNumInc=1000, matrixSolver=DIRECT, 
				matrixStorage=UNSYMMETRIC, nlgeom=ON)
			input = True
		elif answer == None:
			break
		else:
			getWarningReply('Invalid answer. Please enter 1 or 2.', 'OK')
			pass

def E__Create_Loads():
	model = mdb.models.keys()[0]
	if len(mdb.models[model].steps.keys()) > 2:
		xl = win32com.client.Dispatch("Excel.Application")	#Type xl. before using Excel procedures
		if xl.ActiveSheet == None: 								#Tests to see if an Excel sheet is open
			getWarningReply('Need to have Excel sheet for desired Test open. Please open and rerun macro.', 'OK')
			return

		a = mdb.models[model].rootAssembly
		if 'AXIAL' and 'PINT' and 'PEXT' not in a.surfaces:	#checks for correct surfaces
			getWarningReply('Need the surfaces named: "AXIAL" for axial pressure (Tension/Compression), "PINT" for internal pressure, and "PEXT" for external pressure. Please create/rename and rerun macro.', 'OK')
			return

			#Create arrays of load point values and step names from active excel sheet
		CellNum = xl.ActiveSheet.UsedRange.Rows.Count			#index of the last used cell in a column
		Steps = []		#Create empty lists
		Loads = []
		PInt = []
		PExt = []
		StepsList = []
		tracker = defaultdict(int)
		for i in range(3,CellNum + 1):
			StepsList.append(int(xl.Cells(i,1).value))				#Format for cell lookup is (ROW,COLUMN) 		
			Loads.append(-xl.Cells(i,3).value) 
			if int(xl.Cells(i,2)) >= 0:								#If pressure is positive, will designate value as internal pressure and set external pressure as 0
				PInt.append(xl.Cells(i,2).value)
				PExt.append(0)
			else:
				PInt.append(0)									#If pressure is negative, will designate absolute value as external pressure and set internal pressure as 0
				PExt.append(-xl.Cells(i,2).value)
		for cell in StepsList:									#Creates naming convention for steps
			tracker[cell]+=1
			Steps.append('LP'+str(cell)+'_'+str(tracker[cell]))

		# Beginning of Load Creation
		# Axial Pressure
		for x in Loads:
			if x != 0:
				iAXIAL = Loads.index(x)
				mdb.models[model].Pressure(name='Axial', createStepName=Steps[iAXIAL], 
					region=a.surfaces['AXIAL'], distributionType=UNIFORM, field='', 
					magnitude=Loads[0], amplitude=UNSET)
				break
		# Internal Pressure
		for x in PInt:
			if x != 0:
				iINT = PInt.index(x)
				mdb.models[model].Pressure(name='InternalPressure', createStepName=Steps[iINT], 
					region=a.surfaces['PINT'], distributionType=UNIFORM, field='', 
					magnitude=PInt[iINT], amplitude=UNSET)
				break
		# External Pressure
		for x in PExt:
			if x != 0:
				iEXT = PExt.index(x)
				mdb.models[model].Pressure(name='ExternalPressure', createStepName=Steps[iEXT], 
					region=a.surfaces['PEXT'], distributionType=UNIFORM, field='', 
					magnitude=PExt[iEXT], amplitude=UNSET)
				break

		# Beginning of Load and Pressure edit for each Load Point
		if sum(Loads) != 0:
			for stepName in Steps[iAXIAL:(len(Steps)+1)]:
				i = Steps.index(stepName)
				mdb.models[model].loads['Axial'].setValuesInStep(stepName=stepName, 
					magnitude=Loads[i])
		if sum(PInt) != 0:
			for stepName in Steps[iINT:(len(Steps)+1)]:
				i = Steps.index(stepName)
				mdb.models[model].loads['InternalPressure'].setValuesInStep(stepName=stepName, 
					magnitude=PInt[i])
		if sum(PExt) != 0:
			for stepName in Steps[iEXT:(len(Steps)+1)]:
				i = Steps.index(stepName)
				mdb.models[model].loads['ExternalPressure'].setValuesInStep(stepName=stepName, 
					magnitude=PExt[i])
	else:
		getWarningReply('Please make sure you have the correct excel sheet active and formatted correctly. Make sure you have already created your desired steps.\n If you are only doing MakeUp, you do not need to run this macro', 'OK')
		return

def C__Assign_Material():							#### Creates section with previously created material and assigns it to the Pin and Box
	model = mdb.models.keys()[0]

	if len(mdb.models[model].materials.keys()) == 0:
		getWarningReply('No material available. Please run "Create_Material" macro first.', 'OK')
		return

	matl = mdb.models[model].materials.keys()[0]
	b=mdb.models[model].parts.keys()[0]      ####Creates individual sets for Pin and Box to apply material section
	c=mdb.models[model].parts.keys()[1]
	f1 = mdb.models[model].parts[b].faces
	f2 = mdb.models[model].parts[c].faces
	faces1 = f1.getSequenceFromMask(mask=('[#1 ]', ), )
	faces2 = f2.getSequenceFromMask(mask=('[#1 ]', ), )
	mdb.models[model].parts[b].Set(faces=faces1, name=b+'_Set')
	mdb.models[model].parts[c].Set(faces=faces2, name=c+'_Set')

	SectionName = mdb.models[model].materials.keys()[0] + ' Section'     ###Creates section with created material
	mdb.models[model].HomogeneousSolidSection(
		name=SectionName, material=matl, thickness=None)

	if len(mdb.models[model].parts[b].sectionAssignments) != 0:
		input = False
		while input == False:
			answer = getInput('It looks like you already have created Section Assignments. Would you like to delete previous Section Assignments and create a new one?\n\nEnter "1" for Yes\n\nEnter "2" for No')
			if answer == '1':
				del mdb.models[model].parts[b].sectionAssignments[0:len(mdb.models[model].parts[b].sectionAssignments)]
				input = True
			elif answer == '2':
				input = True
				pass
			elif answer == None:
				return
			else:
				getWarningReply('Invalid answer. Please enter 1 or 2.', 'OK')

	Part1 = mdb.models[model].parts[b]										####Assigns section to Pin and Box
	region1 = Part1.sets[b+'_Set']
	Part1.SectionAssignment(region=region1, sectionName=SectionName, offset=0.0, 
		offsetType=MIDDLE_SURFACE, offsetField='', 
		thicknessAssignment=FROM_SECTION)
		
	if len(mdb.models[model].parts[c].sectionAssignments) != 0:
		if answer == '1':
			del mdb.models[model].parts[c].sectionAssignments[0:len(mdb.models[model].parts[c].sectionAssignments)]
		elif answer == '2':
			pass

	Part2 = mdb.models[model].parts[c]
	region2 = Part2.sets[c+'_Set']
	Part2.SectionAssignment(region=region2, sectionName=SectionName, offset=0.0, 
		offsetType=MIDDLE_SURFACE, offsetField='', 
		thicknessAssignment=FROM_SECTION)

def F__Create_Boundary_Conditions():
	model = mdb.models.keys()[0]
	a = mdb.models[model].rootAssembly
	if 'FixPin' and 'FixBox' not in a.sets.keys():
		getWarningReply('Need the sets named: "FixPin" and "FixBox". Please create/rename and rerun macro.', "OKAY")
		return

	region1 = a.sets['FixPin']
	mdb.models[model].DisplacementBC(name='FixPin', 
		createStepName='MakeUp', region=region1, u1=UNSET, u2=0.0, ur3=UNSET, 
		amplitude=UNSET, fixed=OFF, distributionType=UNIFORM, fieldName='', 
		localCsys=None)
	region2 = a.sets['FixBox']
	mdb.models[model].DisplacementBC(name='FixBox', 
		createStepName='MakeUp', region=region2, u1=UNSET, u2=0.0, ur3=UNSET, 
		amplitude=UNSET, fixed=OFF, distributionType=UNIFORM, fieldName='', 
		localCsys=None)
	mdb.models[model].boundaryConditions['FixPin'].deactivate(mdb.models[model].steps.keys()[2])

def G__Create_Int_Property():
	model = mdb.models.keys()[0]
	if 'PinThreads' and 'BoxThreads' not in mdb.models[model].rootAssembly.surfaces:
			getWarningReply('Need the surfaces named: "PinThreads" for pin thread surface and "BoxThreads" for box thread surface. Please create/rename and rerun macro.', "OKAY")
			return
	if len(mdb.models[model].steps.keys()) > 2:
		xl = win32com.client.Dispatch("Excel.Application")		#Recreating step array for interaction property reference
		if xl.ActiveSheet == None: 								
			getWarningReply('Need to have Excel sheet for desired Test open. Please open and rerun macro.', 'OK')
			return

		CellNum = xl.ActiveSheet.UsedRange.Rows.Count
		Steps = []
		StepsList = []
		tracker = defaultdict(int)

		for i in range(3,CellNum + 1):
			StepsList.append(int(xl.Cells(i,1).value))
		for cell in StepsList:									#Creates naming convention for steps
			tracker[cell]+=1
			Steps.append('LP'+str(cell)+'_'+str(tracker[cell]))

		#Creates frictionless MakeUp Interaction Property
		MUPropName = getInput('What would you like to name the MakeUp Interaction Property?')

		mdb.models[model].ContactProperty(MUPropName)
		mdb.models[model].interactionProperties[MUPropName].TangentialBehavior(
			formulation=FRICTIONLESS)
		mdb.models[model].interactionProperties[MUPropName].NormalBehavior(
			pressureOverclosure=HARD, allowSeparation=ON, contactStiffness=DEFAULT, 
			contactStiffnessScaleFactor=1.0, clearanceAtZeroContactPressure=0.0, 
			constraintEnforcementMethod=PENALTY)

		#Creates Load Steps Int Property with friction coefficient of 0.04
		LSPropName = getInput('What would you like to name the Load Step Interaction Property?')	
		mdb.models[model].ContactProperty(LSPropName)
		mdb.models[model].interactionProperties[LSPropName].TangentialBehavior(
			formulation=PENALTY, directionality=ISOTROPIC, slipRateDependency=OFF, 
			pressureDependency=OFF, temperatureDependency=OFF, dependencies=0, table=((
			0.04, ), ), shearStressLimit=None, maximumElasticSlip=FRACTION, 
			fraction=0.005, elasticSlipStiffness=None)
		mdb.models[model].interactionProperties[LSPropName].NormalBehavior(
			pressureOverclosure=HARD, allowSeparation=ON, 
			constraintEnforcementMethod=PENALTY)

		#Create Interaction and apply frictionless property to makeup and friction coeff of 0.04 to all other Load Points
		IntName = getInput('What would you like to name the Surface to Surface Interaction?')

		a = mdb.models[model].rootAssembly
		region1=a.surfaces['BoxThreads']
		region2=a.surfaces['PinThreads']
		mdb.models[model].SurfaceToSurfaceContactStd(name=IntName, 
			createStepName='MakeUp', master=region1, slave=region2, sliding=FINITE, 
			thickness=ON, interactionProperty=MUPropName, interferenceType=SHRINK_FIT, 
			adjustMethod=NONE, initialClearance=OMIT, datumAxis=None, 
			clearanceRegion=None)

		mdb.models[model].interactions[IntName].setValuesInStep(
			stepName=Steps[0], interactionProperty=LSPropName)
	else:
		MUPropName = getInput('What would you like to name the MakeUp Interaction Property?')

		mdb.models[model].ContactProperty(MUPropName)
		mdb.models[model].interactionProperties[MUPropName].TangentialBehavior(
			formulation=FRICTIONLESS)
		mdb.models[model].interactionProperties[MUPropName].NormalBehavior(
			pressureOverclosure=HARD, allowSeparation=ON, contactStiffness=DEFAULT, 
			contactStiffnessScaleFactor=1.0, clearanceAtZeroContactPressure=0.0, 
			constraintEnforcementMethod=PENALTY)

		#Creates frictionless Surface to Surface Interaction for MakeUp step alone
		IntName = getInput('What would you like to name the Surface to Surface Interaction?')
		a = mdb.models[model].rootAssembly
		region1=a.surfaces['BoxThreads']
		region2=a.surfaces['PinThreads']
		mdb.models[model].SurfaceToSurfaceContactStd(name=IntName, 
			createStepName='MakeUp', master=region1, slave=region2, sliding=FINITE, 
			thickness=ON, interactionProperty=MUPropName, interferenceType=SHRINK_FIT, 
			adjustMethod=NONE, initialClearance=OMIT, datumAxis=None, 
			clearanceRegion=None)


def H__Create_PressPen():
	if xl.ActiveSheet == None: 								
		getWarningReply('Need to have Excel sheet for desired Test open. Please open and rerun macro.', 'OK')
		return

	model = mdb.models.keys()[0]
	IntName = mdb.models[model].interactions.keys()[0]
	CellNum = xl.ActiveSheet.UsedRange.Rows.Count			#index of the last used cell in a column
	Steps = []		#Create empty lists
	PInt = []
	PExt = []
	StepsList = []
	tracker = defaultdict(int)
	for i in range(3,CellNum + 1):
		StepsList.append(int(xl.Cells(i,1).value))				#Format for cell lookup is (ROW,COLUMN) 		
		if int(xl.Cells(i,2)) >= 0:								#If pressure is positive, will designate value as internal pressure and set external pressure as 0
			PInt.append(xl.Cells(i,2).value)
			PExt.append(0)
		else:
			PInt.append(0)									#If pressure is negative, will designate absolute value as external pressure and set internal pressure as 0
			PExt.append(-xl.Cells(i,2).value)
	for cell in StepsList:									#Creates naming convention for steps
		tracker[cell]+=1
		Steps.append('LP'+str(cell)+'_'+str(tracker[cell]))
	for elem in PInt:
		if elem != 0:
			iInt = PInt.index(elem)
			break
	for elem in PExt:
		if elem != 0:
			iExt = PExt.index(elem)
			break

	a = mdb.models[model].rootAssembly	
	if sum(PInt) != 0:
		if 'IntPressPenMaster' and 'IntPressPenSlave' not in a.sets.keys():
			getWarningReply('Missing the sets named: "IntPressPenMaster" and "IntPressPenSlave". Please create/rename and rerun macro.', "OKAY")
			return
		masterPoint1 = a.sets['IntPressPenMaster']
		masterPoints =(masterPoint1, )
		slavePoint1 = a.sets['IntPressPenSlave']
		slavePoints =(slavePoint1, )
		mdb.models[model].PressurePenetration(name='IntPressPen',       ###Creates Internal Pressure Penetration
			createStepName=Steps[iInt], contactInteraction=IntName, 
			masterPoints=masterPoints, slavePoints=slavePoints, criticalPressure=(0.0, 
			), penetrationPressure=(PInt[iInt], ), amplitude='')

		for StepName in Steps[iInt:(len(Steps)+1)]:			###Sets pressure values into pressure penetration interaction
			i = Steps.index(StepName)
			mdb.models[model].interactions['IntPressPen'].setValuesInStep(
				stepName=StepName, penetrationPressure=(PInt[i], ))

	if sum(PExt) != 0:
		if 'ExtPressPenMaster' and 'ExtPressPenSlave' not in a.sets.keys():
			getWarningReply('Missing the sets named: "ExtPressPenMaster" and "ExtPressPenSlave". Please create/rename and rerun macro.', "OKAY")
			return
		masterPoint2 = a.sets['ExtPressPenMaster']
		masterPoints2 =(masterPoint2, )
		slavePoint2 = a.sets['ExtPressPenSlave']
		slavePoints2 =(slavePoint2, )
		mdb.models[model].PressurePenetration(name='ExtPressPen',       ###Creates External Pressure Penetration
			createStepName=Steps[iExt], contactInteraction=IntName, 
			masterPoints=masterPoints2, slavePoints=slavePoints2, criticalPressure=(0.0, 
			), penetrationPressure=(PExt[iExt], ), amplitude='')

		for StepName in Steps[iExt:(len(Steps)+1)]:			###Sets pressure values into pressure penetration interaction
			i = Steps.index(StepName)
			mdb.models[model].interactions['ExtPressPen'].setValuesInStep(
				stepName=StepName, penetrationPressure=(PExt[i], ))

def I__Mesh():
	model = mdb.models.keys()[0]
	session.viewports['Viewport: 1'].assemblyDisplay.setValues(mesh=ON,               ###Select Mesh Module
		interactions=OFF, constraints=OFF, connectors=OFF, engineeringFeatures=OFF)
	session.viewports['Viewport: 1'].assemblyDisplay.meshOptions.setValues(
		meshTechnique=ON)                                                 

	a = mdb.models[model].rootAssembly  ###Select Seeding and size
	pickedEdges = a.surfaces['AXIAL'].edges + a.surfaces['BoxEnd'].edges + a.surfaces['PINT'].edges + a.surfaces['PEXT'].edges
	a.seedEdgeBySize(edges=pickedEdges, size=0.05, deviationFactor=0.1, 
		constraint=FINER)
	
	a = mdb.models[model].rootAssembly  ###Select Seeding and size for Pin Threads
	pickedEdges = a.surfaces['PinThreads'].edges
	a.seedEdgeBySize(edges=pickedEdges, size=0.007, deviationFactor=0.1, 
		constraint=FINER)

	a = mdb.models[model].rootAssembly  ###Select Seeding and size for Box Threads
	pickedEdges = a.surfaces['BoxThreads'].edges
	a.seedEdgeBySize(edges=pickedEdges, size=0.008, deviationFactor=0.1, 
		constraint=FINER)

	f1 = a.instances[a.instances.keys()[0]].faces                ###Select Mesh Styling
	faces1 = f1.getSequenceFromMask(mask=('[#1 ]', ), )
	f2 = a.instances[a.instances.keys()[1]].faces
	faces2 = f2.getSequenceFromMask(mask=('[#1 ]', ), )
	pickedRegions = faces1+faces2
	a.setMeshControls(regions=pickedRegions, elemShape=QUAD)
	partInstances =(a.instances[a.instances.keys()[0]], a.instances[a.instances.keys()[1]], )
	a.generateMesh(regions=partInstances)
	
def J__Format_Plots():									### Sets up viewport format for post-processing step printing. Sets font, background, free edges, etc	
	session.graphicsOptions.setValues(backgroundStyle=SOLID, 
		backgroundColor='#FFFFFF')
	inp = getInput('What size font do you want to use?', '14')
	if inp == None:
		return
	for key in session.viewports.keys():
		session.viewports[key].odbDisplay.display.setValues(plotState=(
			CONTOURS_ON_DEF, ))
		session.viewports[key].odbDisplay.commonOptions.setValues(
			visibleEdges=FREE)
		session.viewports[key].viewportAnnotationOptions.setValues(
		triadFont='-*-verdana-bold-r-normal-*-*-' + inp + '0-*-*-p-*-*-*',
		legendFont='-*-verdana-bold-r-normal-*-*-' + inp + '0-*-*-p-*-*-*',
		titleFont='-*-verdana-bold-r-normal-*-*-' + inp + '0-*-*-p-*-*-*',
		stateFont='-*-verdana-bold-r-normal-*-*-' + inp + '0-*-*-p-*-*-*')
		session.viewports[key].viewportAnnotationOptions.setValues(title=OFF, 
			statePosition=(15, 95), compass=OFF)

def K__ViewCPress():
	import visualization
	import xyPlot
	import displayGroupOdbToolset as dgo

	currentViewport = session.currentViewportName
	session.viewports[currentViewport].view.setValues(nearPlane=21.8619,
		farPlane=33.3082, width=12.5514, height=6.59917, viewOffsetX=1.41986,
		viewOffsetY=-0.497329)
	leaf = dgo.LeafFromPartInstance(partInstanceName=('Box-1', ))
	session.viewports[currentViewport].odbDisplay.displayGroup.remove(leaf=leaf)
	session.viewports[currentViewport].odbDisplay.setPrimaryVariable(
		variableLabel='CPRESS', outputPosition=ELEMENT_NODAL, )
	session.viewports[currentViewport].odbDisplay.display.setValues(
		plotState=CONTOURS_ON_DEF)
	session.viewports[currentViewport].odbDisplay.contourOptions.setValues(
		maxValue=161264, minValue=0)

def L__ProcessXYData():
	import visualization
	import xyPlot
	import displayGroupOdbToolset as dgo

	inp=getInput('Enter path #:','Path-1')
	if inp == None:
		return
	lastStep = len(session.odbs[session.odbs.keys()[0]].steps)	
	for x in range(0,lastStep) :
		step_name = session.odbs[session.odbs.keys()[0]].steps.keys()[x] 		#Name/key of the current step
		lastFrame = session.odbs[session.odbs.keys()[0]].steps[step_name].frames[-1].incrementNumber
		currentViewport = session.currentViewportName
		session.viewports[currentViewport].odbDisplay.setFrame(step=x, frame=lastFrame)

		if len(session.paths.keys()) == 0:
			reply = getWarningReply(message = 'Must create a path', buttons = ('OK'))	#Message box. Does not include "Ok" button
			break

		if inp <> []:
			pth = session.paths[inp]
		if inp == []:
			pthKey = session.paths.keys()[0]
			pth = session.paths[pthKey]

		session.XYDataFromPath(name= step_name , path=pth, includeIntersections=False, shape=DEFORMED, labelType=TRUE_DISTANCE)

def M__PrintAllSteps():			
	import visualization
	import xyPlot
	import displayGroupOdbToolset as dgo
	import time

    #Count the Steps
	lastStep = len(session.odbs[session.odbs.keys()[0]].steps)		#last step of the current output database'

	for x in range(0,lastStep):
		for key in session.viewports.keys():
			vp = session.viewports[key]
			vp.makeCurrent
			step_name = session.odbs[session.odbs.keys()[0]].steps.keys()[x] 		#Name/key of the current step
			lastFrame = session.odbs[session.odbs.keys()[0]].steps[step_name].frames[-1].incrementNumber	#index of the last frame of the current step
			vp.odbDisplay.setFrame(step=x, frame=lastFrame)
			od = vp.odbDisplay

			Var = od.primaryVariable[0]				#This is the symbol representation of the current load type
			var2 = od.primaryVariable[5]

			session.pngOptions.setValues(imageSize = (2272,1704))
			session.printOptions.setValues(rendition=COLOR, vpDecorations=OFF, vpBackground=OFF)

			#Saves a file with a unique name for each step of a load type
			session.printToFile(fileName= 'Stress' + '_' + Var + '_' + var2 + '_Step_' + str(x+1) + '_(' + step_name + ')', format = PNG, canvasObjects= (vp,))