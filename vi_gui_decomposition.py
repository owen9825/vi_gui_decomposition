#!/usr/bin/env python
from gimpfu import *
import os
from ctypes import *
import comtypes.client
import math

def viGuiDecomposition(imgFilename,toolpath,middleLayer,targetVI):
    #"""Decompose an image"""
	
	img = pdb.gimp_file_load(imgFilename,imgFilename)
	imgWidth = pdb.gimp_image_width(img)
	imgHeight = pdb.gimp_image_height(img)
	desiredLayer = pdb.gimp_image_get_layer_by_name(img,middleLayer)
	numLayers, layerIds = pdb.gimp_image_get_layers(img)
	
	#initialise COM section
	print "Initialising COM"
	comtypes.CoInitialize()
	TypeLibPath = "C:/Program Files/National Instruments/LabVIEW 2011/resource/labview.tlb"
	print comtypes.client.gen_dir
	comtypes.client.gen_dir = None
	comtypes.client.GetModule(TypeLibPath)
	
	try:
		Application = comtypes.client.CreateObject("LabVIEW.Application",None,None,comtypes.gen.LabVIEW._Application)
		
		#prepare parameters for insertTextVI.vi
		Text = "new title"
		#magenta
		TextColour = 16728787
		FontName = "Bradley Hand ITC"
		FontSize = 18
		TargetPath = targetVI
		LeftEdge = 200
		TopEdge = 100
		ParameterNames = ["Text","TextColour","FontName","FontSize","TargetPath","LeftEdge","TopEdge"]
		#Parameters = [Text,TextColour,FontName,FontSize,TargetPath,LeftEdge,TopEdge]
		VIPath = toolpath + "\\viAccessors\\insertTextInVi.vi"
		
		#Get VI Reference
		VirtualInstrument = Application.GetVIReference(VIPath)
		
		#Open VI front panel
		VirtualInstrument.OpenFrontPanel(True,1)
	
		#back to GIMP operations
		layers = img.layers
		middleIndex = numLayers
		if desiredLayer == None:
			if numLayers == 1:
				print numLayers,"layer found:"
			else:
				print numLayers,"layers found:"
			for lay in layers:
				print "layer",pdb.gimp_item_get_name(lay)
		else:
			middleIndex = pdb.gimp_image_get_item_position(img,desiredLayer)
		#imgType = pdb.gimp_image_base_type(img)
		#textLayerType = 2*imgType
		for lay in layers:
			pos = pdb.gimp_image_get_item_position(img,lay)
			if(pos < middleIndex):
				#this is perhaps our text layer
				if(pdb.gimp_item_is_text_layer(lay)):
					print("idenfied text layer")
					Text = pdb.gimp_text_layer_get_text(lay)
					print 'text to insert:',Text
					FontName = pdb.gimp_text_layer_get_font(lay)
					print 'font:',FontName
					newTextColour = pdb.gimp_text_layer_get_color(lay)
					TextColour = pow(16,4)*newTextColour[0] + pow(16,2)*newTextColour[1] + newTextColour[2]
					FontSizes = pdb.gimp_text_layer_get_font_size(lay)
					FontSize = FontSizes[0]
					print 'font size:',FontSize
					LeftEdge,TopEdge = pdb.gimp_drawable_offsets(lay)
					# layerName = "imgLay"+pdb.gimp_item_get_name(lay)
					# textImgLayer = pdb.gimp_layer_new(img,imgWidth,imgHeight,textLayerType,layerName,100,0)
					# insertionPos = 1 + pdb.gimp_image_get_item_position(img,lay)
					# pdb.gimp_image_insert_layer(img,textImgLayer,None,insertionPos)
					# mergedLayer = pdb.gimp_image_merge_down(img,lay,1)
					# pdb.plug_in_autocrop_layer(img,mergedLayer)
					# offX,offY = pdb.gimp_drawable_offsets(mergedLayer)
					print "layer offset",LeftEdge,",",TopEdge
					#Calling VI
					Parameters = [Text,TextColour,FontName,FontSize,TargetPath,LeftEdge,TopEdge]
					VirtualInstrument.Call(ParameterNames,Parameters)
				else:
					continue
			else:
				continue
	
		#VirtualInstrument.CloseFrontPanel()
		
		#prepare parameters for recolourVi.vi
		
		#chartreuse green
		DesiredColour = 13369106
		recolourParamNames = ["DesiredColour","TargetPath"]
		recolourParams = [DesiredColour,TargetPath]
		VIPath = toolpath + "\\viAccessors\\recolourVi.vi"
		
		VirtualInstrument = Application.GetVIReference(VIPath)
		
		#Open VI front panel
		VirtualInstrument.OpenFrontPanel(True,1)
		VirtualInstrument.Call(recolourParamNames,recolourParams)
		#VirtualInstrument.CloseFrontPanel()
		
	except:
		VirtualInstrument = None
		Application = None
		
		#rethrow the exception to get the full trace on the console
		raise
	
	VirtualInstrument = None
	Application = None
		
	print "finished"
	
register(
    "VI_image_decomposition",
	"This script analyses the image layers to modify a LabVIEW virtual instrument",
	"",
	"Owen Miller",
	"Precision Mechatronics",
	"February 2013",
	"<Toolbox>/Xtns/Languages/Python-Fu/_VI Scripts/_Decomposition",
	"",
	[
	    (PF_STRING, "imgFilename","Input name",""),
		(PF_STRING, "toolpath","Decomposition Toolpath",""),
		(PF_STRING, "middleLayer","VI Layer's Name","VI"),
		(PF_STRING, "targetVI","Name of target VI","")
	],
	[],
	viGuiDecomposition)
	
main()