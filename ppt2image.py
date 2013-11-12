<<<<<<< HEAD
import os
import comtypes.client
from PIL import Image

class ppt2Image:
	"""This class processes powerpoint files as images"""	
	
	# Initialize powerpoint located in path_to_ppt 
	def __init__(self, directory): 
		self.location = directory

	# Get a dict of all original ppt names matched to their "unspaced" versions
	def namingDict(self, directory):
		names = {} 
		for l in os.listdir(str(directory)): 
			if ('pptx' in l) and (' ' in l):
				unspaced = l.replace(' ', '')
				pdfversion = unspaced.replace('pptx', 'pdf')
				names[pdfversion] = l.replace('pptx', 'pdf')
		return names


	# Save a powerpoint file in path_to_ppt as a JPG named "Slide1" in the same directory
	def saveAsJPG(self, pptfile, directory):
		# Make sure no filename has spaces in it
		[os.rename(l, l.replace(' ', '')) for l in os.listdir(str(directory)) if ' ' in l]
		# Get full path to the ppt file
		path_to_ppt = str(os.path.dirname(directory)) + "/" + str(pptfile)
		print path_to_ppt
		# Initialize ppt object pointing to the ppt file
		powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
		powerpoint.Presentations.Open(path_to_ppt)
		powerpoint.Visible = True
		# Export the ppt to JPG
		powerpoint.ActivePresentation.Export(os.path.dirname(directory), "JPG")
		# Close ppt
		powerpoint.Quit()

	# Rename JPG image "Slide1.JPG" to the (unspaced) ppt file it refers to
	def renameJPG(self, pptfile, directory):
		for l in os.listdir(directory):
			if ("Slide1.JPG" in l):
				print "Slide1.JPG found!"
				fullJPGpath = directory+"Slide1.JPG"
				newJPGpath = directory+pptfile.replace('pptx', 'JPG')
				print fullJPGpath
				print newJPGpath
				os.rename(fullJPGpath, newJPGpath)

	# Crop the blue edges off each JPG image, resize, and save as PDF to upload on the web
	def cropAndResize(self, jpgfile, directory):
		imgpath = directory+jpgfile
		outfile = imgpath.replace('JPG', 'PDF')
		print outfile
		im = Image.open(imgpath)
		cropbox = (56, 0, 960, 664)
		resizedimg = im.crop(cropbox).resize((448, 336))
		resizedimg.load()
		resizedimg.save(outfile)
		return

	# "space" the name of each edited pdf file to match previously existing versions
	def matchPdfName(self, namedict, pdffile, directory):
		for l in os.listdir(directory):
			if ("pdf" in l):
				newname = namedict[l]
				print newname				

	def test(self):
		return self.location


"""Test Code""" 
#This is some change I've made.
# Get instance of ppt2
prs = ppt2Image("M:/Brad_Slide_Updates2/Practice/")
#prs.saveAsJPG("EquityREITReturnsduringaPeriodofRisingInterestRates.pptx", "M:/Brad_Slide_Updates2/Practice/")
#prs.renameJPG("EquityREITReturnsduringaPeriodofRisingInterestRates.pptx", "M:/Brad_Slide_Updates2/Practice/")
#prs.cropAndResize("EquityREITReturnsduringaPeriodofRisingInterestRates.JPG", "M:/Brad_Slide_Updates2/Practice/")
namedict = prs.namingDict("M:/Brad_Slide_Updates2/Practice/")
prs.matchPdfName(namedict, "EquityREITReturnsduringaPeriodofRisingInterestRates.PDF","M:/Brad_Slide_Updates2/Practice/")
#for k in prs.namingDict("M:/Brad_Slide_Updates2/Practice/"):
#	print k
#	print prs.namingDict("M:/Brad_Slide_Updates2/Practice/")[k]


"Here's a change"