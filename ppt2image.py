import os
import comtypes.client
from PIL import Image
import re
import shutil

""" 
Author: Shruthi Venkatesh
Date: 11/13/13

This program converts Brad's updated slides (located in N:/Research/_Long Term Storage/Presentations)
to JPEGs; then crops, resizes, and converts the JPEG to PDF format to upload them
to the reit.com website via FileZilla.

The date, resolution, and updated slide list MUST be edited below. 
""" 
# IMPORTANT: Edit the below variable assignments
date = "110113"
directory = "M:/Brad_Slide_Updates2/" + date + "/"
res_val = 3000 
updated = ["AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", 
"AL", "AM", "AN", "AO", "AP", "AQ", "AR", "BA", "BB", "BC", "BD", "BE",
"CA", "DB", "DC", "CB", "CC", "DC", "CD", "CE", "CF", "DJ", "CG", "CH", 
"CI", "CJ", "CK", "DO", "DP"] # Enter a list of two-letter strings of slide names here

def copyUpdatedSlides(directory, updated): 
	# Generate new directory to place COPIES of Brad's most recently updated files
	if not os.path.exists(directory):
		os.mkdir(directory)
	else: 
		print "New designated directory already exists!"
	# Make copy of recently updated files and place in above directory
	regexes = {}
	for letters in updated: 
		regexes[re.compile("^Slide " + letters + ".*")] = letters
	for n in os.listdir("N:/Research/_Long Term Storage/Presentations/"):
		for r in regexes.keys():
			if (r.match(n)):
				org_full_path = "N:/Research/_Long Term Storage/Presentations/" + n
				new_full_path = directory + n
				shutil.copy2(org_full_path, new_full_path)
				print "Copying " + n	

class ppt2Image:
	"""This class processes powerpoint files as images"""	
	
	# Initialize powerpoint located in path_to_ppt 
	def __init__(self, directory): 
		self.location = directory

	# Save a powerpoint file in path_to_ppt as a JPG named "Slide1" in the same directory
	def saveAsJPG(self, pptfile, directory):
		# Make sure no filename has spaces in it
		[os.rename(l, l.replace(' ', '')) for l in directory if ' ' in l]
		# Get full path to the ppt file
		path_to_ppt = str(directory) + str(pptfile)
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
		gen_name = re.compile('^Slide[0-9]+.JPG$')
		for l in os.listdir(directory):
			if gen_name.match(str(l)):
				fullJPGpath = directory+l
				newJPGpath = directory+pptfile.replace('pptx', 'JPG')
				print newJPGpath
				if not os.path.exists(newJPGpath):
					os.rename(fullJPGpath, newJPGpath)


	# Crop the blue edges off each JPG image, resize, and save as a JPG again
	def cropAndResize(self, jpgfile, directory, res_val):
		imgpath = directory+jpgfile
		outfile = imgpath.replace("jpg", "pdf")
		im = Image.open(imgpath)
		# Check if image is in RGB mode: 		
		if im.mode not in ('L', 'RGB', 'RGBA'):
		 	im = im.convert('RGB')
		# Designate dimensions of cropped image (remove blue edges)
		cropbox = (56, 0, 960, 664)
		# Crop and resize image
		im = im.crop(cropbox)
		im.thumbnail((575, 420), Image.ANTIALIAS)
		# Save edited image as pdf
		im.save(outfile, "PDF", resolution = res_val)

	def test(self):
		return self.location

"""Implement the above functions for a given date and updated list""" 
# Make copies of all updated files in a new directory
# copyUpdatedSlides(directory, updated)
# Convert each slide in directory to cropped and resized PDF! 
for pptfile in os.listdir(directory):
	print pptfile 
	jpgfile = (pptfile.replace("pptx", "jpg")).strip()
	prs = ppt2Image(directory)
	prs.saveAsJPG(pptfile, directory)
	prs.renameJPG(pptfile, directory)
	prs.cropAndResize(jpgfile, directory, res_val)


"""
Test Code
res_val = 3000
directory = "M:/Brad_Slide_Updates2/Practice/"
pptfile = "Slide QM (Public and Private Returns by Leverage, Common Available Period).pptx"
jpgfile = (pptfile.replace("pptx", "jpg")).strip()

# Get instance of ppt2image file 
prs = ppt2Image(directory)

# Save the ppt file as jpg
prs.saveAsJPG(pptfile, directory)

# Rename the JPG from "Slide0" to its actual name
prs.renameJPG(pptfile, directory)

# Edit the JPG
prs.cropAndResize(jpgfile, directory, res_val)

# Rename the edited JPG to match original ppt name 
#namedict = prs.namingDict(directory)
#prs.matchPdfName(namedict, directory, jpgfile)

#copyUpdatedSlides("M:/Brad_Slide_Updates2/Practice/", ["AK"]) 
"""