import os
import comtypes.client

class ppt2image:
	"""This class processes powerpoint files as images"""	
	
	# Initialize powerpoint located in path_to_ppt 
	def __init__(self, directory): 
		self.location = directory

	# Save a powerpoint file in path_to_ppt as a JPG in path_to_image
	def saveAsJPG(self, directory):
		[os.rename(l, l.replace(' ', '')) for l in os.listdir(directory)]
		for l in os.listdir(directory):
			path_to_ppt = os.path.abspath(l)
			path_to_image = path_to_ppt
			powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
			# Needed for script to work, don't know why... 
			powerpoint.Presentations.Open(path_to_ppt)
			powerpoint.Visible = True
			# Export to JPG	
			powerpoint.ActivePresentation.Export(path_to_image, "JPG")
			powerpoint.Presentations[1].Close()
			powerpoint.Quit()

	# Rename JPG image in path_to_image 	
	def renameJPG(self, path_to_image):
		for filename in os.listdir(path_to_image):	
			if filename == "Slide1.JPG":
				print "found"
				os.renames("Slide1.JPG", slidename)
				break
			else: 
				continue 
		#	else: 
		#		print "Image file not found"	

	def test(self):
		return self.location



"""Test Code""" 
prs = ppt2image("M:/Brad_Slide_Updates2/Practice")
prs.saveAsJPG("M:/Brad_Slide_Updates2/Practice")






#instance of ppt2image
#

# prs.renameJPG("M:/")





#print ppt2image("N:/Research/_Long Term Storage/Presentations/Public REIT Equity Market Capitalization", "M:/Compensation").test()


#.renameJPG("test1", "M:/")
#t.renameJPG("test", "M:/")
