"""This program practices editing Powerpoint with windows COM client"""

# Import windows COM client - an interface for windows applications
import win32com.client, sys

# Open Powerpoint
app = win32com.client.Dispatch("PowerPoint.Application")

# Create new presentation 
prs = app.Presentations.Add()

# Add a blank slide. 12 is the code for blank slide. 
sld = prs.Slides.Add(1, 12)

# Open PowerPoint
Application = win32com.client.Dispatch("PowerPoint.Application")

# Add a presentation
Presentation = Application.Presentations.Add()

# Add a slide with a blank layout (12 stands for blank layout)
Base = Presentation.Slides.Add(1, 12)

# Add an oval. Shape 9 is an oval.
oval = Base.Shapes.AddShape(9, 100, 100, 100, 100)