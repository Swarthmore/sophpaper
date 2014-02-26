# Take a CSV file with the output of a Sophomore paper survey in Qualtrics and output
# individual PDF's for each student.

import string
import re
import csv

from pyPdf import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas  
from reportlab.lib.units import inch 
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph, SimpleDocTemplate,Spacer,PageBreak

from reportlab.pdfbase import _fontdata_enc_winansi
from reportlab.pdfbase import _fontdata_enc_macroman
from reportlab.pdfbase import _fontdata_enc_standard
from reportlab.pdfbase import _fontdata_enc_symbol
from reportlab.pdfbase import _fontdata_enc_zapfdingbats
from reportlab.pdfbase import _fontdata_enc_pdfdoc
from reportlab.pdfbase import _fontdata_enc_macexpert
from reportlab.pdfbase import _fontdata_widths_courier
from reportlab.pdfbase import _fontdata_widths_courierbold
from reportlab.pdfbase import _fontdata_widths_courieroblique
from reportlab.pdfbase import _fontdata_widths_courierboldoblique
from reportlab.pdfbase import _fontdata_widths_helvetica
from reportlab.pdfbase import _fontdata_widths_helveticabold
from reportlab.pdfbase import _fontdata_widths_helveticaoblique
from reportlab.pdfbase import _fontdata_widths_helveticaboldoblique
from reportlab.pdfbase import _fontdata_widths_timesroman
from reportlab.pdfbase import _fontdata_widths_timesbold
from reportlab.pdfbase import _fontdata_widths_timesitalic
from reportlab.pdfbase import _fontdata_widths_timesbolditalic
from reportlab.pdfbase import _fontdata_widths_symbol
from reportlab.pdfbase import _fontdata_widths_zapfdingbats
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from bs4 import BeautifulSoup

REMOVE_ATTRIBUTES = ['font']


csv_file_name = "Soph_Plan_Submission.csv"


pdfmetrics.registerFont(TTFont('trebuchet ms', '/Library/Fonts/Trebuchet MS.ttf'))
pdfmetrics.registerFont(TTFont('trebuchet ms-Bold', '/Library/Fonts/Trebuchet MS Bold.ttf'))
pdfmetrics.registerFont(TTFont('trebuchet ms-Italic', '/Library/Fonts/Trebuchet MS Italic.ttf'))
pdfmetrics.registerFont(TTFont('trebuchet ms-BoldItalic', '/Library/Fonts/Trebuchet MS Bold Italic.ttf'))
           
           
# Return only the objects with keys that start with s           
# From http://stackoverflow.com/a/4558999
def slicedict(d, s):
    return {k:v for k,v in d.iteritems() if k.startswith(s)}




def generate_pdf(response):

	# Make sure we have all the proper information
	#if not (response["Name"] and response["EmailAddress"] and response["Banner ID"]):
	#	print "Not enough information to process response #%s" % response["ResponseID"]
	#	return

	print response["Name"]

	# Set up the PDF file
	pdf_file_name = "Sophomore_Paper_" + response["Banner ID"] + ".pdf"
	# If the banner ID isn't present, save with the response ID
	if not response["Banner ID"]: 
		# Use the username from the email address
		pdf_file_name = "Sophomore_Paper_" + response["EmailAddress"].split("@")[0] + ".pdf"
	pdf = SimpleDocTemplate(pdf_file_name, pagesize = letter, rightMargin=72,leftMargin=72, topMargin=36,bottomMargin=18)

		
	# Create the content 
	story = []
	style = getSampleStyleSheet()
	
	# Name
	story.append(Paragraph("Sophomore Paper for %s" % response["Name"], style["Heading2"]))
	
	# Sub-heading (class year, email, date completed)
	if not response["Class Year"]:
		response["Class Year"] = "unknown"
	
	
	sub_heading = "Class year: %s <br/>%s<br/>Completed %s" % (response["Class Year"], response["EmailAddress"], response["EndDate"])
	story.append(Paragraph(sub_heading, style["Heading4"]))
	
	
	# Add a space
	story.append(Paragraph("<br/>", style["Normal"]))
	
	##################################
	# Majors
	##################################
	# Find all the questions that match the major question wording
	major_dict = slicedict(response, "Major (one required); you may select up to 2 majors")
	
	# Loop through all the possible majors, looking for a selection
	majors = []
	for k,v in major_dict.iteritems():
		if v:
			# Neat way to strip off a prefix (in this case, the question text).  See http://stackoverflow.com/a/5293388
			major = k.split("Major (one required); you may select up to 2 majors-").pop()
			
			if major == "Special Major":
				# If the special major value is filled out, include it otherwise print unknown
				if response["Special Major: you must fill out the Special Major form / available from the Registrar's Office and submit it to the / appropriate departments. / Title of your Special Major:"]:
					special_major_title = response["Special Major: you must fill out the Special Major form / available from the Registrar's Office and submit it to the / appropriate departments. / Title of your Special Major:"]
				else:
					special_major_title = "Unknown"

				major = "Special Major: %s" % special_major_title
				
			majors.append(major)
				
	
	# Print out majors
	if len(majors) < 1:
		majors = ["None"]
	story.append(Paragraph("Major", style["Heading2"]))
	story.append(Paragraph("<br/>".join(majors), style["Normal"]))		
			
	# Add a space
	story.append(Paragraph("<br/>", style["Normal"]))

	##################################
	# Minors
	##################################
	# Find all the questions that match the minor question wording
	minor_dict = slicedict(response, "Minor(s); no more than 2 minors")
	
	# Loop through all the possible majors, looking for a selection
	minors = []
	for k,v in minor_dict.iteritems():
		if v:
			# Neat way to strip off a prefix (in this case, the question text).  See http://stackoverflow.com/a/5293388
			minor = k.split("Minor(s); no more than 2 minors-").pop()
			minors.append(minor)
				
	
	# Print out minors
	if len(minors) < 1:
		minors = ["None"]	
	story.append(Paragraph("Minor", style["Heading2"]))
	story.append(Paragraph("<br/>".join(minors), style["Normal"]))	
	
	# Add a space
	story.append(Paragraph("<br/>", style["Normal"]))	


	##################################
	# Other notes
	##################################
	
	# Advisor
	story.append(Paragraph("Sophomore Plan Advisor(s)", style["Heading2"]))
	advisor = response["Sophomore Plan Advisor(s): please list the name of your advisor(s) / for both the major or special major and any minor"]
	story.append(Paragraph(advisor, style["Normal"]))	
	
	story.append(Paragraph("Honors?", style["Heading2"]))
	honors = "No"		# Assume no
	if response["Honors?"] == "1":
		honors = "Yes"
	story.append(Paragraph(honors, style["Normal"]))	
	
	# Add a space
	story.append(Paragraph("<br/>", style["Normal"]))	
	
	story.append(Paragraph("Teacher Certification?", style["Heading2"]))
	teaching_cert = "No"		# Assume no
	if response["Teacher Certification"] == "1":
		teaching_cert = "Yes"
	story.append(Paragraph(teaching_cert, style["Normal"]))		

	# Add a space
	story.append(Paragraph("<br/>", style["Normal"]))		
	
	##################################
	# Plan of Study Narrative
	##################################
	story.append(Paragraph("Plan of Study Narrative", style["Heading2"]))
	story.append(Paragraph(response["Plan of Study Narrative (200 to 500 words)"].replace("/","<br/>"), style["Normal"]))		
	
	# Add a space
	story.append(Paragraph("<br/>", style["Normal"]))	
	
	##################################
	# Course plan
	##################################
	story.append(PageBreak())
	story.append(Paragraph("List of courses for the last two years", style["Heading2"]))
	
	# Output course listings.  Multiline box replaces newlines with "/ ", so switch it back.
	story.append(Paragraph("Junior Fall", style["Heading3"]))
	story.append(Paragraph(response["Junior Fall"].replace("/","<br/>"), style["Normal"]))		
	
	story.append(Paragraph("Junior Spring", style["Heading3"]))
	story.append(Paragraph(response["Junior Spring"].replace("/","<br/>"), style["Normal"]))		
	
	story.append(Paragraph("Senior Fall", style["Heading3"]))
	story.append(Paragraph(response["Senior Fall"].replace("/","<br/>"), style["Normal"]))				
	
	story.append(Paragraph("Senior Spring", style["Heading3"]))
	story.append(Paragraph(response["Senior Spring"].replace("/","<br/>"), style["Normal"]))			
	
	pdf.build(story)



	# Return an index for the student with 'Name', 'BannerID', 'Major', 'Minor'
	index = {}
	index["Name"] = response["Name"]
	index["BannerID"] = response["Banner ID"]
	index["Major"] = "|".join(majors)
	index["Minor"] = "|".join(minors)
	
	return index








################# MAIN PROGRAM #################


# Open the PDF file, skip the first header row, and use the next row as the headers
f = open(csv_file_name,'rb')
f.seek(0)
next(f)
input_file = csv.DictReader(f, delimiter=',')

# Setup the index file
index_file = open('sophomore_paper_index.csv','wb')
fieldnames = ['Name', 'BannerID', 'Major', 'Minor']
csvwriter = csv.DictWriter(index_file, delimiter=',', fieldnames=fieldnames)
csvwriter.writerow(dict((fn,fn) for fn in fieldnames))

for row in input_file:
	index_data = generate_pdf(row)
	csvwriter.writerow(index_data)
exit(0)
