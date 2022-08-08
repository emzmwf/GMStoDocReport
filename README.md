# GMStoDocReport
Python script to automate exporting a .docx report with images and metadata with Gatan Microscopy Suite

Requires:
Python running with Gatan Microscopy Suite

numpy

pathlib

docx

docxcompose

Template files:

Header.docx

meta.docx

Endtext.docx

meta file selects the metadata to report. The script currently accesses the following .dm tags - 

    "DataBar:Acquisition Date",
    "Session Info:Microscope",
    "DataBar:Signal Name",
    "Microscope Info:Formatted Indicated Mag",
    "Microscope Info:Stage Position:Stage X",
    "Microscope Info:Stage Position:Stage Y",
    "Microscope Info:Stage Position:Stage Z"


Any additional microscope info tags could be coded into the script and added to the meta.docx file to report. 
restoration of PIL/pillow support for python filtering planned for future version
