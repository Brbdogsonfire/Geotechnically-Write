# public
geotechnically_write

Geotechnically Write is a program to help create a base line geotechnical report minus engineering and geologic analysis. It is meant to pull information from a file created during data entry from software called Gint used for boring log creation, and from text input, and button selections from the user. 

The program uses several python libraries including Docx, Pygame, and openpyxl, Pandas. The Docx library loads an  existing microsoft word file and edits into it the data it processes through the other libraries. This choice to require a blank text file is due to  most geotechnical reports requiring a specific "stamp" from a geotechnical engineer that is unique to each. By having the document the stamp and any iconography of a company can easily be integrated into this program without the user having knowledge of programming. The openpyxl and pandas libraries are used to sort through and organize the data file "geo.xlsx" and organizes the data into a dataframe that is further broken down and engineering calculations are performed. Once the data has been prepared it is fed back into the original docx file to create a report from the lab data. The pygame library is used as the user interface.

The excel file from Gint needs to be named "Geo.xlsx", and a base report must be provided named "main_report.docx"

Any files produced from this should be reviewed by a professional and should not be considered as accurate or reliable. This program is meant to organize the data for an engineer to be able to review more quickly and to be more productive.
