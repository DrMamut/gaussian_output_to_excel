# gaussian_output_to_excel

Welcome to Gaussian Output to Excel program.
                                          written by DrMamut

 This python script is capable of finding, reading, and extracting important information from all the Gaussian output files (.out) found on the directory and all subdirectories where it is stored.

 It will grab the typical information needed from those files, like energies or the presence of imaginary frequencies from optimized jobs that were sent to Gaussian. Is capable of reading DFT, TD-DFT, transition states, solvent, and other types of data that could be found on those jobs.

 To use it, make sure that an up-to-date version of python with the Xlswriter plugin is installed, to run python
and to write the excel file. 

 I believe, the Anaconda version of python includes Xlswriter, but it could also be found here:

https://xlsxwriter.readthedocs.io/

And python for windows could be found here:
https://www.python.org/downloads/release/python-3100/

 Once they are installed properly, to run this program, just place gaussian_output_to_excel.py in a folder that contains .out files in it and its subfolder. Double click on it and wait. A new excel file will be created with all the information in it.

 Enjoy.
