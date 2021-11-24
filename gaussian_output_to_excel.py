import re
import xlsxwriter
import os
#import fnmatch
import os.path

#getting the directory information to get ready to build the excel file
filepath = os.getcwd()
#foldername = os.path.dirname(filepath)
excelfilename=filepath.replace(":","").replace("/","_").replace("\\",'_')

#writing the excel file and formatting
workbook = xlsxwriter.Workbook("Gaussian_ouput_summary_"+excelfilename+'.xlsx')
worksheet = workbook.add_worksheet()
cell_format = workbook.add_format()
bold = workbook.add_format({'bold': True})
center = workbook.add_format({'align': 'center'})

worksheet.set_column(0, 0, 148)
worksheet.set_column(1, 5, 10)
worksheet.set_column(5, 5, 20)
worksheet.set_column(6, 10, 7)
worksheet.set_column(10, 12, 20)
worksheet.set_column(12, 12, 16)


worksheet.write('A1', 'Cruntch3')
worksheet.write('D1', 'Abstract')
worksheet.write('E1', 'Comments')
worksheet.write('G1', 'Cartesian Basis')
worksheet.write('K1', 'Enthalpies')
worksheet.write('L1', 'Free Energies')
worksheet.write_rich_string('A2',  bold, 'Path and filename',  " ", center )
worksheet.write_rich_string('B2',  bold, 'Functional',  " ",  center  )
worksheet.write_rich_string('C2',  bold, 'Basis Set',  " ", center )
worksheet.write_rich_string('D2',  bold, 'Charge',  " ",  center )
worksheet.write_rich_string('E2',  bold, 'Multiplicity',  ' ', center )
worksheet.write_rich_string('F2',  bold, 'Stoichiometry',  " ", center )
worksheet.write_rich_string('G2',  bold, "Bf's",  " ", center )
worksheet.write_rich_string('H2',  bold, 'Alpha',  " ", center )
worksheet.write_rich_string('I2',  bold, 'Beta',  " ", center )
worksheet.write_rich_string('J2',  bold, 'imag v',  " ", center )
worksheet.write_rich_string('K2',  bold, 'H (Hartrees/particles)',  " ", center )
worksheet.write_rich_string('L2',  bold, 'G (Hartrees/particles)',  " ", center )
worksheet.write_rich_string('M2',  bold, 'Solvent',  " ", center )
worksheet.write('A3', filepath, bold)


row=2

#finding the ouput files

for dirpath, dirnames, file in os.walk("."):
    
    for file in [f for f in file if f.endswith(".out")]:
            file=os.path.join(filepath+dirpath, file).replace(".\\","\\")
            worksheet.write(row+1, 0, file)
            with open(file, "r") as file:
                for line in file: 
                    line=line.rstrip() 
                    if re.search('^ #.+/', line):
                        functional=re.search(r'\S+/',line).group()
                        basis_set=re.search(r'/\S+',line).group()
                        worksheet.write(row+1, 1, str(functional).replace("/",""))
                        worksheet.write(row+1, 2, str(basis_set).replace("/",""))      


                    if re.findall('Charge =', line):
                        charge=re.search(r'([0-9]+)',line).group()
                        multiplicity=re.search(r'\S+$',line).group()
                        worksheet.write(row+1, 3, float(charge))
                        worksheet.write(row+1, 4, float(multiplicity)) 
                        
                    if re.findall(" Stoichiometry", line):
                        stoi=re.search(r'\S+$',line).group()
                        worksheet.write(row+1, 5, str(stoi))
                                
                    if re.findall("basis functions,", line):
                        basisf= re.search(r'(?<!basis functions)\w+',line).group()
                        worksheet.write(row+1, 6, float(basisf))
                                       
                    if re.findall("alpha electrons", line):
                        alpha=re.search(r'(?<!alpha electrons)\w+', line).group()
                        beta=re.findall('^.*alpha electrons.* ([0-9]+) ', line)
                        worksheet.write(row+1, 7, float(alpha))
                        worksheet.write(row+1, 8, int(str(beta).replace("'", "").replace("[","").replace("]","")))
                    
                    if re.search('^ Solvent.+:', line):
                        solvent=str(re.search(r':\s\S+', line).group())
                        worksheet.write(row+1, 12, solvent.replace(":","").replace(",",""))    

                    if re.findall('thermal Enthalpies=', line):
                        enthalpy=re.search(r' \S+$', line).group()
                        worksheet.write(row+1, 10, float(enthalpy))
                        
                    if re.findall(' Sum of electronic and thermal Free Energies=', line):
                        G=re.search(r'\S+$', line).group()
                        worksheet.write(row+1, 11, float(G))
                    
                    try:
                        if re.findall("negative Signs", line):
                            vimag=float(re.search(r'(?<!imaginary frequencies)\w+', line).group())
                            worksheet.write(row+1, 9, vimag)

                        elif re.findall("mag=0" or "Imag=0" or "NImag=0", line):
                            worksheet.write(row+1, 9, float(0))

                        elif re.findall('g_write', line):
                            worksheet.write(row+1, 9, "incomplete")

                        elif re.findall("Error termination", line):
                            worksheet.write(row+1, 13, str(line))
                            worksheet.write(row+1, 9, "error") 
                    except:
                        worksheet.write(row+1, 9, float(0))
                    
            row=row+1


workbook.close()

#Created by DrMamut