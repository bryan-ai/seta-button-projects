import pandas as pd
import os
from pathlib import Path
import subprocess
import logging
from time import sleep
import logging

# Default path for many functions.
WORKING_DIR = Path(os.getcwd())

logging.basicConfig(filename='logging.log', encoding='utf-8', level=logging.DEBUG)

def setup():
    setup_string = f"./setup.sh"
    logging.info(f"TERMINAL COMMAND: {setup_string}")
    subprocess.call(setup_string, shell=True)

def get_files(team="",files_path=WORKING_DIR):
    #TODO find a better way to handle non submissions. This is an awful hack
    try:
        mylist = os.listdir(files_path)
    except Exception as e:
        logging.error(e)
        print(f"No submission on record. Did {team} submit a zip file?")
        print(f"Simulating a blank submission for {team}")
        subprocess.call(["mkdir",files_path])
        mylist=[]

    # rm_ipynb_string = "rm -rf pages/*.ipynb"
    # subprocess.call(rm_ipynb_string, shell=True)
    # # mylist = os.listdir("pages")
    # print(mylist)
    accepted_file_dictionary = {}
    #TODO Figure out a config file
    accepted_filetype_list =["png","pdf","ipynb","pptx","docx","jpeg"]
    blank_files = False
    for i in range(len(mylist)):
        file = mylist[i]
        try:
            filetype = file.split('.')[1].lower()
            logging.info(f"Filename is {mylist[i]} and filetype is ${filetype}$")
            if filetype in accepted_filetype_list:
                if filetype not in accepted_file_dictionary:
                    accepted_file_dictionary[filetype]=[]
                    accepted_file_dictionary[filetype].append(file)
                else:
                    accepted_file_dictionary[filetype].append(file)
            else:
                
                logging.info(f"removing {file}")
                logging.info(f"TERMINAL COMMAND: rm -f {os.path.join(files_path,file)}")
                subprocess.call(["rm","-f",os.path.join(files_path,team,file)])
        except Exception as e:
            print(e, "WARNING: Could not get filetype - suspect no file extension")
            blank_files = True
            continue
    
    '''Accounting for IPYNBs that were not labelled as such, that you manually rename, and have not been converted to PDF. e.g. AE5'''
    if blank_files:
        print(f"WARNING: Team {team} submitted at least 1 file of unknown type. I need your help to identify any important files and add the filetype, or delete the irrelavent files.")
        open_finder=input(f"Would you like to open a finder window for team {team}'s submission? [Y]es or [N]o: \n")
        if open_finder.lower() == "yes" or open_finder.lower() == 'y':
            folder_to_show = os.path.join(files_path)
            print(f"opening {folder_to_show}")
            logging.info(f"TERMINAL COMMAND: open -R {folder_to_show}")
            subprocess.call(["open", "-R", folder_to_show])
            continue_string = input("When you have dealt with the files, press Enter to continue")
    return accepted_file_dictionary

def lower_file_extension_case(team="",filename="", input_path=WORKING_DIR, output_path=WORKING_DIR):
    input_file_path = os.path.join(input_path,filename)
    output_file_path = os.path.join(output_path,filename.lower())
    print(f"Renaming {filename} to {filename.lower()}")
    logging.info(f"TERMINAL COMMAND: mv {input_file_path} {output_file_path}")
    subprocess.run(["mv", input_file_path,output_file_path])

def files_converter_helper(team="",filetype="",file_list=[], input_path=WORKING_DIR, output_path=WORKING_DIR):
    #TODO you need to figure out how to direct users to issues
    
    if len(file_list) < 1:
        print(f"No files found")
    elif len(file_list) > 0:
        if filetype=="png": 
            png_to_pdf(team=team,input_path=input_path,output_path=output_path)
        if filetype=="pptx":
            pptx_to_pdf_helper(team=team,file_list=file_list,input_path=input_path, output_path=output_path)
        if filetype=="ipynb":
            ipynb_to_pdf(team=team,file_list=file_list,input_path=input_path,output_path=output_path)
    


def jpeg_to_png(team="",input_path=WORKING_DIR, output_path=WORKING_DIR):
    logging.info(f"team is {team} and input_path is {input_path} and output_path is {output_path}")
    input_file_path = os.path.join(input_path, "*.jpeg")
    output_file_path = os.path.join(output_path, "4images.png")
    logging.info(input_file_path)
    logging.info(output_file_path)
    logging.info(f"TERMINAL COMMAND: magick {input_file_path} {output_file_path}")
    subprocess.call(["magick", input_file_path, output_file_path])


def png_to_pdf(team="",input_path=WORKING_DIR, output_path=WORKING_DIR):

    input_file_path = os.path.join(input_path, "*.png")
    output_file_path = os.path.join(output_path, "4images.pdf")
    terminal_string = ' '.join(["magick", input_file_path, output_file_path])
    logging.info(f"TERMINAL COMMAND: {terminal_string}")
    subprocess.call(["magick", input_file_path, output_file_path])

def ipynb_to_pdf(team="", file_list=[], input_path=WORKING_DIR, output_path=WORKING_DIR):
    wkhtmltopdf_flag_string = "-q -s A4 --print-media-type --disable-smart-shrinking --margin-top 15mm --margin-bottom 15mm --margin-left 15mm --margin-right 15mm --no-background"
    for file in file_list:
        print(f"Converting {team}'s {file} to html")
        ipynb_home_dir=os.path.join(input_path,file)
        logging.info(f"html_ouptu_dir is {ipynb_home_dir}")
        logging.info(f"output_path is {output_path}")
        terminal_string = " ".join(["jupyter-nbconvert", "--to","html","--template","portfolio",ipynb_home_dir])
        logging.info(f"TERMINAL COMMAND: {terminal_string}")
        subprocess.call(["jupyter-nbconvert", "--to","html","--template","portfolio",ipynb_home_dir])
        logging.info(file_list)

    
        print(f"Converting {file.split('.')[0]}.html to pdf")
        html_input_path = os.path.join(input_path,f"{file.split('.')[0]}.html")
        pdf_output_path = os.path.join(output_path,f"{file.split('.')[0]}.pdf")
        logging.info(f"html_input_path is {html_input_path}")
        logging.info(f"pdf_output_path is {pdf_output_path}")
        string = " ".join(["wkhtmltopdf", wkhtmltopdf_flag_string,html_input_path,pdf_output_path])
        logging.info(f"TERMINAL COMMAND: {string}")
        subprocess.run(["wkhtmltopdf","-q","-s","A4","--print-media-type", "--disable-smart-shrinking","--margin-top","15mm","--margin-bottom","15mm","--margin-left","15mm","--margin-right","15mm","--no-background", html_input_path,pdf_output_path])




def pptx_to_pdf_helper(team="",file_list=[],input_path=WORKING_DIR, output_path=WORKING_DIR):
    for filename in file_list:
        
        file_renamed = False
        file_to_show = os.path.join(input_path,filename)
        file_rename = os.path.join(input_path,"2"+filename)
        print(f"Renaming {filename} to 2{filename}")
        try:
            logging.info(f"TERMINAL COMMAND: mv {file_to_show} {file_rename}")
            subprocess.call(["mv",file_to_show,file_rename])
            file_to_show = file_rename
        except Exception as e:
            logging.info(e) 
            print(f"ERROR: Could not rename {file_rename} perhaps the files is already renamed?")
            file_renamed = True
        print(f"Team {team} submitted a .pptx file {filename}. I need your help to convert it to pdf. Please use your tool of choice to convert the pptx to PDF, and delete the pptx")
        open_finder=input("Would you like to open a finder window for the team's submission? [Y]es or [N]o: ")
        if open_finder.lower() == "yes" or open_finder.lower() == 'y':
            # file_to_show = os.path.join(input_path,team,filename)
            logging.info(f"TERMINAL COMMAND: open -R {file_to_show}")
            subprocess.call(["open", "-R", file_to_show])
            pass
        else: 
            print(f"continuing")

def pdf_mover(team="",filetype="",file_list=[], input_path=WORKING_DIR, output_path=WORKING_DIR):
    #TODO something is wrong here... I keep getting "usage: mv [-f | -i | -n] [-hv] source target mv [-f | -i | -n] [-v] source ... directory" in stdout
    for file in file_list:
        original_file_path = os.path.join(input_path,file)
        new_file_path = os.path.join(output_path,file)
        print(f"Moving {file} to {new_file_path}")
        logging.info(f"TERMINAL COMMAND: mv {original_file_path} {new_file_path}")
        subprocess.run(["mv", original_file_path,new_file_path])

def multi_file_helper(team="",filename="", filetype="pdf",input_path=WORKING_DIR, output_path=WORKING_DIR):
    file_to_show = os.path.join(input_path,filename)
    print(f"team {team}'s folder has one or more unlabelled {filetype} in their folder. I need your help to ensure their names help us merge the files in order.")

    print("1.Cover page")
    print("2.Presentation")
    print("3.Notebook")
    print("4.Images")
    print("5.Student marks")
    print("6.Rubric")
    print("7.Comments")
    print("8.Skip")
    print("9.Delete File")
    
    mv_or_open=input(f"Please choose {filename}'s position by selecting the number or delete the file by slecting 9 \nAlternatively, would you like to open a finder window for team {team}'s folder to rename them yourself? [1-9], [Y]es or [N]o: \n")
    if mv_or_open.lower() == "yes" or mv_or_open.lower() == 'y':
        file_to_show = os.path.join(input_path,filename)
        try:
            logging.info(f"TERMINAL COMMAND: open -R {file_to_show}")
            subprocess.call(["open", "-R", file_to_show])
        except Exception as e:
            logging.info(e)
            print("WARNING: Could not find file, opening folder instead")
            folder_to_show =os.path.join(input_path,team)
            logging.info(f"TERMINAL COMMAND: open -R {file_to_show}")
            subprocess.call(["open", "-R", file_to_show])
        pass
    elif mv_or_open.isdigit():
        if int(mv_or_open) < 8:
            file_rename = os.path.join(input_path,mv_or_open+filename)
            print(f"renaming {filename} to {mv_or_open}{filename}")
            logging.info(f"TERMINAL COMMAND: mv {file_to_show} {file_rename}")
            subprocess.call(["mv", file_to_show,file_rename])
            print(f"continuing")
        elif int(mv_or_open) == 7:
            pass
        elif int(mv_or_open) == 9:
            print(f"Deleting {file_to_show}")
            logging.info(f"TERMINAL COMMAND: rm -f {file_to_show}")
            subprocess.call(["rm","-f", file_to_show])

if __name__ == "__main__":
    team = "AM6"
    files_dictionary = get_files(team)
    print(files_dictionary)
    filetypes_list = list(files_dictionary)
    for filetype in filetypes_list:
        if(filetype != "png") and (filetype!="jpeg") and (len(files_dictionary[filetype])>1):
            for filename in files_dictionary[filetype]:
                multi_file_helper(team=team, filename=filename, filetype=filetype)
    files_dictionary = get_files(team)


