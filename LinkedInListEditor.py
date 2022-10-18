"""Reads an excel file that contains education and work experience 
cut-and-pasted from LinkedIn Recruiter results lists, 
and outputs the education and work experience entries in a more readable format. 
Note that the education output can have some arrangement problems that need to be fixed 
in cases when the original input string has commas used in the degree name.
"""

import pandas as pd
import openpyxl
import sys
import os

print('This program accepts an excel file (.xlxs) ' +
      'with work experience and education entries ' +
      'cut-and-pasted from LinkedIn Recruiter results pages ' +
      'and puts the entries into an easier to read format.\n')

filepath = input('Enter the path to your file, ' +
                 'including the filename with the filetype extension: ')

# Check that extension is xlsx
try:
    extension = os.path.split(filepath)[1].split('.')[1]
except IndexError:
    print('\nError:\nPlease ensure that the file extension ' +
          'is included in the filename and try again.')
    input('Press enter to exit.')
    sys.exit(1)

if extension != 'xlsx':
    print('\nError:\nThe file type must be xlsx. ' +
          'Please ensure the file is in the proper format and try again.')
    input('Press enter to exit.')
    sys.exit(1)

# Read the file and handle file / directory exception.
try:
    df = pd.read_excel(filepath)

except FileNotFoundError:
    print('\nFolder or Filename Error:\nThe folder path and/or filename entered does not exist. ' +
          'Format example: C:\\Users\\username\\foldername\\filename.xlsx' +
          'Please check and try again.')
    input('Press enter to exit.')
    sys.exit(1)

work_exp_column_name = input('Enter the column name that contains ' +
                             'the work experience entries: ')

education_column_name = input('Enter the column name that contains ' +
                              'the education entries: ')

# Replace carriage returns + line feeds with just line feeds.
# Also check the column entries while doing this.
try:
    df[work_exp_column_name] = df[work_exp_column_name].str.replace('\r\n', '\n')

except KeyError:
    print('\nKey Error:\nThe column name provided for the work experience entries does not exist. ' +
          'Please check the name and try again.')
    input('Press enter to exit.')
    sys.exit(1)

try:
    df[education_column_name] = df[education_column_name].str.replace('\r\n', '\n')

except KeyError:
    print('\nKey Error:\nThe column name provided for the education entries does not exist. ' +
          'Please check the name and try again.')
    input('Press enter to exit.')
    sys.exit(1)

# Replace carriage returns with just line feeds
df[work_exp_column_name] = df[work_exp_column_name].str.replace('\r', '\n')
df[education_column_name] = df[education_column_name].str.replace('\r', '\n')

# Replace NaN with empty strings to prevent errors in using string methods in code below
df[work_exp_column_name] = df[work_exp_column_name].fillna('')
df[education_column_name] = df[education_column_name].fillna('')


# Experience formatter function
def format_exp (input_exp_string):
    """Takes a work experience input string, as cut-and-pasted from LinkedIn Recruiter results page, 
    and puts it into a more reader friendly format.
    """
    
    # Split the work experience string into a list at each new line
    exp_list = input_exp_string.split('\n')
    
    # For each experience in the experience list, 
    # pull out the title, company, and period and put them in a list, 
    # then append each list to a full experience list, 
    # creating a nested list.
    full_exp_split_list = []
    
    try:
        for exp in exp_list:
            # Handle cases of "Profile experience" being copied over in the string.
            if exp == "Profile experience":
                print('\nWarning:\nThe text string \"Profile experience\" was found in one of the work experience entries. ' +
                      'This string is sometimes inadvertently copied over with the experiences when cutting and pasting. ' +
                      'All experience entry formatting has been completed, ' +
                      'but you will need to check all entries in your output file and remove it manually.\n')
                full_exp_split_list.append(exp)
            else:
                title_split_list = exp.split(' at ', 1)
                title = title_split_list[0]
                company, period = title_split_list[1].rsplit(' · ', 1)
                exp_split_list = [title, company, period]
                full_exp_split_list.append(exp_split_list)
    except IndexError:
        print ('\nWarning:\nThere is an entry in one of the work experiences that is not in the proper format. ' +
               'All experiences must be in the form [title] at [company] · [period]. ' +
               'Any entries that are not in the proper format will appear as blank cells in the returned file.\n')
    
    # Rebuild the work experience string (exp_string) in a more readable way
    exp_string = ''

    for idx, exp in enumerate(full_exp_split_list):  
        # Format the first experience in the list and add them to the new string
        if idx == 0:
            if exp == 'Profile experience':
                exp_string = exp_string + exp
            else:
                exp_string = exp_string + '- ' + exp[1] + '\n' + '  · ' + exp[0] + ' (' + exp[2] + ')'
        
        # Format all other experiences in the list and add them to the new string
        else:
            if full_exp_split_list[idx][1] == full_exp_split_list[idx-1][1]: # current and previous company check
                new_string = '  · ' + exp[0] + ' (' + exp[2] + ')'
                exp_string = exp_string + '\n' + new_string
            else:
                new_string = '- ' + exp[1] + '\n' + '  · ' + exp[0] + ' (' + exp[2] + ')'
                exp_string = exp_string + '\n' + new_string
                
    return exp_string

# Format the work experience
df[work_exp_column_name] = df[work_exp_column_name].apply(format_exp)

# Education formatter function
def format_edu (input_edu_string):
    """Takes a education input string, as cut-and-pasted from LinkedIn Recruiter results page, 
    and puts into a more reader friendly format.
    """
    # Split the education string into a list at each new line
    edu_list = input_edu_string.split('\n')
    
    # Initiate a string to hold all the education in a new format
    edu_string = ''
    
    # Rebuild the education string (edu_string) in a more readable way. 
    # Note that for education entries in LinkedIn 'school' is the only required field, others are optional. 
    # If the optional fields are missing, question marks (?) will be added to indicate that. 
    # In some cases of unregistered school names, it seems the school name sometimes does not appear in the listing, 
    # and only the degree and/or period, if entered, show up.  
    # These different cases are all handled below. 
    for idx, edu in enumerate(edu_list):
        # Print a warning if "Profile education" is present in the education entry.
        if edu == "Profile education":
            print('\nWarning:\nThe text string \"Profile education\" was found in one of the education entries. ' +
                  'This string is sometimes inadvertently copied over with the education when cutting and pasting. ' +
                  'All education entry formatting has been completed, ' +
                  'but you will need to check all entries in your output file and remove it manually.\n')
        
        if (',' in edu) and ('·' in edu): # All parts of the education list exist
            school, degree_period_list = edu.rsplit(', ', 1) # rsplit used to account for comma in school name case
            degree, period = degree_period_list.split(' · ')
            if idx == 0:
                edu_string = edu_string + '- ' + degree + ': ' + school +  ' (' + period + ')'
            else:
                edu_string = edu_string + '\n' + '- ' + degree + ': ' + school +  ' (' + period + ')'

        elif (',' in edu) and ('·' not in edu): # The school AND degree exist, but not the period
            school, degree = edu.rsplit(', ', 1) # rsplit used to account for comma in school name case
            if idx == 0:
                edu_string = edu_string + '- ' + degree + ': ' + school +  ' (????)'
            else:
                edu_string = edu_string + '\n' + '- ' + degree + ': ' + school +  ' (????)'

        elif (',' not in edu) and ('·' in edu): # The school OR degree and period exist
            # Cannot easily distinguish between school and degree 
            # so will just output whichever one is there.
            # There are also some rare cases in which only the period exists. 
            # In these cases a value error results when the split below is performed. 
            # In these cases we can add it in the proper format when the error is caught. 
            error_flag = False
            try:
                school_or_degree, period = edu.split(' · ')
            except ValueError:
                error_flag = True
                
            if error_flag:
                period = edu.lstrip('· ')
                if idx == 0:
                    edu_string = edu_string + '- '  + '??' + ': ' + '??' +  ' (' + period + ')'
                else:
                    edu_string = edu_string + '\n' + '- '  + '??' + ': ' + '??' +  ' (' + period + ')'
            
            else:               
                if idx == 0:
                    edu_string = edu_string + '- ' + school_or_degree +  ' (' + period + ')'
                else:
                    edu_string = edu_string + '\n' + '- ' + school_or_degree +  ' (' + period + ')'

        else: # Only the school or degree exists (only period case is handled above)
            school_or_degree = edu
            if idx == 0:
                edu_string = edu_string + '- ' + school_or_degree +  ' (????)'
            else:
                edu_string = edu_string + '\n' + '- ' + school_or_degree +  ' (????)'
        
    return edu_string


# Format the education
df[education_column_name] = df[education_column_name].apply(format_edu)

# Function to add an initial quote at the start of each education and experience, 
# so excel recognizes them as text not formulas.
def add_quote(input_string):
    input_string = '\'' + input_string
    return input_string

# Add the quotes
df[education_column_name] = df[education_column_name].apply(add_quote)
df[work_exp_column_name] = df[work_exp_column_name].apply(add_quote)

# Build the full export filepath including a new 'output' file
path, filename = os.path.split(filepath)
export_filename = filename.split('.')[0] + '_output' + '.xlsx'
export_filepath = os.path.join(path, export_filename)

# Export (to original filepath)
df.to_excel(export_filepath, encoding='utf-8')

# End program prompt
input('Export complete.\n' +
      'The exported file has the same original filename plus _output at the end.\n' +
      'Hit enter to close the program.')
sys.exit(0)