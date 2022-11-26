"""Reads an excel file that contains education and work experience 
cut-and-pasted from LinkedIn Recruiter results lists, 
and outputs the education and work experience entries in a more readable format. 
Any entries that start with a dash(-) are considered to be already formatted and are left unchanged. 
Note that the education output can have some arrangement problems that need to be fixed 
in cases when the original input string has commas used in the degree name.
"""

import pandas as pd
import openpyxl
import sys
import os

print('This program accepts a csv or txt file ' +
      'with work experience and education entries ' +
      'cut-and-pasted from LinkedIn Recruiter results pages, ' +
      'puts the entries into an easier to read format, ' +
      'and exports it as an excel file. ' + 
      'Any entries that start with a dash(\'-) ' +
      'are considered to be already formatted ' +
      'and are left unchanged.\n')

# Accept the path and filename and build the filepath to pass into pandas
path = input('Enter the path to the folder that contains your file: ')
filename = input('Enter the filename including the filetype extension: ')
filepath = os.path.join(path, filename)

# Check that extension is included and is csv or txt
try:
    extension = filename.split('.')[1]
except IndexError:
    print('\nError:\nPlease ensure that the file extension ' +
          'is included in the filename and try again.')
    input('Press enter to exit.')
    sys.exit(1)

if extension not in ['csv','txt']:
    print('\nError:\nThe file type must be either csv or txt. ' +
          'Please ensure the file is in the proper format and try again.')
    input('Press enter to exit.')
    sys.exit(1)

# Accept the delimiter type for the data to be passed to pandas
delimiter = input('Enter the delimiter type of your data. ' +
                  'Only \'comma\' and \'tab\' are accepted options: ')

if delimiter not in ['comma','tab']:
    print('\nError:\nThe delimiter must be either \'comma\' or \'tab\'. ' +
          'Please try again and select the right delimiter for your data.')
    input('Press enter to exit.')
    sys.exit(1)

# Read the file and handle file / directory exception.
try:
    if delimiter == 'csv':
        df = pd.read_csv(filepath, encoding='utf-8')
    else:
        df = pd.read_csv(filepath, delimiter = '\t', encoding='utf-8')

except FileNotFoundError:
    print('\nFolder or Filename Error:\nThe folder path and/or filename entered does not exist.\n' +
          'Folder path format example: C:\\Users\\username\\foldername\n' +
          'Filename format example: filename.txt\n' +
          'Please check and try again.')
    input('Press enter to exit.')
    sys.exit(1)

work_exp_column_name = input('Enter the column name that contains ' +
                             'the work experience entries: ')

# Replace carriage returns + line feeds with just line feeds for the work experience.
# Also check the column entries while doing this.
try:
    df[work_exp_column_name] = df[work_exp_column_name].str.replace('\r\n', '\n')

except KeyError:
    print('\nKey Error:\nThe column name provided for the work experience entries does not exist. ' +
          'Please check the name and try one more time.')

    # One more chance to enter the work experience column name
    work_exp_column_name = input('Enter the column name that contains ' +
                                 'the work experience entries: ')

    try:
        df[work_exp_column_name] = df[work_exp_column_name].str.replace('\r\n', '\n')
    
    except KeyError:
        print('\nKey Error:\nThe column name provided for the work experience entries does not exist. ' +
              'Please check the name and try again.')
        input('Press enter to exit.')
        sys.exit(1)

education_column_name = input('Enter the column name that contains ' +
                              'the education entries: ')

# Replace carriage returns + line feeds with just line feeds for the education.
# Also check the column entries while doing this.
try:
    df[education_column_name] = df[education_column_name].str.replace('\r\n', '\n')

except KeyError:
    print('\nKey Error:\nThe column name provided for the education entries does not exist. ' +
          'Please check the name and try one more time.')
    
    # One more chance to enter the educaion column name
    education_column_name = input('Enter the column name that contains ' +
                                  'the education entries: ')
    
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
    
    # If cells containing an initial single quote are clicked on excel, 
    # those single quotes are no longer recognized at import into pandas.
    # So need to add it back if the first character is a dash, 
    # as it will be in the ones that are already formatted.
    try:
        if input_exp_string[0] == "-":
            input_exp_string = "'" + input_exp_string
    except IndexError:
        pass

    # Skip ones that have already been formatted by recognizing the starting pattern ('- )
    try:
        if input_exp_string[0:3] == "\'- ":
            return input_exp_string
    except IndexError:
        pass

    # Split the work experience string into a list at each new line
    exp_list = input_exp_string.split('\n')
    
    # For each experience in the experience list, 
    # pull out the title, company, and period and put them in a list, 
    # then append each list to a full experience list, 
    # creating a nested list.
    full_exp_split_list = []
    
    try:
        for exp in exp_list:
            # Handle cases of "Profile experience" that are sometimes inadvertently copied over in the string.
            if exp == "Profile experience":
                continue
            
            # Replace cases of 'no backspace' characters, somehow added by Excel in some cases, with regular spaces.
            if '\xa0' in exp:
                exp = exp.replace('\xa0', ' ')
            
            # Handle cases of additional line feeds inadvertently added that create empty string entries.
            if exp == '':
                continue

            else:
                title_split_list = exp.split(' at ', 1)
                title = title_split_list[0]
                company, period = title_split_list[1].rsplit(' · ', 1)
                exp_split_list = [title, company, period]
                full_exp_split_list.append(exp_split_list)
    except (IndexError, ValueError):
            print ('\nWarning:\nThere is an entry in one of the work experiences that is not in the proper format. ' +
               'All experiences must be in the form [title] at [company] · [period]. ' +
               'Any entries that are not in the proper format will appear as they were with no changes made.\n')
            return input_exp_string
    
    else:
        # Rebuild the work experience string (exp_string) in a more readable way
        exp_string = ''
        
        for idx, exp in enumerate(full_exp_split_list):  
            # Format the first experience in the list and add them to the new string
            if idx == 0:
                # Remove any initial quotes that might exist at the start of the entry
                if exp[0] == "'":
                    exp = exp.lstrip("'")
                
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
    """Takes an education input string, as cut-and-pasted from LinkedIn Recruiter results page, 
    and puts into a more reader friendly format.
    """

    # If cells containing an initial single quote are clicked on excel, 
    # those single quotes are no longer recognized at import into pandas.
    # So need to add it back if missing.
    try:
        if input_edu_string[0] == "-":
            input_edu_string = "\'" + input_edu_string
    except IndexError:
        pass

    # Skip ones that have already been formatted by recognizing the starting pattern ('- )
    try:
        if input_edu_string[0:3] == "\'- ":
            return input_edu_string
    except IndexError:
        pass

    # Split the education string into a list at each new line
    edu_list = input_edu_string.split('\n')

    if edu_list[0] == "Profile education":
        edu_list.remove("Profile education")
    
    # Initiate a string to hold all the education in a new format
    edu_string = ''
    
    # Rebuild the education string (edu_string) in a more readable way. 
    # Note that for education entries in LinkedIn 'school' is the only required field, others are optional. 
    # If the optional fields are missing, question marks (?) will be added to indicate that. 
    # In some cases of unregistered school names, it seems the school name sometimes does not appear in the listing, 
    # and only the degree and/or period, if entered, show up.  
    # These different cases are all handled below. 
    for idx, edu in enumerate(edu_list):

        # Remove any initial quotes that might exsit at the start of the entry
        if idx ==0:
            try:
                if edu[0] == "'":
                    edu = edu.lstrip("'")
            except IndexError:
                pass
        
        # Replace cases of 'no backspace' characters, somehow added by Excel in some cases, with regular spaces.
        if '\xa0' in edu:
            edu = edu.replace('\xa0', ' ')

        if (',' in edu) and ('·' in edu): # All parts of the education list exist
            try:
                school, degree_period_list = edu.rsplit(', ', 1) # rsplit used to account for comma in school name case
                degree, period = degree_period_list.split(' · ')
            except ValueError:
                print('\nWarning:\nThere is an education entry that is not in the proper format. ' +
                      'Any entries that are not in the proper format will appear as they were with no changes made.\n')
                return input_edu_string
            
            if idx == 0:
                edu_string = edu_string + '- ' + degree + ': ' + school +  ' (' + period + ')'
            else:
                edu_string = edu_string + '\n' + '- ' + degree + ': ' + school +  ' (' + period + ')'

        elif (',' in edu) and ('·' not in edu): # The school AND degree exist, but not the period
            try:
                school, degree = edu.rsplit(', ', 1) # rsplit used to account for comma in school name case
            except ValueError:
                print('\nWarning:\nThere is an education entry that is not in the proper format. ' +
                      'Any entries that are not in the proper format will appear as they were with no changes made.\n')
                return input_edu_string
            
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

# Function to add an initial quote at the start of education and experience, 
# so excel recognizes them as text not formulas.
def add_quote(input_string):
    input_string = '\'' + input_string
    return input_string

# Add initial quotes to any newly added entry starting with a dash 
# that doesn't already have an initial quote
edu_mask = df[education_column_name].str.startswith("-", na=False) == True
df.loc[edu_mask, education_column_name] = df.loc[edu_mask][education_column_name].apply(add_quote)

exp_mask = df[work_exp_column_name].str.startswith("-", na=False) == True
df.loc[exp_mask, work_exp_column_name] = df.loc[exp_mask][work_exp_column_name].apply(add_quote)

# Build the full export filepath including a new 'output' file
export_filename = filename.split('.')[0] + '_output' + '.xlsx'
export_filepath = os.path.join(path, export_filename)

# Export (to original folder)
try:
    df.to_excel(export_filepath)
except PermissionError:
    print('Error: The file you are trying to write to is open. ' +
          'Close the file and try again.')
    input('Press enter to exit.')
    sys.exit(1)

# End program prompt
input('Export complete.\n' +
      'The exported file has the same original filename plus _output at the end, ' +
      'and it is now in excel format.\n' +
      'Press enter to close the program.')
sys.exit(0)