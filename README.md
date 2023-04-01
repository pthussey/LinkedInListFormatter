# LinkedInListFormatter

This program takes an Excel file that has work experience and education information copy-and-pasted from LinkedIn Recruiter search results 
and puts it into a more readable format.

It works only on text copied from paid LinkedIn account (Premium or Corporate Recruiter) search results pages because the text is separated according to set rules.  
Work experience entries are always written as [Job title] at [Company name] · [Period].  
Education entries are a bit more complicated because some of the fields are optional but these are seperable as well.

This program uses the separation rules to split the text into parts and then recombine it in more readable format.

Below is an example of how it works.

Search results example in LinkedIn Recruiter:  
<img src="https://raw.github.com/pthussey/LinkedInListFormatter/main/assets/images/Paul_Hussey_search_results.png" width="470" height="270"/>

If you copy the text from the Experience and Education sections of Recruiter search results and paste it into columns in an Excel file, then you can run this LinkedInListFormatter program on the Excel file to format the text.

Example of search results text copied-and-pasted into an Excel file:  
<img src="https://raw.github.com/pthussey/LinkedInListFormatter/main/assets/images/excel_before_conversion.png" width="528" height="166"/>

Once you have all the search results you need copied into an excel file, you can then run the program to do the formatting.  
If you have Python installed on your machine, you can run it using the Python file (pandas and openpyxl packages are also required).
Or, if you don't have Python installed, you can run it using the executable file provided in this repository. 

When running the program just follow the intructions at the prompts.

<img src="https://raw.github.com/pthussey/LinkedInListFormatter/main/assets/images/enter_path.png" width="410" height="18"/>  

<img src="https://raw.github.com/pthussey/LinkedInListFormatter/main/assets/images/enter_filename.png" width="410" height="19"/>

<img src="https://raw.github.com/pthussey/LinkedInListFormatter/main/assets/images/enter_work_experience.png" width="470" height="20"/>

<img src="https://raw.github.com/pthussey/LinkedInListFormatter/main/assets/images/enter_education.png" width="440" height="19"/>

Then your Excel output file will be produced.  
If you select the columns that hold the work experience and education text in your output file and enable "Wrap Text" you will be able to see the full text.

The end result will look like this:  
<img src="https://raw.github.com/pthussey/LinkedInListFormatter/main/assets/images/excel_after_conversion.png" width="685" height="177"/>
