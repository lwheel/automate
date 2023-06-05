import pandas as pd

# Read the Excel file
df = pd.read_excel("/Users/lilywheeler/Desktop/students.xlsx", engine="openpyxl")
print(df.columns)
# Extract the relevant information for each student
students = df['Student ']
countries = df['Country ']
genders = df['Gender ']
q1 = df['Q_1']
q2 = df['Q_2']
q3 = df['Q_3']
q4 = df['Q_4']
q5 = df['Q_5']
totals = df['Total ']
percentages = df['Percentage']

# Iterate over each student's data
for i in range(len(students)):
    # Extract the relevant information for each student
    student_name = students.loc[i]
    country = countries.loc[i]
    gender = genders.loc[i]
    q1_value = q1.loc[i]
    q2_value = q2.loc[i]
    q3_value = q3.loc[i]
    q4_value = q4.loc[i]
    q5_value = q5.loc[i]
    total = totals.loc[i]
    percentage = percentages.loc[i]

    # Generate the HTML content for the student's webpage
    html_content = f'''
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head><meta http-equiv="content-type" content="application/xhtml+xml; charset=UTF-8" /><link href="App_Themes/fav-logo.ico" rel="shortcut icon" type="image/x-icon" /><link href="App_Themes/design.css" rel="stylesheet" type="text/css" /><link href="App_Themes/print.css" rel="stylesheet" type="text/css" media="print" /><title>
	International Mathematical Olympiad
</title>
<style>
</style></head>
<body>
    <div id="header">
        <div id="h1">
            <h1 class = "centered-heading"><a href="index.html">East African Mathematical Olympiad</a></h1>
        </div>
        
        <div id="sub">   
            <span class="centered-image"><a href="index.html"><img src="https://www.imo-official.org/logo/IMOLogo_30px.png" alt="IMO" style="width: 100px; height: 66px;"/></a></span>
      
        </div>
      
        <div id="menu">
            <a id="ctl00_HyperLink_Countries" href="countries.html">Countries</a> &bull;
            <a id="ctl00_HyperLink_Results" href="results.html">Results</a> &bull;
            <a id="ctl00_HyperLink_Problems" href="problems.html">Problems</a> &bull;
            
            <a id="ctl00_HyperLink_Links" href="links.html">Links and Resources</a>
        </div>
      </div>



    <div id="main">
        
<h2>{student_name}</h2>

<table><thead><tr><th rowspan="2" class="highlightDown"><a href="kp5.html">Year</a></th><th rowspan="2"><a>Country</a></th><th rowspan="2"><a href="kp5.html">P1</a></th><th rowspan="2"><a href="kp5.html">P2</a></th><th rowspan="2"><a href="kp5.html">P3</a></th><th rowspan="2"><a href="kp5.html">P4</a></th><th rowspan="2"><a href="kp5.html">P5</a></th><th rowspan="2" style="display: none;"></th><th colspan="2" rowspan="2"><a href="kp5.html">Total</a></th><th rowspan = "2">Rank</th><th rowspan="2"><a href="kp5.html">Award</a></th></tr></tr></thead><tbody>
<tr class="imp"><td align="center"><a>2023</a></td><td><a href="{countries}.html">Kenya</a></td><td align="center">{q1}</td><td align="center">{q2}</td><td align="center">{q3}</td><td align="center">{q4}</td><td align="center">{q5}</td><td align="center" style="display: none;"></td><td align="right">{total}</td><td align="right">{percentage}</td><td align="right">24</td></td><td>Bronze medal</td></tr>
</tbody></table>

<p>Results may not be complete and may include mistakes.
Please send relevant information to the webmaster: <a id="ctl00_CPH_Main_HL1" href="mailto:ea@globtalent.org">ea@globtalent.org</a>.</p>


    </div>

</body>
</html>
    '''

    # Save the HTML content to a file
    filename = f'{student_name.replace(" ", "_")}.html'
    with open(filename, 'w') as file:
        file.write(html_content)

    print(f'Successfully created webpage for {student_name}')

print('Webpage creation completed.')