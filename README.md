# Stevens Esports Application Sheet Splitter

This script takes an incoming Excel Spreadsheet full of Google Form responses for applications, and splits them up into their respective games. It asks for a single character, designating the row where applicants selected their games, and then outputs an .xlsx spreadsheet file with all of the responses divided between sheets representing games. Column-width is accounted for to promote readability.

You must provide an .xlsx file, and ensure the name is the same in the python script(see `df = pd.read_excel('file.xlsx')`)

Upon completion, the program will output a new .xlsx file called `SortedTeams.xlsx`.