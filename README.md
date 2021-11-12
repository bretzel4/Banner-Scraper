# Banner-Scraper
 
Created by Lilac Damon http://www.meltedlilacs.com

This program uses pasted info from Banner in order to compile a list of student emails and other details.

## Usage
Create a text file named `students.txt`. Copy and paste all text from the Banner page into the text file. If there are multiple pages worth of students, paste each page sequentially. Next, run `bannerscraper.exe` by double-clicking it. If everything works correctly, an Excel spreadsheet should be generated. That spreadsheet has two tabs, the results and the errors. Before using the results, look through the errors and try to resolve any you can. A file `log.txt` is also generated, but that can be ignored.

## Privacy
1) This program never directly connects to Banner. It therefore has no way to access or manipulate privileged information.
2) This program only connects to public websites (UVM directories).
3) The source code does not reveal any information about Banner other than basic information
    about where items are located on the page (ex: student names appear near 9-digit numbers).
4) All of these claims may be verified by looking at the source code.

## Contribution
This source code is a mess. It was cobbled together with no intention to extend it further. The `.exe` is created with the `auto-py-to-exe` pip package.
Todos:
- Reduce file size
- Clean up code (especially create a more encompassing DataFrame)
- Add settings for what the output should look like