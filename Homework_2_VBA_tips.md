# Homework 2 VBA tips

Here are some notes that I have sent out to differerent people over the past week that may be helpful - mostly related to the homework.

**VBA Day 1 Summary:**  

1.  So I suggest that at the end of the class you write down what you think the most important parts of today's class and does any of it relate to the homework.  Here is what I would take notes of today:

2.  `Activity 6` shows ways of doing cell referencing (Range, Cells)
3.  `Activity 7` shows an example of using the Long data type.  There is also a link in TTh about the data types and their limits.  We said an integer max value goes a bit above 32,000.  If you go over 32,000 in your code then it will give you an error about data size.
4.  `Activity 8` shows how to get values from the sheets and set them as a variable and it shows how to write values/calculations/strings back to a cell in a sheet.
5.  `Activity 12` shows if statements so it tells you how to compare two values and then execute code only for that situation.
6.  Also when dealing with decimals there is the data type Single and Double.  Double works for a larger set of numbers.  If you use single then you could get an error if you exceed its limit. 
7.  Below is a link about `Range(...).Interior.ColorIndex = #`  color codes. I think there is an example of this in `Activity 6` where the chessboard black and white squares are created.
8.  Please note:  **`all vba code needs to be saved to a .xlsm or your vba script you wrote will be deleted if saved as another format like .xls.`**  
I note these above because I think almost each of these is related to the homework.
<br><br>  

**VBA Day 2 Summary:**  
1.  `Activity 6` is very pertinent to the homework since it is iterating down rows and doing a comparison and storing info  in a table to the side of the data.  
2.  `Activity 7` is also important because you will need to use nested loops in your homework.  
3.  We have another activity in Day 3 that will greatly help with figuring out the homework.  `Look over the homework files (instructions and what the data looks like)so you will know which activity I am talking about when we cover it.`  Maybe this will help you come up with some helpful questions.
<br><br>

**VBA Day 3 Summary:**  
1.  The early activities are helpful with formatting, especially `Activity 4` which uses formatting inside looping through cells. 
2.  `Activity 5` and `Activity 6` are very helpful for the homework for probably the most difficult part of the HW.
2.  `Activity 7` has a way of looping through sheets (referencing each sheet as `ws.`), finding the last row (there are actually multiple methods in Excel VBA but here is one), inserting columns, finding the last column 
3.  `Activity 7` and `Activity 8 Part ii` has a nice way of getting and looping through each spreadsheet tab with a `For each` - I think this is the most common way vba coders do this.
3.  As in `Activity 7 and 8` it is best to get one sheet to do what you need and then go back and get your code to work with multiple sheets (like the looping from sheet to sheet)
3.  The `Day 3 > Activities > Extra_Content` folder has some info about how to reference multiple sheets in the code.
<br><br>

**Other Notes - Citations, Git, etc.**
1.  If you work with someone then just include at the top of your VBA script a comment the tells us everyone that you worked with like this:
`'worked/discussed assignment with {first name} {last name}`
2. Cite any websites that you pull code from as a comment near the area where you use the code.  If you are just looking up a command or using instructions from the source documentation then no need to cite anything but if you find a chunk of code on stackoverflow or a general website that does what you need (like a multistep process) then give the person credit for their work - provide a URL link in the comment. 
3.  (`For Projects`) When you find datasets then you should always provide a citation of where you got the data.  It does not have to be a formal citation but say who made the dataset and where you found it and a date when you pulled it.
4.  FYI - As you get solutions for activities, you can add these into your repo manually or you can wait until the end of class I will push the solutions to the repo.  If you add the files into your solutions folder but make changes to it then you might get an error when you do a 'git pull'.  It will see that the master copy on gitlab is different than the one in your folder.  Two ways of fixing this:  1) you can save the file  with a different name in the Solved folder or you can put the file in the unsolved folder or 2) when you get the git pull error, you can do a 'git stash' after you see the error.  This will move your changes to the background and allow the file to be brought in but you won't see the changes you added to the file initially.
* **`Key Reminder`**:  Do not clone a repo (from Github or Gitlab) inside an existing repo - keep them as separate folders - independend of each other. 
* **`If your Macro Enable Excel file is over 100MB then I would suggest putting that in a google drive folder and submitting that link for the homework.  Github does  not like files over 100MB and I think that starter file is already at 96MB.`**  You can submit the code as a .vbs file or as a .bas file and any other files in Github along with a ReadMe.md file.  It would be helfpul if you put the google drive link in the ReadMe.md file and write up a few sentences about your code.
<br><br>

**Links:**
* Here are the colorIndex codes: [Link](https://www.automateexcel.com/excel-formatting/color-reference-for-color-index/)
* Here are the different data types (integer, long, etc) and their limits:  [Link](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/data-type-summary)
* To import code into the VBA IDE, go to File and Import and you can import files that have the extension renamed from .vbs to .bas.  You might get a security warning about macros so this link might be helfpul:  [Link](https://support.microsoft.com/en-us/office/enable-or-disable-macros-in-office-files-12b036fd-d140-4e74-b45e-16fed1a7e5c6?ui=en-us&rs=en-us&ad=us)
* Using the & (concatenate operator) with print statements: [Link](https://www.tutorialspoint.com/vba/vba_concatenation_operators.htm)
* Using the + str() (string casting and other examples of casting) with print statements: [Link](https://www.automateexcel.com/vba/convert-text-string-to-number/)
<br><br>

**Helpful Videos**
1.  Software Setup Checks and Git Basic Operations:  [Link](https://northwestern.hosted.panopto.com/Panopto/Pages/Viewer.aspx?id=c5a12560-9e09-4404-81d7-acd20173d4d1) 


The video covers this:
```
• Adding the Developer tab to Excel
• Making a copy of gitlab on your local machine (cloning a repo)
• Getting updates of a gitlab repo (pulling a repo)
• Using VSCode
• Adding VSCode Extensions
• Testing Python in VSCode
• Testing Python in Anaconda Prompt
• Making a Github Repo in Github and copying it to your computer (Step 1)
• Adding Files to a Github Repo on your computer (Step 2)
• Adding Files to your online Github Repo by pushing (Step 3)
```