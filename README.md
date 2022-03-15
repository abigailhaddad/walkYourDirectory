# walkYourDirectory
this is for pulling text/metadata from all of the files in a folder/subfolders

This walks through all of the files in your folder and its subfolders and returns a data frame containing information on each of those files: its location, creator, who last modified it, and the file extension
if textPull=True, which is the default, this will also get additional info for xlsx, docx, ppt, and pdf files: created on/modified on dates as well as the actual text of that document
you can also modify the wordlist function to search for particular strings within the text you've pulled out
the formulas parameter, if you set it True, will specifically pull out the formulas in the text of Excel files (if you've left textPull to true) and put those in a separate column of the final data frame
