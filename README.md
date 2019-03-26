<h1 align="center">
  <br>
  <img src="https://raw.githubusercontent.com/tinyqubit/Lenovo_BatchWarranty_EndDates/master/Images/lenovo_logo.png" alt="Reddit" width="500">
  </br>
  Lenovo: Warranty Batch End Dates Lookup
  <br>
</h1>

<p align="center">
  <a href="#instructions">Instructions</a> •
  <a href="#instructions-getting-started">Getting Started</a> •
  <a href="#instructions-subreddit_onlineusers_collector_py">Online Users</a> •
  <a href="#future-features">Future Features</a>
</p>

# Purpose
Automate the lookup process of Lenovo product warranty's end dates. This allows you to add as many Lenovo product warranties, and gives you the end dates for each warranty all at once.

## Instructions
**Step 1**:
You'll want to get your Lenovo Client ID (or Token ID). You'll have to get in touch with Lenovo to see what the process would be for you to get your own Client ID depending on your needs.

Once you have your Client ID, you can add it here:
```
$TokenID = "TOKEN_ID"
```

**Step 2**:
Add your list of warranties to a .xlsx file. You can create one in Excel, Google Sheets, or Apple Numbers by exporting the file to a .xlsx file. We will be using this file to write the end dates to.
*Make sure you add in a column name so it can be searched by the script. It should look like this:*

<p align="center">
<img src="https://raw.githubusercontent.com/tinyqubit/Lenovo_BatchWarranty_EndDates/master/Images/example_2.png" alt="Reddit" width="300">
</p>

**Step 3**:
Export the same file to a .csv. This will be used to gather all the warranties. Make sure both files use the same name.
*This will be changed in the future, where you will only need one .csv. This was created with excel files in mind. This will be changed in the future to make this process more universal.*

**Step 4**:
Edit these variables.
```powershell
# File paths.
(string) $csvFilePath
(string) $excelFilePath

# The name of the column in your .csv file. Look at Step 2 for reference.
(string) $nameOfSerialNumbersColumn

# What column number do you want to write your end dates to? NOTE! Column 1 DOES NOT start with 0. Column 1 equals 1.
(int) $excelWriteEndDateRow

# From Step 1.
(string) $tokenID
```

Once these steps are complete, you can run the program.

## Example Output
<p align="center">
<img src="https://raw.githubusercontent.com/tinyqubit/Lenovo_BatchWarranty_EndDates/master/Images/example_3.png" alt="Reddit" width="300">
</p>
