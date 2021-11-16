### Links

[csv module](https://docs.python.org/3/library/csv.html)

---

## Install Instructions (Windows 10)

1. Download [Git for Windows](https://gitforwindows.org/)
...
Run 'pip install pywin32'
git submodule add https://github.com/matthew-webber/my_modules my_modules


```bash

+--------------------------------------+
|                                      |
|    Hello and welcome to the          |
|  SCA "Membership Email Generator"!   |
|                                      |
|  To begin, simply follow the prompts |
|   after ensuring you have read and   |
|     understand the notice below.     |
|                             .        |
|          Happy mailing!      .       |
|                       ><((('>        |
+--------------------------------------+

           =====NOTICE=====

* Make sure you have the .csv file in
  the same folder as this script.

* This script will start with the record
  on row 2 and generate a
  "giver" email and "recipient" email
  for each record.

* It will iterate over 3 records
  at a time, pause after each iteration,
  and ask to continue until you tell it
  to stop or until it reaches the last
  record.

* You can change which row the script
  starts on and how many records it
  will process at a time, or solely
  the row it starts on.  You cannot
  change solely the number of records
  to process each iteration.  See the
  examples below.


Examples:
(running from command prompt - start on row 10,
  process 5 records each iteration):

    python path/to/this/script 10 5 (Mac)
    python "C:\Users\your_home_folder_here\path\to\this\script" 10 5 (Windows)

(running from command prompt - start on row 12,
  process default (3) records each iteration):

    python path/to/this/script 12 (Mac)
    python "C:\Users\your_home_folder_here\path\to\this\script" 12 (Windows)


--------------------------------------

Starting row: 2
emails at a time: 3

Right now, the emails you generate are set to
save in your drafts folder as they are created.

Enter "display" to change this, q to "quit",
or just press Enter to start.

?:

```