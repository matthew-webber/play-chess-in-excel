# play-chess-in-excel
Upload these VBA files to any workbook and start playing chess, today!

![alt text](https://github.com/matthew-webber/play-chess-in-excel/blob/master/preview.png?raw=true)

## Who It's For
This can be great if you're playing chess via a shared workbook in the cloud or on a network drive.  It's also pretty handy if you're playing chess via e-mail -- just copy+paste the board to your favorite e-mail client and send!  Algebraic notation included for ease of communication.

## Get It Going
Simply 'clone' or download these files to your computer and then upload them into any workbook from the VBA menu in Excel (you only need the 3 files -- the 2 "Spawn_form" files and the "chess_main" file.

### Uploading to a workbook
To access the VBA menu in Excel, simply enter Alt + F11 while Excel is open.  Then choose 'File' > 'Import File...' and choose 'chess_main.bas' > 'Open.'  Repeat this for 'Spawn_form.frm' as well.  The .frx file needs to be in the same directory (I think), but it probably already is if you downloaded / cloned this project, so no worries.  It may not be necessary, I'm not really sure...

Once you've added the files above, simply double-click the "chess_macro" file from the VBA window and run "setup()" -- you can do this by simply clicking anywhere in the "Sub setup()" submodule and then choosing "Run" > "Run Sub/UserForm".  Alternatively, just press the F5 button to run it.  A new worksheet called "Chess!" will be created with all of the necessary buttons, logic, etc.

# If it doesn't work, you might have Excel 64-bit!  It won't work in Excel 64-bit!
## Check this link on how to check your version of Office
https://support.microsoft.com/en-us/office/about-office-what-version-of-office-am-i-using-932788b8-a3ce-44bf-bb09-e334518b8b19?ui=en-us&rs=en-us&ad=us#ID0EAAAACAAA=Windows

