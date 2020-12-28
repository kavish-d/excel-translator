## A python utility to convert xlsx and xls files from one language to another while trying to keep as much as possible struture and formatting.


To run First time do 
` pip install -r requirements.txt `


Run as 
`python translator.py `


It will ask some option

`Mode` : mode of conversion 
There are 3 modes  for xlsx file as they are zips with proprieatry xml data for formatting some things are not translated properly
Mode 1 tries to keep as much formatting as possible butsince all xlsx can't be handled
it generates 2 target files one is translated xlsx with formatting ( in case it doesn't open try renaming extension to xls or using any other software than excel)
other is compatibility file whichs remove formatting and maintain cell position and text in a xls file

Mode 2 will give a xlsx file which will definitely open in excel though formatting is not sure in it

Mode 3 first convert xlsx to xls keeping only the formatting that is backward compatible by ms and work on xls and is fastest Use this if file type is xls

Mode 1 is preferred for use

`Source and destiantion language code` is code used by google translate  

`es` for spanish
`en` for english
more search on google

Last input is for `extension` which to be converted xlsx or xls
Program will crawl in current directory and all it's subdirectory for files of that extension and translates them



Sheets that need to be translated there names should be in `SHEETS` constant in python file


In case there is error from googletrans library package It is likely that there were too many request from your ip in that try using vpn or install `cloudfare warp`
