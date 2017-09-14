# Send Mail from Spreadsheet using Google Apps Script(GAS)
This is a Google App script(GAS) that send your mail data from a google spreadsheet.


## Before you begin
  Create a Google Spreadsheet with three seets.

1 "template" sheet is a mail template. 
   * It needs to have a button to start the script. The initial script name is 'click'. You can create this button from "Insert -> Drawing" on the toolbar. Click the righ mouse button on the button, and assign a script name 'click'.
   * Import mail.gs script from 'Tools -> Script editor' on the toolbar.
![](readme_images/sheet1.PNG)

2 "rcptst" sheet is used to manage recipients.
![](readme_images/sheet2.PNG)

3 "log" sheet is used to store sent history.
![](readme_images/sheet3.PNG)
   
   
## Usage
1 Create a templete for mail on "template" sheet.

2 Insert recipients you want to send mail on "rcpts" sheet.

3 Click the button on "template" sheet. After the operation check the dialogue box to confirm selected operations.

That's all. The transmission history are stored in "log" sheet.


## License
The script is available as open source under the terms of the MIT License.


## Authors
http://alpha-netzilla.blogspot.com/
