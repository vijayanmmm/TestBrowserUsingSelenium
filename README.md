# TestBrowserUsingSelenium
This is the sample project along with the helper file to automate/test web sites

********************TOOLS Used 
  Language - C#.Net     
  MicroSoft Visual Studio Express 2013
  Selenium Web Driver
  AutoItx
  HtmlAgilityPack

********************PURPOSE
  This is the sample project about how to use the selenium with google chrome and two helper classes to speed up your coding


********************Targetted Users
 For beginners who try to use selenium
 For experienced users -
    the class "clsAutomation.cs" & "Helpers.cs"
    
********************ClsAutomation.cs
This class consist of several methods which are helpful to speed up the automation. This class consist of settext,gettext, selecttext,
and working with html tables.

********************clsHelpers.cs
This class consist of several methods which are helpful in 
  1. Excel to data table conversion
  2. Data table to excel
  3. string functions like getalpha, getnumeric..etc, this also has functions to compare a text with 85% match....
  4. Outlook auto click allow button - if you are in a situation that when you send emails from outlook using coding and its
  displaying warning where we need to click allow every time/email. (Using Auto ITx)
  5. Browser auto click login - If you are in a situation that after your browser a URL, it asking authendication details 
  then selenium struck there waiting for you to update manually. In this case the helper functions will click ok automatically.
  (make sure you have already entered and clicked remember option in chrome)   (Using Auto ITx)
