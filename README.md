# OutlookEmailMenu
Creates a menu which helps sorting emails.


### Purpose of this project
Some people tend to create a huge ammount of folders with subfolders to properly sort the emails. This results in an unacceptable waste of time scrolling the folders on the left side of Outlook. This outlook project aims to simplify this task by creating a menu that can be called as a pop-up (it is a userform). An initial set-up is required the first time to tell Outlook which are the paths and to upload the images to be used.

### On clic sorting
There are two types of functions to sort the emails. 
Type 1: Directly move to a hardcoded folder. This can be used to create a button for each of the top used folders
Type 2: If there is a folder with many subfolders it may be convinient to search for a specific folder given an input string (hashtags). For example, if one has a folder with claims and one claim per customer and product, then it can be convenient to call this function, input the customer name and then get a menu with the folders that match that customer's name. The macro loops all subfolder in the specified folder and compares it to the hashtags to be searched.

# Installation Steps

## 1. Installing the package
This step depends on whether a VBA project is already present. 

### Case A: no project
Drop the project directly on the Outlook folder. This usually can be found here:
C:\Users\xxxxxx\AppData\Roaming\Microsoft\Outlook

where xxxxxx is the username of the current user.

### Case B: import modules and userforms
1. Open Outlook
2. Open the VBA editor from inside Outlook
3. Import all provided modules and userforms
  i.   UF_FolderMenu
  ii.  UF_MainMenu
  iii. MainMenu_Emails


## 2. Customizing 
1. Open Visual Basic in Outlook
2. Double clic on MainMenu_Emails
3. For each folder to which emails must directly be moved (case 1) or in which subfolders are to be looped (case 2):
    i.   Create a button  
    
    ii.  Customize button (size, image, etc.) 
    
    iii. Double clic and add code to the folder (see below)  
    
    
### Case 1: directly move to a given folder
Use function "MoveToFolder". Arguments for the function are:
1. An Array with all the folders and subfolders
2. Optional argument: if the folders are inside the Inbox, introduce the username of email account of the inbox

Example A. Folder PST_Archive > Claims > RearDoor  
Call MoveToFolder(Array("PST_FolderName","Subfolder1", "Subfolder2"))  

Example B. Move to the subfolder "Private" contained in the Inbox of user afischer  
Call MoveToFolder(Array("PST_FolderName","Subfolder1", "Subfolder2"),"AdrianFischer")

### Case 2: loop subfolders of given subfold
It has the same logic as Case 1, but the function to be used is SuggestFolders:  
Call SuggestFolders(Array("PST_FolderName","Subfolder1", "Subfolder2"))  
Call SuggestFolders(Array("PST_FolderName","Subfolder1", "Subfolder2"),"AdrianFischer")  


### Special comment for Inbox
There are two ways to assign the Inbox. Besides the mentioned method of using the optional Inbox parameter. Alternatively, it is also possible to directly use the path and inputing the email as first Folder. For example: 
Call MoveToFolder(Array("adrian.fischer@dummycompany.com","Inbox","Subfolder1")) 

is equivalent to

Call MoveToFolder(Array("Subfolder1"),"AdrianFischer") 




## 3. Add button to Outlook
An icon can now be added to the Quick Access Toolbar or the Ribbon.
File > Options > Quick Access Toolbar
File > Options > Customize Ribbon

I suggest using the Quick Access Toolbar. Later it can be called pressing ALT + icon number. 

Eventually Outlook will distrust the macro project. You might have to self-sign the project to prevent Outlook from blocking it.
More information here: https://support.microsoft.com/en-us/office/digitally-sign-your-macro-project-956e9cc8-bbf6-4365-8bfa-98505ecd1c01

### Disclaimer
I initially wrote this function for myself some years ago. Thus I commented most of the code in my mother tongue, Spanish. 
Since it was my first attempt using VBA for Outlook it is rather unefficient. 
Recently I was asked by a colleague to share the code with him. 
As a result I re-sorted and cleaned the code partially, built in some new functions (such as moving files to subfolders of the Inbox) and translated some functions and comments to English.
It is still a bit redundant sometimes, so questions and suggestions are welcomed.
