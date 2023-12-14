# [Create Project Batch File]

## What is it?
A Python script that aids in creating project-specific batch files that prepare your workspace by opening everything you would open yourself using just one double-click action.

The batch files created using this script can open ***folders***, ***files***, ***programs*** and ***URLs***.

As an illustrator, I  very often  find myself opening the same stuff, in the same order, over and over again, day after day. That's where these batch files come in. 

Instead of wasting time preparing my workspace each time I sit down to work, I have begun creating batch files that open, for example, my references (PureRef files), my project folders, my programs, and so on. However, the creation of said batch files is a very manual process. 

That's why I've sat down to create this script, to make that process as simple and user-friendly as possible. But, since I'm not a programmer, this script was created with the help ***ChatGPT 3.5*** and a metric ton of elbow grease. 

At this point, as far as the functionality I require out of it, it is fully functional as is. But, I am very aware of how unoptimized and redundant some parts of it are. That's exactly why I'm sharing it on GitHub in the hope of it becoming the best it can be. Feel free to rewrite it in any form you think would be better suited to the task at hand.

## How does it work?

This script works by opening an always-on-top window that you can drag and drop everything you would open yourself into. After you're done with that, you just click on the "Write to file and exit" button and it will create a batch file that will open everything you selected for you in the order you specified. 

It also features a Dark Mode that is persistent across startups by utilizing the config.ini file to store the setting.

![[Create Project Batch File] - Light Mode](https://github.com/BoundToMakeArt/Create-Project-Batch-File/blob/main/Images/%5BCreate%20Project%20Batch%20File%5D%20_-_Light_Mode.png)    ![[Create Project Batch File] - Light Mode](https://github.com/BoundToMakeArt/Create-Project-Batch-File/blob/main/Images/%5BCreate%20Project%20Batch%20File%5D%20_-_Dark_Mode.png)

When the created batch file has completed opening everything you've specified, it will notify you using a toast notification.

![[Create Project Batch File] - Toast Notification](https://github.com/BoundToMakeArt/Create-Project-Batch-File/blob/main/Images/%5BCreate%20Project%20Batch%20File%5D_-_Toast_Notification.png)

For larger PureRef files (500mb+), the Python script adds 32 second increments for each 500mb of a PureRef file's size to the delay in the batch file. This is done to make sure that large PureRef files have time to fully load before the script notifies you that it has completed opening everything.

Icon for the toast notification acquired from [Icon Link](https://www.flaticon.com/free-icon/open-box_869078)

## Requirments

Being a Python script, it requires Python to be installed before use. You can download it from [Python.org](https://www.python.org/)

## Dependencies

As far as the dependencies go, this script uses the following Python libraries: ***pyqt5***, ***pywin32***, ***pywin32-ctypes*** and ***win10toast***.

But, you **don't have to install anything yourself!**

I've created a ***"[Create Project Batch File] [Install Script Dependencies]"*** script that installs all the dependencies for you!

After you've run that script once, you can delete it.

## Usage
- Make sure that the "[Create Project Batch File].py" script and the "[Create Project Batch File] Util" folder are in the same directory.
- Start the "[Create Project Batch File].py" script
- Drag and drop any of the following into the main window: folders, files, programs and/or URLs (or use the "Select Files/Folders" button to start the file picker)
- Click on the "Write to file and exit" button
- Rename the newly created "open_prepared_items.bat" file to the name of your project
- Store the .bat file where your project is stored
- Each time you're ready to work, just double-click the .bat file and grab a drink as the batch file opens everything for you, preparing your wokrspace
- If you ever move the script and its Util folder to a new directory (they must remain together in the same directory for the script to work), just make the batch file again using the same process and it will automatically update the path to the toast notification script (that's how it displays toast notifications)
