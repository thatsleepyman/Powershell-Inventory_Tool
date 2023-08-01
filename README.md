# Powershell-Inventory_Tool
This is a Powershell script I wrote for a summer job I had. It was used to easily take inventory of old Desktops that had to be replaced.

First off, these scripts were made by Stefan Meeuwessen to make taking Inventory of computers a lot easier.

Now off to the good part...

*NOTE:*
To be able to run Powershell scripts on any Computer be sure to open Powershell as Administrator and then open the scripts in Powershell ISE.

- First off run the script `Set-ExecutionPolicy.ps1` to allow scripts to be used on this machine.
> WARNING, this will make it possible to run ANY script! This can be very dangerous if abused!

- Then run `Inventory_Tool.ps1` to get the computer's data.
- This will create a folder with all your data in the same Directory as the `Inventory_Tool.ps1` script

Do with that data what ever you wish.

This is one of my first Powershell scripts so hopefully this script will have been useful to someone.
Also its totally fine to use for private or commercial reasons or to change the script to your liking and or needs.
Though it would be nice to get a __smidgen__ of credit for the script :)

Also, if you're curious on the development of this script, you might be interested in this Reddit post I made a long time ago:
[Reddit Post](https://www.reddit.com/r/PowerShell/comments/izfjh0/so_im_a_new_ps_user_and_i_made_this_script_mind/)


Powershell-Inventory_Tool
Overview

This repository contains a Powershell script created by Stefan Meeuwessen for an inventory management task during a summer job. The script facilitates the process of taking inventory of old desktop computers that need to be replaced.
Purpose

The main purpose of this script is to simplify the inventory process for computers, making it more efficient and less error-prone.
Prerequisites

Before running the Powershell scripts, please ensure that you have the necessary permissions. To do this:

    Open Powershell as an administrator.
    Run the script Set-ExecutionPolicy.ps1 to allow the execution of scripts on this machine.

Note: Enabling script execution can have security implications, as any script can be executed. Please be cautious and use this feature responsibly.
How to Use

    Run the script Inventory_Tool.ps1 after following the prerequisite steps.
    This script will gather and organize the computer's data.
    The script will create a folder in the same directory as Inventory_Tool.ps1 containing all the collected data.

Feel free to utilize the collected data as needed for your specific use case.
Important Note

This script is one of the first Powershell scripts developed by the creator and has been shared here in the hope that it may be useful to others. You are welcome to use it for both private and commercial purposes or modify it to suit your specific requirements.

Acknowledgment: While it's not mandatory, giving credit to the original author (Stefan Meeuwessen) would be appreciated if you find the script useful.
Development Background

For additional insights into the development of this script, you can refer to the [Reddit post](Reddit Post) made by the creator some time ago.

We hope this script proves valuable in simplifying your inventory management tasks. If you have any questions, suggestions, or improvements, feel free to contribute to the repository.
