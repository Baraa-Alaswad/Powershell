User Creation Tool (PowerShell)
Overview

This project is a PowerShell GUI tool used to create new users in Active Directory.
It helps IT administrators generate user IDs, create email addresses, and automatically add users to the correct Organizational Unit (OU) and security group.

The tool is made to simplify the user creation process and reduce manual errors.

What This Project Does

Generates a unique user ID (sAMAccountName) based on first name and surname

Generates an email address automatically

Checks if the user ID or email already exists in Active Directory

Creates the user in the correct OU based on the selected department

Adds the user to the corresponding Active Directory group

Shows a graphical interface (GUI) for easy use

Technologies Used

PowerShell

.NET Windows Forms (System.Windows.Forms, System.Drawing)

Active Directory PowerShell module

How It Works

Enter the user's First Name and Surname

Click Get Id to generate a username and email

Select the Department/Section

Click Create User

The script creates the user in Active Directory and shows a confirmation window

Requirements

Windows with PowerShell

Active Directory module installed

Permission to create users in Active Directory

Run the script as Administrator

Notes

Default password is hardcoded in the script (should be changed in production)

Email domain is currently set to @ehbalpha.be

OU and group names must match your Active Directory structure

Author

This project was created as an automation tool for system and network administration.