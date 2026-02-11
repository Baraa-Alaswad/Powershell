# Active Directory User Creation Tool

## Description
This project is a PowerShell GUI tool that automates user creation in Active Directory.  
It generates usernames, creates email addresses, and assigns users to the correct Organizational Unit (OU) and security groups.

## Features
- Generate unique sAMAccountName (User ID)
- Automatically generate email addresses
- Check if username and email already exist in Active Directory
- Create users in the correct OU based on department
- Add users to the corresponding AD security group
- Windows Forms graphical user interface (GUI)

## Technologies
- PowerShell
- .NET Windows Forms
- Active Directory PowerShell Module

## Installation
1. Download the project files
2. Open PowerShell as Administrator
3. Ensure the Active Directory module is installed
4. Run the script

## Usage
1. Enter First Name and Surname
2. Click "Get Id" to generate username and email
3. Select the Department
4. Click "Create User"
5. The user is created in Active Directory

## Configuration
- Default email domain: `@ehbalpha.be`
- OU names must match your Active Directory structure
- Default password is hardcoded and should be changed

## Requirements
- Windows OS
- Active Directory environment
- Admin permissions to create users
- PowerShell 5.1 or higher

## Notes
- Test the script in a lab environment before production use
- Improve security by changing the default password handling
- Logging can be added for enterprise usage

## Author
System and Network Engineering Student

