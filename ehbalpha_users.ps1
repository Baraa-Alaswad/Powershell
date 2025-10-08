####################################################
# This scipt is to create user id's #
####################################################

#system.windows.forms zorgt ervoor dat je .NET framework kan gebruiken voor GUI
#system.drawing idem ivm .NET, zorgt ervoor dat je met graphics, images and drawing kan werken

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#Dpi awareness for sharper Gui scaling
#This part of the script uses a here-string (@' ... '@) to define a block of C# code
Add-Type -TypeDefinition @'
using System.Runtime.InteropServices;
public class ProcessDPI {
    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool SetProcessDPIAware();      
}
'@

#Oproepen van SetProcessDPIAware method of the ProcessDPI class. 
#The result of the method call is not captured ($null = ...), 
#and it is generally used when the return value is not needed

$null = [ProcessDPI]::SetProcessDPIAware()

##############################
# Functions clicking buttons #
##############################

#Function clicking Get-Id button.
#################################

##Checking if mandatory fields are filled in.##
function Get-Id () {

    if (($Input_Box_First_Name.Text -eq "") -and ($Input_Box_Surname.Text -eq "")) {
        $Input_Box_First_Name.BackColor = "MistyRose"
        $Input_Box_Surname.BackColor = "MistyRose"
        Error-Message-Name
        }

    elseif ($Input_Box_First_Name.Text -eq "") {
        $Input_Box_First_Name.BackColor = "MistyRose"
        $Input_Box_Surname.BackColor = "White"
        Error-Message-Name
        }

    elseif ($Input_Box_Surname.Text -eq "") {
        $Input_Box_First_Name.BackColor = "White"
        $Input_Box_Surname.BackColor = "MistyRose"
        Error-Message-Name
        }

    else {
        $Input_Box_First_Name.BackColor = "White"
        $Input_Box_Surname.BackColor = "White"
        Generate-Id
        }
}

#Function clicking Create-User button.
#################################

##Checking mandatory fields are filed in.##
function Create-User () {

    if (($Input_Box_First_Name.Text -eq "") -and ($Input_Box_Surname.Text -eq "")) {
        $Input_Box_First_Name.BackColor = "MistyRose"
        $Input_Box_Surname.BackColor = "MistyRose"
        Error-Message-Name
        }

    elseif ($Input_Box_First_Name.Text -eq "") {
        $Input_Box_First_Name.BackColor = "MistyRose"
        $Input_Box_Surname.BackColor = "White"
        Error-Message-Name
        }

    elseif ($Input_Box_Surname.Text -eq "") {
        $Input_Box_First_Name.BackColor = "White"
        $Input_Box_Surname.BackColor = "MistyRose"
        Error-Message-Name
        }

##Check if section is selected.##

    elseif ($Combo_Box_1.Text -eq "Select...") {
        $Combo_Box_1.BackColor = "MistyRose"
        Error-Message-ComboBox
        }

    else {
        $Input_Box_First_Name.BackColor = "White"
        $Input_Box_Surname.BackColor = "White"
        $Combo_Box_1.BackColor = "White"
        Check-Mail
        }
}

##################
# Error messages #
##################

##Error message when name fields are left empty##
function Error-Message-Name () {

    #Message Box Main Form.
    $mainFormErrorName = New-Object System.Windows.Forms.Form
    $mainFormErrorName.Size = New-Object System.Drawing.Size(300,200)
        #.Topmost betekent naar de voorgrond
    $mainFormErrorName.TopMost = $true
    $mainFormErrorName.Text = "Error..."
    $mainFormErrorName.StartPosition = "CenterScreen"


    #Message Text Line 1.
    $errormessageTextNameLine1 = New-Object system.Windows.Forms.Label
    $errormessageTextNameLine1.Location = New-Object System.Drawing.Size(45,35)
    $errormessageTextNameLine1.AutoSize = $true
    $errormessageTextNameLine1.Text = "One or more mandatory"
    $errormessageTextNameLine1.Font = 'Microsoft Sans Serif,11'

    $mainFormErrorName.Controls.add($errormessageTextNameLine1)


    #Message Text Line 2.
    $errorMessageTextNameLine2 = New-Object system.Windows.Forms.Label
    $errorMessageTextNameLine2.Location = New-Object System.Drawing.Size(92,65)
    $errorMessageTextNameLine2.AutoSize = $true
    $errorMessageTextNameLine2.Text = "fields are empty!"
    $errorMessageTextNameLine2.Font = 'Microsoft Sans Serif,11'

    $mainFormErrorName.Controls.add($errorMessageTextNameLine2)


    #Ok Button.
    $errorOkButtonName = New-Object System.Windows.Forms.Button
    $errorOkButtonName.Location = New-Object System.Drawing.Size(102,118)
    $errorOkButtonName.Size = New-Object System.Drawing.Size(80,25)
    $errorOkButtonName.Text = "Ok"
    $errorOkButtonName.Font = 'Microsoft Sans Serif,11'
    $errorOkButtonName.Add_Click({$mainFormErrorName.Close()})

    $mainFormErrorName.Controls.add($errorOkButtonName)

    #Shows box.
    $mainFormErrorName.ShowDialog()
}

##Error message when section is not selected.##
function Error-Message-ComboBox () {

    #Message Box Main Form.
    $mainFormErrorBox = New-Object System.Windows.Forms.Form
    $mainFormErrorBox.Size = New-Object System.Drawing.Size(300,200)
    $mainFormErrorBox.TopMost = $true
    $mainFormErrorBox.Text = "Error..."
    $mainFormErrorBox.StartPosition = "CenterScreen"


    #Message Text Line 1.
    $errormessageTextBoxLine1 = New-Object system.Windows.Forms.Label
    $errormessageTextBoxLine1.Location = New-Object System.Drawing.Size(50,35)
    $errormessageTextBoxLine1.AutoSize = $true
    $errormessageTextBoxLine1.Text = "Please select a section"
    $errormessageTextBoxLine1.Font = 'Microsoft Sans Serif,11'

    $mainFormErrorBox.Controls.add($errormessageTextBoxLine1)

    #Ok Button.
    $errorOkButtonBox = New-Object System.Windows.Forms.Button
    $errorOkButtonBox.Location = New-Object System.Drawing.Size(102,118)
    $errorOkButtonBox.Size = New-Object System.Drawing.Size(80,25)
    $errorOkButtonBox.Text = "Ok"
    $errorOkButtonBox.Font = 'Microsoft Sans Serif,11'
    $errorOkButtonBox.Add_Click({$mainFormErrorBox.Close()})

    $mainFormErrorBox.Controls.add($errorOkButtonBox)

    #Shows Box.
    $mainFormErrorBox.ShowDialog()
}


##Error message when a mailbox already exists in ad.##
function Error-Message-Email () {

    #Message Box Main Form.
    $mainFormErrorEmail = New-Object System.Windows.Forms.Form
    $mainFormErrorEmail.Size = New-Object System.Drawing.Size(300,200)
    $mainFormErrorEmail.TopMost = $true
    $mainFormErrorEmail.Text = "Warning..."
    $mainFormErrorEmail.StartPosition = "CenterScreen"


    #Message Text Line 1.
    $errormessageTextEmailLine1 = New-Object system.Windows.Forms.Label
    $errormessageTextEmailLine1.Location = New-Object System.Drawing.Size(45,35)
    $errormessageTextEmailLine1.AutoSize = $true
    $errormessageTextEmailLine1.Text = "This email already"
    $errormessageTextEmailLine1.Font = 'Microsoft Sans Serif,11'

    $mainFormErrorEmail.Controls.add($errormessageTextEmailLine1)

    #Message Text Line 2.
    $errorMessageTextEmailLine2 = New-Object system.Windows.Forms.Label
    $errorMessageTextEmailLine2.Location = New-Object System.Drawing.Size(98,65)
    $errorMessageTextEmailLine2.AutoSize = $true
    $errorMessageTextEmailLine2.Text = "exists!"
    $errorMessageTextEmailLine2.Font = 'Microsoft Sans Serif,11'

    $mainFormErrorEmail.Controls.add($errorMessageTextEmailLine2)

    #Ok Button.
    $errorOkButtonMail = New-Object System.Windows.Forms.Button
    $errorOkButtonMail.Location = New-Object System.Drawing.Size(100,118)
    $errorOkButtonMail.Size = New-Object System.Drawing.Size(80,25)
    $errorOkButtonMail.Text = "Ok"
    $errorOkButtonMail.Font = 'Microsoft Sans Serif,11'
    $errorOkButtonMail.Add_Click({$mainFormErrorEmail.Close()})

    $mainFormErrorEmail.Controls.add($errorOkButtonMail)

    #Shows Box.
    $mainFormErrorEmail.ShowDialog()
}

##Error message when creating a user when mail address already exists.##
function Mail-Check-Error-Message () {

    #Message Box Main Form.
    $maiCheckError = New-Object System.Windows.Forms.Form
    $maiCheckError.Size = New-Object System.Drawing.Size(300,200)
    $maiCheckError.TopMost = $true
    $maiCheckError.Text = "Error..."
    $maiCheckError.StartPosition = "CenterScreen"


    #Message Text Line 1.
    $maiCheckErrorLine1 = New-Object system.Windows.Forms.Label
    $maiCheckErrorLine1.Location = New-Object System.Drawing.Size(48,35)
    $maiCheckErrorLine1.AutoSize = $true
    $maiCheckErrorLine1.Text = "Email already exists"
    $maiCheckErrorLine1.Font = 'Microsoft Sans Serif,11'

    $maiCheckError.Controls.add($maiCheckErrorLine1)


    #Message Text Line 2.
    $maiCheckErrorLine2 = New-Object system.Windows.Forms.Label
    $maiCheckErrorLine2.Location = New-Object System.Drawing.Size(60,65)
    $maiCheckErrorLine2.AutoSize = $true
    $maiCheckErrorLine2.Text = "in Active Directory!"
    $maiCheckErrorLine2.Font = 'Microsoft Sans Serif,11'

    $maiCheckError.Controls.add($maiCheckErrorLine2)


    #Ok Button.
    $errorOkButtonmaiCheckError = New-Object System.Windows.Forms.Button
    $errorOkButtonmaiCheckError.Location = New-Object System.Drawing.Size(102,118)
    $errorOkButtonmaiCheckError.Size = New-Object System.Drawing.Size(80,25)
    $errorOkButtonmaiCheckError.Text = "Ok"
    $errorOkButtonmaiCheckError.Font = 'Microsoft Sans Serif,11'
    $errorOkButtonmaiCheckError.Add_Click({$maiCheckError.Close()})

    $maiCheckError.Controls.add($errorOkButtonmaiCheckError)

    #Shows box.
    $maiCheckError.ShowDialog()
}


########################
# Functions commando's #
########################

##Function to generate user id and mailbox.##
#bv. "np"
function Generate-Id () {

    #Please Wait pop up.
    $mainFormWait = New-Object System.Windows.Forms.Form
    $mainFormWait.Size = New-Object System.Drawing.Size(300,200)
    $mainFormWait.TopMost = $true
    $mainFormWait.StartPosition = "CenterScreen"
    $mainFormWait.ControlBox = $False
    $mainFormWait.Visible = $True

    $waitText = New-Object System.Windows.Forms.Label
    $waitText.Location = New-Object System.Drawing.Size(89,80)
    $waitText.AutoSize = $true
    $waitText.Text = "Please wait..."
    $waitText.Font = 'Microsoft Sans Serif,11'

    $mainFormWait.Controls.Add($waitText)

    $mainFormWait.Update()

    #Put the First Name and Surname in a variable.
    $FirstNameBrut = $Input_Box_First_Name.Text
    $SurNameBrut = $Input_Box_Surname.Text

    #Remove uppercase letters.
    $FirstNameBrutNoUpper = $FirstNameBrut.ToLower()
    $SurNameBrutNoUpper = $SurNameBrut.ToLower()

    #Remove all the -, é, è, ç, ö, ô and ' in the names.
    $FirstName = $FirstNameBrutNoUpper.Replace("é","e")
    $FirstName = $FirstName.Replace("è","e")
    $FirstName = $FirstName.Replace("ç","c")
    $FirstName = $FirstName.Replace("ö","o")
    $FirstName = $FirstName.Replace("ô","o")
    $FirstName = $FirstName.Replace("'","")
    #$FirstName = $FirstName.Replace("-"," ")

    $SurName = $SurNameBrutNoUpper.Replace("é","e")
    $SurName = $SurName.Replace("è","e")
    $SurName = $SurName.Replace("ç","c")
    $SurName = $SurName.Replace("ö","o")
    $SurName = $SurName.Replace("ô","o")
    $SurName =  $SurName.Replace("'","")
    #$SurName = $SurName.Replace("-"," ")

    #Remove all the spaces for email address.
    $mailForeName = $FirstName.Replace(" ","")
    $mailSurname = $SurName.Replace(" ","")

    #Creating a start Id.
    #gaat eerste letters nemen van voornaam, familienaam
    $FullName = "$FirstName" + " " + "$Surname"
    $Spaces = ($FullName -split "( )" -match "^\s").Count
    $FullNamePart = $FullName -split " "

    for ($Counter1 = 0 ; $Counter1 -le $Spaces ; $Counter1++){
        $FirstId = $FirstId + $FullNamePart[$Counter1].Substring(0,1)
        }

    #Checking FirstId already exists.
    $IdCheck = [bool] (Get-ADUser -Filter {SamAccountName -eq $FirstId})

    #Adding letters to the FirstId and checking existance.
    $FirstNameVar = $FirstName.Replace(" ","")
    $SubstringCounterA = 1
    $InsertCounterA = 1

    $SurNameVar = $SurName.Replace(" ","")
    $SubstringCounterB = 1
    $InsertCounterB = 3

    While ($IdCheck -eq "True"){
        $AddLetterA = $FirstNameVar.Substring($SubstringCounterA,1)
        $FirstId = $FirstId.Insert($InsertCounterA,$AddLetterA)
        $SubstringCounterA++
        $InsertCounterA++
        $IdCheck = [bool] (Get-ADUser -Filter {SamAccountName -eq $FirstId})

        if ($IdCheck -eq "True"){$FirstId
        $AddLetterB = $SurNameVar.Substring($SubstringCounterB,1)
        $FirstId = $FirstId.Insert($InsertCounterB,$AddLetterB)
        $SubrstringCounterB++
        $InsertCounterB++
        $InsertCounterB++
        $IdCheck = [bool] (Get-ADUser -Filter {SamAccountName -eq $FirstId})
        }
    }

    $Output_Box_Id.Text = $FirstId

    #-erroraction stop zorgt ervoor dat wanneer er een error voorkomt bij import-module dan gaat het script niet verder
    Import-Module ActiveDirectory -ErrorAction Stop

    #################################################################
    $mainFormWait.Close()

    #Generate the email address.
    $EmailAddress = "$mailForeName" + "." + "$mailSurname" + "@ehbalpha.be"


    #Check if email already exist.
    $Seeking = Get-ADUser -Filter {mail -eq $EmailAddress}
    
    if ($Seeking -eq $null){
        $Output_Box_Email.Text = $EmailAddress
        }

    else {
        $Output_Box_Email.BackColor = "MistyRose"
        $Output_Box_Email.Text = $EmailAddress
        Error-Message-Email
        }

}


##Checking if mail does not exist before creating user in ad.##
function Check-Mail () {

    $mailToCheck = $Output_Box_Email.Text
    $mailCheck = Get-ADUser -Filter {mail -eq $mailToCheck}

    if ($mailCheck -eq $null) {
    Create-In-Ad
    }

    else {
    Mail-Check-Error-Message
    }

}

##Function creating user in ad.##
Function Create-In-Ad () {

    #Variables.
    $dateCreated = Get-Date -Format "dd/MM/yyyy - HH:mm"
    $department = $Combo_Box_1.Text
    $firstName = $Input_Box_First_Name.Text
    $surname = $Input_Box_Surname.Text
    $sAMAccountName = $Output_Box_Id.Text
    $description = "$department" + " - " + "$firstName" + " " + "$Surname"

    #Defining the distinguishedName attribute.
    #CN=Boekhouding,OU=Boekhouding,DC=ehbalpha,DC=be

    if ($department -eq "Boekhouding") {
        $distinguishedName =  "OU=Boekhouding,DC=ehbalpha,DC=be"
        }

     elseif ($department -eq "Aankoopdienst") {
        $distinguishedName = "OU=Aankoopdienst,DC=ehbalpha,DC=be"
        }
    
    elseif ($department -eq "Dienst_na_verkoop") {
        $distinguishedName = "OU=Dienst_na_verkoop,DC=ehbalpha,DC=be"
        }

    elseif ($department -eq "Directie") {
        $distinguishedName = "OU=Directie,DC=ehbalpha,DC=be"
        }

    elseif ($department -eq "Dispatching") {
        $distinguishedName = "OU=Dispatching,DC=ehbalpha,DC=be"
        }
    
    elseif ($department -eq "IT-dienst Infrastructuur") {
        $distinguishedName = "OU=IT-dienst Infrastructuur,DC=ehbalpha,DC=be"
        }
    
    elseif ($department -eq "IT-dienst Ontwikkeling") {
        $distinguishedName = "OU=IT-dienst Ontwikkeling,DC=ehbalpha,DC=be"
        }
    
    elseif ($department -eq "Kader") {
        $distinguishedName = "OU=Kader,DC=ehbalpha,DC=be"
        }
    
    elseif ($department -eq "Magazijn") {
        $distinguishedName = "OU=Magazijn,DC=ehbalpha,DC=be"
        }
    
    elseif ($department -eq "Raad_van_bestuur") {
        $distinguishedName = "OU=Raad_van_bestuur,DC=ehbalpha,DC=be"
        }
    
    elseif ($department -eq "Verkoopdienst") {
        $distinguishedName = "OU=Verkoopdienst,DC=ehbalpha,DC=be"
        }

    
    $Attributes = @{
    company = "ehbalpha"
    department = "$department"
    surname = "$surname"
    displayName = "$firstName" + " " + "$surname"
    description = "$description"
    emailaddress = $Output_Box_Email.Text
    Path = "$distinguishedName"
    givenName = $Input_Box_First_Name.Text
    name = "$sAMAccountName"
    sAMAccountName = "$sAMAccountName"
    userPrincipalName = $Output_Box_Email.Text
    Enabled = $true

   AccountPassword = "Azerty123" | ConvertTo-SecureString -AsPlainText -Force
   }

   #cmd let om nieuwe user aan te maken
   New-ADUser @Attributes
   
   
   ###AD user toevoegen aan geselecteerde groep###
   
    Add-ADGroupMember -Identity "$department" -Members "$sAMAccountName"


    ##Confirmation and resume of the program.##
    #Get the AD info and trim to string
    $resumeUserId = Get-ADUser $sAMAccountName -Properties * | Select-Object sAMAccountName | Out-String
    
    $resumeUserId = $resumeUserId.Replace("sAMAccountName","")
    $resumeUserId = $resumeUserId.Replace("-","")
    $resumeUserId = $resumeUserId.Trim()
   
    $resumeUserEmail = Get-ADUser $sAMAccountName -Properties * | Select-Object mail | Out-String
    $resumeUserEmail = $resumeUserEmail.Replace("mail","")
    $resumeUserEmail = $resumeUserEmail.Replace("-","")
    $resumeUserEmail = $resumeUserEmail.Trim()

    
    #Message Box Main Form
    $mainFormResume = New-Object System.Windows.Forms.Form
    $mainFormResume.Size = New-Object System.Drawing.Size(400,460)
    $mainFormResume.TopMost = $true
    $mainFormResume.Text = "Resume..."
    $mainFormResume.StartPosition = "CenterScreen"


    #Message Text.
    $messageText1 = New-Object system.Windows.Forms.Label
    $messageText1.Location = New-Object System.Drawing.Size(100,10)
    $messageText1.AutoSize = $true
    $messageText1.Text = "The user is created."
    $messageText1.Font = 'Microsoft Sans Serif,10'

    $mainFormResume.Controls.add($messageText1)

    #Message Text.
    $messageText2 = New-Object system.Windows.Forms.Label
    $messageText2.Location = New-Object System.Drawing.Size(80,35)
    $messageText2.AutoSize = $true
    $messageText2.Text = "Thanks for using this tool."
    $messageText2.Font = 'Microsoft Sans Serif,10'

    $mainFormResume.Controls.add($messageText2)


    $resumeOutput_Box_Id = New-Object System.Windows.Forms.RichTextBox
    $resumeOutput_Box_Id.Location = New-Object System.Drawing.Size(40,70)
    $resumeOutput_Box_Id.Size = New-Object System.Drawing.Size(300,290)
    $resumeOutput_Box_Id.Font = 'Microsoft Sans Serif,10'
    $resumeOutput_Box_Id.Text =
    "User Id:`n" + "$resumeUserId`n`n" + "Email Address:`n" + "$resumeUserEmail`n`n"

    $mainFormResume.Controls.Add($resumeOutput_Box_Id)

    #Ok Button
    $OkButton = New-Object System.Windows.Forms.Button
    $OkButton.Location = New-Object System.Drawing.Size(160,375)
    $OkButton.Size = New-Object System.Drawing.Size(80,25)
    $OkButton.Text = "Ok"
    $OkButton.Font = 'Microsoft Sans Serif,11'
    $OkButton.Add_Click({$mainFormResume.Close()})

    $mainFormResume.Controls.add($OkButton)

    $mainFormResume.ShowDialog()

    
    ##clears the form after closing confirmation.##
    $Input_Box_First_Name.Clear()
    $Input_Box_First_Name.BackColor = "White"

    $Input_Box_Surname.Clear()
    $Input_Box_Surname.BackColor = "White"

    $Output_Box_Id.Clear()

    $Output_Box_Email.Clear()
    $Output_Box_Email.BackColor ="White"
    
    #Reset the combo boxes.

    $Combo_Box_1.SelectedIndex = 0
    $Combo_Box_1.BackColor = "White"


}


###########################
# Main Graphic Interface. #
###########################

#Graphic Interface Body.
$Main_Form = New-Object System.Windows.Forms.Form
$Main_Form.Size = New-Object System.Drawing.Size(560,535)
$Main_Form.Text = "User aanmaken"

#First Name text.
$First_Name_Text = New-Object System.Windows.Forms.Label
$First_Name_Text.Location = New-Object System.Drawing.Size(35,30)
$First_Name_Text.AutoSize = $true
$First_Name_Text.Text = "First Name:"
$First_Name_Text.Font = 'Microsoft Sans Serif,11'

$Main_Form.Controls.Add($First_Name_Text)

#Inputbox for First Name.
$Input_Box_First_Name = New-Object System.Windows.Forms.RichTextbox
$Input_Box_First_Name.Location = New-Object System.Drawing.Size(145,30)
$Input_Box_First_Name.Size = New-Object System.Drawing.Size(350,25)
$Input_Box_First_Name.Font = 'Microsoft Sans Serif,11'
$Input_Box_First_Name.Multiline = $false

$Main_Form.Controls.Add($Input_Box_First_Name)

#* First Name.
$First_Name_Text_Ster = New-Object System.Windows.Forms.Label
$First_Name_Text_Ster.Location = New-Object System.Drawing.Size(500,30)
$First_Name_Text_Ster.AutoSize = $true
$First_Name_Text_Ster.Text = "*"
$First_Name_Text_Ster.Font = 'Microsoft Sans Serif,14'

$Main_Form.Controls.Add($First_Name_Text_Ster)

#Surname text.
$Surname_Text = New-Object System.Windows.Forms.Label
$Surname_Text.Location = New-Object System.Drawing.Size(35,80)
$Surname_Text.AutoSize = $true
$Surname_Text.Text = "Surname:"
$Surname_Text.Font = 'Microsoft Sans Serif,11'

$Main_Form.Controls.Add($Surname_Text)

#Inputbox for Surname.
$Input_Box_Surname = New-Object System.Windows.Forms.RichTextbox
$Input_Box_Surname.Location = New-Object System.Drawing.Size(145,80)
$Input_Box_Surname.Size = New-Object System.Drawing.Size(350,25)
$Input_Box_Surname.Font = 'Microsoft Sans Serif,11'
$Input_Box_Surname.Multiline = $false

$Main_Form.Controls.Add($Input_Box_Surname)

#* Surname.
$Surname_Text_Ster = New-Object System.Windows.Forms.Label
$Surname_Text_Ster.Location = New-Object System.Drawing.Size(500,80)
$Surname_Text_Ster.AutoSize = $true
$Surname_Text_Ster.Text = "*"
$Surname_Text_Ster.Font = 'Microsoft Sans Serif,14'

$Main_Form.Controls.Add($Surname_Text_Ster)

#Get Id button.
$Get_Id_Button = New-Object System.Windows.Forms.Button
$Get_Id_Button.Location = New-Object System.Drawing.Size(28,123)
$Get_Id_Button.Size = New-Object System.Drawing.Size(100,40)
$Get_Id_Button.Text = "Get Id"
$Get_Id_Button.Font = 'Microsoft Sans Serif,11'
$Get_Id_Button.Add_Click({Get-Id})

$Main_Form.Controls.add($Get_Id_Button)

#Output box for new user id.
$Output_Box_Id = New-Object System.Windows.Forms.RichTextBox
$Output_Box_Id.Location = New-Object System.Drawing.Size(145,130)
$Output_Box_Id.Size = New-Object System.Drawing.Size(350,25)
$Output_Box_Id.Font = 'Microsoft Sans Serif,11'
$Output_Box_Id.Multiline = $false

$Main_Form.Controls.Add($Output_Box_Id)

#Email text.
$Email_Text = New-Object System.Windows.Forms.Label
$Email_Text.Location = New-Object System.Drawing.Size(35,180)
$Email_Text.AutoSize = $true
$Email_Text.Text = "Email:"
$Email_Text.Font = 'Microsoft Sans Serif,11'

$Main_Form.Controls.Add($Email_Text)

#Output box Email.
$Output_Box_Email = New-Object System.Windows.Forms.RichTextBox
$Output_Box_Email.Location = New-Object System.Drawing.Size(145,180)
$Output_Box_Email.Size = New-Object System.Drawing.Size(350,25)
$Output_Box_Email.Font = 'Microsoft Sans Serif,11'
$Output_Box_Email.Multiline = $false

$Main_Form.Controls.Add($Output_Box_Email)

#Section text.
$Section_Text = New-Object System.Windows.Forms.Label
$Section_Text.Location = New-Object System.Drawing.Size(35,230)
$Section_Text.AutoSize = $true
$Section_Text.Text = "Section:"
$Section_Text.Font = 'Microsoft Sans Serif,11'

$Main_Form.Controls.Add($Section_Text)

#Drop Down Menu Section.
$descriptions = @("Select...","Aankoopdienst","Boekhouding","Dienst_na_verkoop","Directie","Dispatching","IT-dienst Infrastructuur","IT-dienst Ontwikkeling","Kader","Magazijn","Raad_van_bestuur","Verkoopdienst")

#Drop down menu section.
$Combo_Box_1 = New-Object System.Windows.Forms.ComboBox
$Combo_Box_1.Location = New-Object System.Drawing.Point(145,230)
$Combo_Box_1.Size = New-Object System.Drawing.Size(150, 15)
$Combo_Box_1.Font = 'Microsoft Sans Serif,11'
$Combo_Box_1.DataSource = $descriptions
$Combo_Box_1.add_SelectedIndexChanged($Combo_Box_1_SelectedIndexChanged)

$Main_Form.Controls.Add($Combo_Box_1)

#* Section 1.
$Section1_Ster = New-Object System.Windows.Forms.Label
$Section1_Ster.Location = New-Object System.Drawing.Size(301,230)
$Section1_Ster.AutoSize = $true
$Section1_Ster.Text = '*'
$Section1_Ster.Font = 'Microsoft Sans Serif,14'

$Main_Form.Controls.Add($Section1_Ster)

#AD Groups text.
$Ad_Groups_Text = New-Object System.Windows.Forms.Label
$Ad_Groups_Text.Location = New-Object System.Drawing.Size(35,280)
$Ad_Groups_Text.AutoSize = $true
$Ad_Groups_Text.Text = 'The user will be added to the corresponding OU and group in AD.'
$Ad_Groups_Text.Font = 'Microsoft Sans Serif,12'

$Main_Form.Controls.Add($Ad_Groups_Text)

#Create Button.
$createUserButton = New-Object System.Windows.Forms.Button
$createUserButton.Location = New-Object System.Drawing.Size(215,380)
$createUserButton.Size = New-Object System.Drawing.Size(110,50)
$createUserButton.Text = 'Create User'
$createUserButton.Font = 'Microsoft Sans Serif,11'
$createUserButton.Add_Click({Create-User})

$Main_Form.Controls.Add($createUserButton)

#Shows Form
$Main_Form.ShowDialog()