# Microsoft Learn Student Ambassador Certificate Automation

### Event done ? Ready to send certificates ? 

### Click here to access certificate generator: https://mlsa.sabyasachiseal.com/

Backup website: https://mlsa-2.sabyasachiseal.com/ (Takes time to load)

### Tutorial

1. Enter your name and event name, and select a CSV file (probably from registration form). Make sure the csv file you 
   upload has a "Name" and "Email" column.
   ![image](https://github.com/user-attachments/assets/e53cf703-a1c4-4968-a238-2b34d494a89c)
2. Click on Generate Certificates, it might take some time (it might feel like its stuck at 100%) but it will generate 
   all the certificates.
   ![image](https://github.com/user-attachments/assets/430e426c-2436-42f5-82db-7f33620bf9fb) ![image](https://github.com/user-attachments/assets/add03930-718f-4d83-8ed1-708d53cf229a)
3. You can now click on "Send emails" and send all the certificates to the participants. Remember, all the emails might 
   not reflect on the participants mails instantly, it might take upto 6 hours for a send email to reflect on the 
   participant's mailbox.
   ![image](https://github.com/user-attachments/assets/1fcefcbf-4cc0-4e4a-8bee-89a130473b10)


Example of what the participants receive on their mail:
![image](https://github.com/user-attachments/assets/8598c318-4737-4197-b8b6-e542ae95588f)

And the certificates are attached as such:
![image](https://github.com/user-attachments/assets/20f151cf-8267-4de4-a128-f5356d706ece)


---

### Other event related details, links and rules: https://stdntpartners.sharepoint.com/sites/SAProgramHandbook/SitePages/Hosting-Events.aspx 


---
<h1>Instructions Below are for you to run this project locally. Use <code>master</code> branch for the same </h1>


### Tutorial Video for running locally:

[![IMAGE ALT TEXT](http://img.youtube.com/vi/OUAbqdLDTZQ/0.jpg)](http://www.youtube.com/watch?v=OUAbqdLDTZQ "How to use Certificate Generator MLSA")

This repo simply use a template certificate docx file and generates certificates
both docx and pdf.

###  Automatic Mail Working on Windows 10 only.

## Prerequisite

We want give you some advices or tips before to run or execute the app locally, in order to follow python best practice 
given by the community.

Instead, to install all the package required to run the app, locally we suggest you to use a virtual environment that 
gives some advantage, like:
- Not create conflicts with previous library or package installed locally
- Resolve version library conflicts that can happen if a same library is installed with different version
- And all package installed in the venv(Virtual environment) are just for this project not other one so that can affect 
  the behavior of another app

So in order to install or create a virtual environment, you to follow this steps :

- Step 1
  - Open the terminal(Cli: Command Line Interface) in order to execute some command there
  - Verify if you are in the right directory before to execute any command, the end of path showed in the terminal, must
    be the name of the project
- Step 2 : Create virtual environment
  - Run this command to create a virtual environment :
    If you haven't configured the `PATH` and `PATHEXT` variables for your <a href="https://docs.python.org/3/using/windows.html#using-on-windows">Python installation</a>, use this command :
    ```
    c:\>Python35\python -m venv c:\path\to\myenv
    ```
    else
    ```
    c:\>python -m venv c:\path\to\myenv
    ```
  - Replace the `c:\path\to\myenv` by the correct path of your virtual environment.
  - It's recommended to place the virtual environment at the parent directory of your project.
- Step 3 : Activate the virtual environment
  - Activation of a virtual environment depends on the platform you use.
    - if you are using a posix based platform :
      - bash/zsh : `$ source <venv>/bin/activate`
      - fish : `$ source <venv>/bin/activate.fish`
      - csh/tcsh : `$ source <venv>/bin/activate.csh`
      - PowerShell : `$ <venv>/bin/Activate.ps1`
    - If is Windows system :
      - cmd.exe : `C:\> <venv>\Scripts\activate.bat`
      - PowerShell : `PS C:\> <venv>\Scripts\Activate.ps1`
  - The `<venv>` correspond to the entire path of virtual environment if you are in the root directory project instead 
    you can the relative path
  
Be sure to activate the virtual environment before to run other command to install the required packages.

After you have created correctly created the virtual environment be sure to add that in the .gitignore file in order to 
not version that. If you want to learn more about virtual environment or other good practices in python you can visit the official website of the python community <a href="https://docs.python.org/3/library/venv.html">Click Here</a>

## Run these commands in your terminal

```
git clone https://github.com/Sabyasachi-Seal/Certificate-Generator-MLSA
cd Certificate-Generator-MLSA
```
Now Copy your Participant List to the Data Folder and rename it as `ParticipantList.csv`. <br>
<e><i>The list must have the following fields: ```Name, Email```</i></e>. It may have more, but these 2 are essential.
```
pip install -r requirements.txt
python main_certificate.py
```

## *Important*
### Do not use this button to run your code:
![Screenshot from 2023-08-26 10-36-02](https://github.com/Sabyasachi-Seal/Certificate-Generator-MLSA/assets/36451386/6e4ddf15-c97a-4c1a-9e64-9cd4db416511)

### Use ```python main_certificate.py``` to run your code


## Customization
- You can change the certificate template file in the `Data` folder.
- You can change the email template file in the `Data` folder.

## How to send emails?
- You can use the `Mail.xlsm` file to send emails to the participants. Open this with Excel. Press ```Allow Content``` 
  if required.
- Do not need to change anything in the file itself.
- All you need to do is to search for ```View Macros```  in Excel and then select the ```Send_Mails``` macro and then 
  click on ```Run```.
- Now open outlook and login.
- Click on outbox and see the mails being sent one by one.

<h2></h2>


Source Repo(I made some subtle improvements) : <a href="https://github.com/muhammedogz/MLSA-Certificate-Automate">Click Here</a>
