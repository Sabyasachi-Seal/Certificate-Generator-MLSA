# Microsoft Learn Student Ambassador Certificate Automation

### Event done ? Ready to send certificates ? 

### Click here to access certificate generator: https://mlsa-certificate-generator.sabyasachiseal.com/


### Other event related details, links and rules: https://stdntpartners.sharepoint.com/sites/SAProgramHandbook/SitePages/Hosting-Events.aspx 


---
<h1>Instructions Below are for you to run this project locally. Use <code>master</code> branch for the same </h1>


### Tutorial Video for running locally:

[![IMAGE ALT TEXT](http://img.youtube.com/vi/OUAbqdLDTZQ/0.jpg)](http://www.youtube.com/watch?v=OUAbqdLDTZQ "How to use Certificate Generator MLSA")

This repo simply use a template certificate docx file and generates certificates
both docx and pdf.

###  Automatic Mail Working on Windows 10 only.

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
- You can use the `Mail.xlsm` file to send emails to the participants. Open this with Excel. Press ```Allow Content``` if required.
- Do not need to change anything in the file itself.
- All you need to do is to search for ```View Macros```  in excel and then select the ```Send_Mails``` macro and then click on ```Run```.
- Now open outlook and login.
- Click on outbox and see the mails being sent one by one.

<h2></h2>


Souce Repo(I made some subtle improvements) : <a href="https://github.com/muhammedogz/MLSA-Certificate-Automate">Click Here</a>
