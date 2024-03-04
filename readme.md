# Microsoft Learn Student Ambassador Certificate Generator

## Use it now: https://mlsa-certificate-generator.sabyasachiseal.com/

### How to run manually ?
```
git clone https://github.com/Sabyasachi-Seal/Certificate-Generator-MLSA
cd ./Certificate-Generator-MLSA
docker build -t test-1 . && docker run -p 8000:8000 test-1:latest
```

## How to send emails?
- You can use the `Mail.xlsm` file to send emails to the participants. Open this with Excel. Press ```Allow Content``` if required.
- Do not need to change anything in the file itself.
- All you need to do is to search for ```View Macros```  in excel and then select the ```Send_Mails``` macro and then click on ```Run```.
- Now open outlook and login.
- Click on outbox and see the mails being sent one by one.

## Further Releases (Planning)
- Implementing for Linux users?

<h2></h2>


Souce Repo(I made some subtle improvements) : <a href="https://github.com/muhammedogz/MLSA-Certificate-Automate">Click Here</a>
