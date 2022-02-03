# Microsoft Learn Student Ambassador Certificate Automation

This repo simply use a template certificate docx file and generates certificates
both docx and pdf.

## Setup

- Working on Windows only.
- Required Microsoft Doc

```
git clone https://github.com/muhammedogz/MLSA-Certificate-Automate.git
cd MLSA-Certificate-Automate
pip install -r requirements.txt
python main_certificate.py
```

## Customization

In, `main_certificate.py` folder. Change your participants and your name with path and name.
In case, you have to deal with your own participate file due to your needs, you should update `list_participants` data.

Also, functions are very simple. You can implement your own versions as well.

## Further Releases ( Not Sure :) )

- Improving coding style
- Adding e-mail automation to script?
- Implementing for Linux users?
