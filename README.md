# Email File Downloader

### Application for downloading attached files to email messages on POP3 mail server.

## Configuration
Use config.ini for configuration. Change mail server:
```
mailserver = pop3.vitebsk.energo.net
```
To delete messages with files after downloading, change:
```
delete_proceeded = True
```
## Run Application
```
python email-file.py <your_user> <your_password>
```
## Logs
Set logging level in `config.ini`