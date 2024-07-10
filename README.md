# shereholders
#### Requirements
python 3.11 https://www.python.org/downloads/release/python-3116/

#### Init
```bash
source venv/bin/activate
```


#### Troubleshooting
1. Google Drive token has been expired or revoked
    ```bash
    google.auth.exceptions.RefreshError: ('invalid_grant: Token has been expired or revoked.', {'error': 'invalid_grant', 'error_description': 'Token has been expired or revoked.'})
    ```
    --> remove token.pickle and then return the script and follow the instruction for re-auth to google drive