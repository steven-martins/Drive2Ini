Drive2Ini
=========

Synchronize a Google Spreadsheet with an ini File



# Create a virtualenv

```bash
virtualenv --python=python2 env
source env/bin/activate
pip install -r requirements.txt
```

# Configure it
```bash
cp local_config.py.sample local_config.py
```

## Example:
```python
config = {
    "username": 'login_x@gmail.com',
    "password": "your_password",
    "worksheet_key": "key_available_from_the_url",
    "sheet_pos": 0 # number of the sheet/tab. From 0 to ...
}

```

# To use it
```
USAGE: Drive2Ini.py toDrive|toFile file.ini
```

