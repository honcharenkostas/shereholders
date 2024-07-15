# shereholders
#### Requirements
python 3.11 https://www.python.org/downloads/release/python-3116/

#### How to run
Add column "Processed" to main.xls. All records except Processed=yes will be processed via the script.
After processing the script will mark processed records as Processed=yes.
```bash
cd ~/code/shereholders ; source venv/bin/activate ; python main.py
```