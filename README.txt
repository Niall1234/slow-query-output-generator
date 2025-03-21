====================SETUP GUIDE====================

1. create a new folder to run this script in and put the query_outputs.py file and requirements.txt file in this folder
2. create a new file in this folder and call it '.env'
3. open .env and type out the following
    USERNAME=<your RISK ID>
    PASSWORD=<your EU Password>
4. save your .env file

SETTING UP THE ENVIRONMENT - MANUAL METHOD
1. open a terminal window in this folder and run the commands below
    python -m venv 11_env
    11_env\Scripts\activate 
    python -m pip install -r requirements.txt 
2. run python query_output.py 

SETTING UP THE ENVIRONMENT - VSCODE METHOD
1. open the folder in VSCode
2. Press CTRL+SHIFT+P > 'Python: Create Environment...' > 'Venv' > 'requirements.txt'