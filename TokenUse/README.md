# Using the extarcted JWT tokens

## Command Line version 
* The command line interface of tool runs in a specific format.
* The current directory `TokenUse` contains source code files to use the JWT tokens

### Setup 
* It is recommended that you use the tool within a VENV environment

  ```bash
  git clone https://github.com/vrikodar/TokenLook
  cd TokenLook/TokenUse

  # configure and activate venv
  python3 -m venv local_venv
  source local_venv/bin/activate

  # run the main script to initiate execution
  python3 main.py
  ```
### Configuring the `tokenlook_config.json`
* This is the main configuration file referenced by the command line interface everytime on execution.
* By default the tool will look for this configuration file in the same directory, from which the main.py file is being run.
