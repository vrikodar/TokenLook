# Using the extarcted JWT tokens

## Command Line version 
* The command line interface of tool runs in a specific format.
* The current directory `TokenUse` contains source code files to use the JWT tokens

### Setup 
* It is recommended that you use the tool from within a VENV environment

  ```bash
  git clone https://github.com/vrikodar/TokenLook
  cd TokenLook/TokenUse/source

  # configure and activate venv
  python3 -m venv local_venv
  source local_venv/bin/activate

  # run the main script to initiate execution
  python3 main.py
  ```
### Configuring the `tokenlook_config.json`
* This is the main configuration file referenced by the command line interface everytime on execution.
* By default the tool will look for this configuration file in the same directory, from which the main.py file is being run.

* The configuration file "tokenlook_config.json" structure is as follows "below we have a example configuration"

```js
{
  "current_jwt": ["JWT-VALUE1", "JWT-VALUE2"],
  "data_dir": "local_data",
  "proxy": ["http://127.0.0.1:8080", "PROXY2", "PROX3"],
  "mail_numb": 10,
  "keyword_search": ["pass", "username", "credentials", "key"]
}
```
* Basic options from the configuration file
  * `current_jwt`: list containing JWT tokens "could be for different users"
  * `data_dir`: directory path to which the tool will save everything on disk, includes emails in JSON and attachments
  * `proxy`: proxy to route traffic through, useful for debugging "Burp proxy etc.", leave this list empty if you don't plan on using a proxy
  * `mail_numb`: The maximum number of emails tool will extract, 10 is the minimum value, specify more if required.
  * `keyword_search`: List of keywords the tool will look for in the extracted emails, useful in cases such as looking for clear text credentials in emails.

# For more detailed documentation related to `TokenUse` refer to the WIKI
[TokenLook ‐ TokenUse](https://github.com/vrikodar/TokenLook/wiki/TokenLook-%E2%80%90-TokenUse)
