# Extracting tokens from the OutLook Process

* There are majorly two research posts talking about extraction of JWT tokens from Outlook APP
* First is [https://mrd0x.com/stealing-tokens-from-office-applications/](https://mrd0x.com/stealing-tokens-from-office-applications/)
* Second is [https://blog.xpnsec.com/wam-bam/](https://blog.xpnsec.com/wam-bam/)


## Extraction from memory 
* The first research post talks about extracting tokens from the outlook process memory, and that is what exactly we would also be doing
* we will combining a custom C sharp memory parser with legitimate tool "ProcessDump" from Microsoft
* Official Binary of ProcessDump can be downloaded from Microsoft site at: (https://learn.microsoft.com/en-us/sysinternals/downloads/procdump)[https://learn.microsoft.com/en-us/sysinternals/downloads/procdump]
  * Once downloaded, make sure that when using the C# parser binary, the process dump binary is present in the same directory
  * In future release the C# parser binary will be extended to include features such as remote download of process dump

* Clone the source code
```bash
git clone https://github.com/vrikodar/TokenLook
cd TokenLook/TokenExtract/source

ls extract_noexfil.cs
# This is the source code for C#, which will create the Outlook process dump using ProcessDump and then parse it 
```

### `extract_noexfil.cs`
* As I am **no expert in C# yet**, the [source code](https://raw.githubusercontent.com/vrikodar/TokenLook/refs/heads/main/TokenExtract/source/extract_noexfil.cs) in this file is a bit AI tweaked/enhanced.
* The program on a high level does following:
  * Looks for the `ProcessDump.exe` in the local directory
  * Once `ProcessDump.exe` binary is found, the program will search for the Outlook app's process ID
  * Once the Outlook app is found to be running on the machine, the Program will use `ProcessDump.exe` to create a full memory dump of the Outlook process
  * The memory dump file is saved to the APP DATA directory of the user under who's context the binary is run 
    * **Note that in some cases "depending on the environment" this memory dump can be of large size >1Gb**
  * Once the memory dump is created, the program will parse the DMP file chunk by chunk and extract all strings from it.
  * All the extracted strings are then saved to a TXT file in the same directory.

### Retrieving JWT tokens from the TXT file \(Manually\)
* Manually look for the strings starting with `"eyJ0"` inside the TXT file
* As tokens are found, we need to [decode](https://jwt.ms/) them to find a token, which on decoding has the `aud` feild set to `https://outlook.office.com` or `https://outlook.office365.com`


### Combining ProcessDump and `Custom C# memory parser`


## Extraction from TBRES files
* The second research post talks about extracting the same tokens from TBRES files, which are DPAPI encrypted
* The researcher has created a tool which uses DPAPI to extract the tokens from TBRES files.
* TBRES file store the tokens as DPAPI encrypted material inside XML files.
* It is however worth noting that use of DPAPI for actions such as TBRES file decryption can easily be flagged by EDRs
  * As of now we won't be touching this method of extraction, possibly in futre updates
 
# **For More Detailed Documentation Refer to the** [WIKI](https://github.com/vrikodar/TokenLook/wiki/TokenLook-%E2%80%90-TokenExtract)
