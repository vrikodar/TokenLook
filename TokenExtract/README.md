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


### Combining ProcessDump and `Custom C# memory parser`


## Extraction from TBRES files
* The second research post talks about extracting the same tokens from TBRES files, which are DPAPI encrypted
* The researcher has created a tool which uses DPAPI to extract the tokens from TBRES files.
* TBRES file store the tokens as DPAPI encrypted material inside XML files.
* It is however worth noting that use of DPAPI for actions such as TBRES file decryption can easily be flagged by EDRs
  * As of now we won't be touching this method of extraction, possibly in futre updates
 
