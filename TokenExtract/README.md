# Extracting tokens from the OutLook Process

* There are majorly two research posts talking about extraction of JWT tokens from Outlook
* First is [https://mrd0x.com/stealing-tokens-from-office-applications/](https://mrd0x.com/stealing-tokens-from-office-applications/)
* Second is [https://blog.xpnsec.com/wam-bam/](https://blog.xpnsec.com/wam-bam/)


## Extraction from memory 
* The first research post above talks about extracting tokens from the outlook process memory, and that is what exactly we would also be doing
* we will combining a custom C sharp memory parser with legitimate tool "ProcessDump" from Microsoft

### Combining ProcessDump and `Custom C# memory parser`


## Extraction from TBRES files
* The second research post talks about extracting the same tokens from TBRES files, which are DPAPI encrypted
* The researcher has created a tool which uses DPAPI to extract the tokens from TBRES files.
* TBRES file store the tokens as DPAPI encrypted material inside XML files.
* It is however worth noting that use of DPAPI for actions such as TBRES file decryption can easily be flagged by EDRs
  * As of now we won't be touching this method of extraction, possibly in futre updates
 
