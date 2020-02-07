
# PSLongPath
Powershell Script which reveals long paths on a drive

# How can it be so fast?
On my tests it took less than 20 seconds to scan my C: and read 500k files. The technology behind this is the same that is being used by the likes of TreeSize, WizTree etc. Instead of scanning the whole disk and creating threads upon threads to get the results, my script with the help of PowerForensicsV2 reads the MFT of the hard drive. 
>[https://github.com/Invoke-IR/PowerForensics](https://github.com/Invoke-IR/PowerForensics)

# How to use
Download the script (.ps1) and open a command prompt on the folder where you've downloaded it. 
Type in 

> powershell.exe -executionPolicy Bypass -File PSLongPath.ps1

A form will appear and the rest is self explanatory. 
Select your drive, set the least path length from which you want to create your report and click Scan Drive.
After the initial scan has finished you will be able to Export to CSV and to HTML.

The Export to HTML is using the PSWriteHTML from EvotecIT
> [https://github.com/EvotecIT/PSWriteHTML](https://github.com/EvotecIT/PSWriteHTML)

From the HTML file you can then export to Excel, CSV if you want, PDF, filter it and more. Please note that if you have a lot of files returned, the HTML file might become large. On my tests for 30k files, it rendered a 7MB file, which Chrome was struggling to open on a Ryzen 7 3700X, 16GB RAM with NVMe SSD.

# Is it stable?
Currently, I have only been able to test this on my personal computer and it won't work with network drives which you don't have physically attached to the computer. I guess VMDKs and VHDs on VMs should work fine. If you test it and it works on your environment, please drop me a message or create a new issue and give me the details, so I can update this section.

# Where was the icon taken from?
See below:
> <a target="_blank" href="https://icons8.com/icons/set/happy-document">Winking Document icon</a> icon by <a target="_blank" href="https://icons8.com">Icons8</a>

# Tips

 - If you click on the first column, it will open the parent folder of the file.
 - Keyboard Shortcuts

| Keys | Function |
|--|--|
| Ctrl + X | Close Window |
| Ctrl + S | Export to C**S**V |
| Ctrl + H| Export to **H**TML |
| Ctrl + I | Open About Form |

# Screenshots
Results (Only 14s)

![Results](https://imgur.com/CksJ49b.png)

File

![File](https://imgur.com/XffELPq.png)

About

![About](https://imgur.com/7bxmz69.png)
