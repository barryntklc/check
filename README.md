# check

Dumps the path, size, and checksum of files into an Excel spreadsheet (.xlsx).

By default, scans the current folder and dumps information to a spreadsheet in the same folder.

USAGE:
    check.win.exe /s [source directory path] | [ /? | 
                                                /v | 
                                                /verbose | 
                                                /log [excel file path] ]

Options:
    /?                          Display this help message
    /v  /verbose                Show extra information, for diagnostic purposes
    /s [source directory path]  Check all files from the [source directory path]
    /log [excel file path]      Save the file to the specified [excel file path]