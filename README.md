# win-inventory-tools
Windows scripts to generate reports from Windows (remote or local) workstations. It is intended to aid in the inventory of physical IT related equipment. It is strongly recommended using the Power Shell ones below instead of the VBS, as they're the most *up to date*.

## VBS scripts

`pInven.vbs`: Generates a CSV file with *custom* columns (some are purposely blank; output file will be `hostname.csv`) with the following information: User, Part Number, Serial Number, Manufacturer, Model, RAM Memory, Processor and Operating System.

`informe.vbs`: Generates a HTML-formatted file with a table showing all the running processes & PIDs in the system. 

## Power Shell scripts
*This* are the ones you should use

`pInven.ps1`: Behavior similar to the previous script, only with different columns and adapted to the current times. 

