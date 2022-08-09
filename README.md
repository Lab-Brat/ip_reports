## ip_reports

Compare xslx sheets and color code the result.  
Documents required for the program to work:
* prp-0.xlsx - previous list of IP address + ports (color coded)
* prp-1.xlsx - current list of IP address + ports
* ip_addresses.csv - list of addresses that are in use
* subnets.txt - list of subnets
* legend.csv - legend that examplains the color meaning

#### How it works
Program reads and saves IPs, ports and cell colors in the format ```{ip: [(port, color), (port, color)...]...}``` of the previous report.  
Then, it proceeds to read the current report, and checks every line for the following:
* If it belongs to any of the subnets;
* If it exists in ```ip_addresses.csv```;
* If it was present in the previous report;
After each step cells are color with an appropriate color.  
The end result is ```prp-1.xlsx``` but with each line colored for better visualization.
