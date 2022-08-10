from tools import Reader, Tools

# document full paths
directory = '/opt/tables/'
ip_adrs   = f'{directory}ip_addresses.csv'
prev_rep  = f'{directory}prp-0.xlsx'
curr_rep  = f'{directory}prp-1.xlsx'
clf_ips   = f'{directory}subnets.txt'
legend    = f'{directory}legend.csv'

rd = Reader(ip_adrs, prev_rep, curr_rep, clf_ips, legend)
Tools(rd).color_xlsx()
