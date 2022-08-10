from tools import Reader, Tools

# necessary documents
ip_adrs = 'ip_addresses.csv'
prev_rep = "prp-0.xlsx"
curr_rep = "prp-1.xlsx"
legend   = "legend.csv"

# optional documets
clf_ips = "subnets.txt"

rd = Reader(ip_adrs, prev_rep, curr_rep, clf_ips, legend)
Tools(rd).color_xlsx()
