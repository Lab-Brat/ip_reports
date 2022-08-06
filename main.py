from tools import Tools

# necessary documents
ip_adrs = 'relevant_list.csv'
prev_rep = "prp-0.xlsx"
curr_rep = "prp-current.xlsx"
legend   = "legend.csv"

# optional documets
clf_ips = "cloudflare_subnets.txt"

tt = Tools(ip_adrs, prev_rep, curr_rep, clf_ips, legend)
tt.color_xlsx()
# tt.test()
