##How to use##

1. download Excel file and Python Script

2. Change credential parameter
  host = "x.x.x.x"
  username = "admin"
  password = "cisco123"

3. Change Tenant name and Application Profile name in **def get_epg(token)** function
  url = "https://" + host + '/api/node/mo/uni/tn-**Tenant_ATX**/ap-**AP_ATX**.json?query-target=subtree&target-subtree-class=fvAEPg'


4. Place Excel file in the same directory with python script.

5. Excecute Python and Excel file will be updated
