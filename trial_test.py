import urllib.request
import json
import datetime

# Fetch API data
print("Fetching data from http://localhost:8080/api/data...")
resp = urllib.request.urlopen("http://localhost:8080/api/data", timeout=10)
data = json.loads(resp.read().decode('utf-8'))

today = data['today']
lookahead_end = data['lookahead_end']

print(f"Today: {today}, Lookahead End: {lookahead_end}")

def compute_status(pkg):
    total_stages = 0
    completed_stages = 0
    delayed_stages = []
    at_risk_stages = []
    ignored_stages = []
    
    for stage in pkg['stages']:
        if not stage['forecast'] and not stage['plan'] and not stage['actual']:
            continue
            
        is_bidder_list = 'bidder list approval' in stage['name'].lower()
        
        if stage['actual']:
            total_stages += 1
            completed_stages += 1
        elif is_bidder_list:
            ignored_stages.append(stage['name'])
        elif stage['forecast']:
            total_stages += 1
            if stage['forecast'] < today:
                delayed_stages.append((stage['name'], stage['forecast']))
            elif stage['forecast'] <= lookahead_end:
                at_risk_stages.append((stage['name'], stage['forecast']))
        elif stage['plan']:
            total_stages += 1
            if stage['plan'] < today:
                delayed_stages.append((stage['name'], stage['plan']))
            elif stage['plan'] <= lookahead_end:
                at_risk_stages.append((stage['name'], stage['plan']))
                
    if delayed_stages:
        status = 'delayed'
    elif at_risk_stages:
        status = 'atrisk'
    elif completed_stages == total_stages and total_stages > 0:
        status = 'completed'
    else:
        status = 'ontrack'
        
    return status, delayed_stages, ignored_stages

all_pkgs = []
for proj_key, proj_data in data['projects'].items():
    for pkg in proj_data['packages']:
        all_pkgs.append((proj_key, pkg))

status_counts = {'delayed': 0, 'atrisk': 0, 'completed': 0, 'ontrack': 0}
bidder_list_ignored_count = 0

print("\n--- Running Trial Verification ---")
for proj, pkg in all_pkgs:
    status, delayed, ignored = compute_status(pkg)
    status_counts[status] += 1
    if ignored:
        bidder_list_ignored_count += 1

print(f"Total Packages Analyzed: {len(all_pkgs)}")
print(f"Status Summary: {status_counts}")
print(f"Packages where Bidder List Approval (blank actual) was ignored: {bidder_list_ignored_count}")

# Print sample package verification
sample_pkg = all_pkgs[0][1]
status, delayed, ignored = compute_status(sample_pkg)
print(f"\nSample Verification -> Package: '{sample_pkg['package_name']}'")
print(f"  Status: {status.upper()}")
print(f"  Ignored stages: {ignored}")
print(f"  Delayed stages: {delayed if delayed else 'None'}")
print("\nTrial Verification Completed Successfully!")
