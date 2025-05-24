import requests
import pandas as pd

def find_invitation(full_name, url, searchLength):
    params = {"full_name": full_name}

    headers = {
        "Accept": "application/json, text/plain, */*",
        "Origin": "https://www.theknot.com",
        "Referer": "https://www.theknot.com/",
        "User-Agent": "Mozilla/5.0 (X11; Linux x86_64; rv:137.0) Gecko/20100101 Firefox/137.0",
    }

    response = requests.get(url, headers=headers, params=params)

    results = []

    if response.status_code == 200:
        data = response.json()

        matches = data.get('partialMatches') or []
        exact_match = data.get('exactMatch')
        
        # Correct handling for exactMatch being None
        if exact_match:
            matches.append(exact_match)

        for match in matches:
            envelope_label = match.get("envelopeLabel", "No Label")
            for person in match["people"]:
                first_name = person.get("firstName", "")
                last_name = person.get("lastName", "")
                email = person.get("email", "")
                invitations = person.get("invitations", [])
                for invitation in invitations:
                    rsvp_status = invitation.get("rsvp")
                    results.append({
                        "Envelope Label": envelope_label,
                        "First Name": first_name,
                        "Last Name": last_name,
                        "Email": email,
                        "RSVP Status": rsvp_status,
                    })
    else:
        print(f"Request failed for '{full_name}' with status code: {response.status_code}")
        if (searchLength < 50): # if i'm searching my entire contacts I don't want it to be putting it in the excel sheet
            results.append({
                "Envelope Label": None,
                "First Name": full_name,
                "Last Name": None,
                "Email": None,
                "RSVP Status": "Not Invited",
            })

    return results

# List of guest names to query, change this based on the wedding
guest_names = []
#guest_names = ["Anna Conforti", "Avery Millspaugh", "China Tinnen", "Dorothy Smith", "Eli Granberry", "Ellie Bunnell", "Gabby Woodie", "Josh Davis", "Kate Noel", "Kenzie Murray", "McKenzie Strasko", "Morgan Elarton", "Riley Barber", "Sam Vinson", "Jordan Griffin", "Brian Richards"]

#or use guestnames from a CSV using one's personal contacts
df = pd.read_excel("SamVinsoniOSContacts05242025.xlsx")
guest_names = df["FullName"].dropna().tolist()


#change the api url depending on the wedding 
url = "https://api.guests.xogrp.com/v1/weddings/f030edc2-6a49-46d1-aacb-9f96563fd8f4/guests"

# Accumulate results
all_results = []
for name in guest_names:
    all_results.extend(find_invitation(name, url, len(guest_names)))

# Save to Excel using pandas
out_df = pd.DataFrame(all_results)

excel_file_name="lists/GoodOleBoy_wedding_list.xlsx"
# Write DataFrame to Excel file
out_df.to_excel(excel_file_name, index=False)

print(f"Invitation data successfully saved to '{excel_file_name}'.")
