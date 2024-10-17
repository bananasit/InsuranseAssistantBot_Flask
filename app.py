from flask import Flask, request
from twilio.twiml.messaging_response import MessagingResponse
import pandas as pd

app = Flask(__name__)

# Define the path to Excel File
EXCEL_FILE_PATH = 'E:/Vidhi/BSIT/Projects/Documents/Bot_Docs/insurance_data.xlsx'

@app.route("/whatsapp", methods=["POST"])
def whatsapp_reply():
    # Get the message the user sent to our Twilio number
    incoming_message = request.values.get('Body', '').strip()
    print(incoming_message)
    # Create a Twilio response object
    response = MessagingResponse()

    resp_msg = ''

    # Logic to respond
    if incoming_message.lower() in ['hi', 'hello']:
        resp_msg = "Hi there! How can I assist you?\nPlease provide me the Policy ID/Claim ID/Contract ID."
    elif incoming_message.isdigit():  # Check if the message is a number
            policy_details = get_policy_details(incoming_message)
            claim_details = get_claim_details(incoming_message)
            contract_details = get_contract_details(incoming_message)

            if policy_details:
                resp_msg = f"Policy Details:\n{policy_details}"
            elif claim_details:
                resp_msg = f"Claim Details:\n{claim_details}"
            elif contract_details:
                resp_msg = f"Contract Details:\n{contract_details}"
            else:
                resp_msg = "No details found for the given number."
    else:
            resp_msg = "I'm sorry, I don't understand that."
        
    response.message(resp_msg)

    print(response)
    return str(response)

def get_policy_details(policy_id):
    policies = fetch_data_from_excel('policies')
    for policy in policies:
        if str(policy['Policy ID']) == policy_id:
            return (
                f"Policy ID: {policy['Policy ID']},\n"
                f"Holder: {policy['Policyholder Name']},\n"
                f"Start Date: {policy['Start Date'].strftime('%Y-%m-%d')},\n"
                f"End Date: {policy['End Date'].strftime('%Y-%m-%d')},\n"
                f"Premium Amount: {policy['Premium Amount']}"
            )
    return None

def get_claim_details(claim_id):
    claims = fetch_data_from_excel('claims')
    for claim in claims:
        if str(claim['Claim ID']) == claim_id:
            return (
                f"Claim ID: {claim['Claim ID']},\n"
                f"Policy Number: {claim['Policy Number']},\n"
                f"Date of Claim: {claim['Date of Claim'].strftime('%Y-%m-%d')},\n"
                f"Amount Claimed: {claim['Amount Claimed']}"
            )
    return None

def get_contract_details(contract_id):
    contracts = fetch_data_from_excel('contracts')
    for contract in contracts:
        if str(contract['Contract ID']) == contract_id:
            return (
                f"Contract ID: {contract['Contract ID']},\n"
                f"Policy Number: {contract['Policy Number']},\n"
                f"Terms and Conditions: {contract['Terms and Conditions']}"
            )
    return None

# Fetching data from Excel sheets
def fetch_data_from_excel(sheet_name):
    try:
        df = pd.read_excel(EXCEL_FILE_PATH, sheet_name=sheet_name)
        return df.to_dict(orient='records')
    except Exception as e:
        print(f"Error fetching data from Excel: {e}")
        return []

if __name__ == '__main__':
    app.run(debug=True)