import json
import requests
from zoho_api.auth import ZohoAuth
from credentials.zoho.zoho_tokens import zp_oauth, zb_oauth, zcrm_oauth, zc_oauth, success_status_codes

zoho_people_auth = ZohoAuth(zp_oauth['client_id'], zp_oauth['client_secret'], zp_oauth['refresh_token'])
zoho_books_auth = ZohoAuth(zb_oauth['client_id'], zb_oauth['client_secret'], zb_oauth['refresh_token'])
zoho_creator_auth = ZohoAuth(zc_oauth['client_id'], zc_oauth['client_secret'], zc_oauth['refresh_token'])
zoho_crm_auth = ZohoAuth(zcrm_oauth['client_id'], zcrm_oauth['client_secret'], zcrm_oauth['refresh_token'])

def api_request(url, source, method, post_data):
    access_headers = {}
    if source == 'zoho_people':
        zoho_people_auth.get_or_refresh_access_token()
        access_headers = {
            'Authorization': f'Zoho-oauthtoken {zoho_people_auth.access_token}'
        }
    elif source == 'zoho_books':
        zoho_books_auth.get_or_refresh_access_token()
        access_headers = {
            'Authorization': f'Zoho-oauthtoken {zoho_books_auth.access_token}'
        }
    elif source == 'zoho_creator':
        zoho_creator_auth.get_or_refresh_access_token()
        access_headers = {
            'Authorization': f'Zoho-oauthtoken {zoho_creator_auth.access_token}'
        }
    elif source == 'zoho_crm':
        zoho_crm_auth.get_or_refresh_access_token()
        access_headers = {
            'Authorization': f'Zoho-oauthtoken {zoho_crm_auth.access_token}'
        }

    if access_headers:
        if method == 'get':
            response = requests.get(url, headers=access_headers)
        elif method == 'put':
            response = requests.put(url, headers=access_headers, data=json.dumps(post_data))
        elif method == 'post':
            response = requests.post(url, headers=access_headers, data=json.dumps(post_data))
        elif method == 'patch':
            response = requests.patch(url, headers=access_headers, data=json.dumps(post_data))
        if response.status_code in success_status_codes:
            return response.json()
        else:
            return None
    else:
        return None
