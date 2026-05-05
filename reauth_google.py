from google_auth_oauthlib.flow import InstalledAppFlow
import pickle

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
]
flow = InstalledAppFlow.from_client_secrets_file(
    r'F:\종소세2026\.credentials\client_secret.json', SCOPES)
creds = flow.run_local_server(port=0)
with open(r'F:\종소세2026\.credentials\token.pickle', 'wb') as f:
    pickle.dump(creds, f)
print('token.pickle 갱신 완료')
