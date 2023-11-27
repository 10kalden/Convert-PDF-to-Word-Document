import requests

api_key = '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'#key 
url = 'https://api.openai.com/v1/engines/davinci/completions' #endpoint
data = {
    'prompt': 'Hello, world!',
    'max_tokens': 5
}
headers = {
    'Content-Type': 'application/json',
    'Authorization': f'Bearer {api_key}'
}

response = requests.post(url, json=data, headers=headers)

if response.status_code == 200:
    print('API key and endpoint are workin.')
else:
    print(f'Error: {response.text}')
