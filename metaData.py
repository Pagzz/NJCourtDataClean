import aiohttp
import asyncio
import pandas as pd  # Ensure pandas is installed

async def get_auth_token():
    async with aiohttp.ClientSession() as session:
        url = 'http://www.datairis.co/V1/auth/subscriber/'
        headers = {
            'SubscriberUsername': 'mintmediaapp',
            'SubscriberPassword': 'FNCcKWyi0bAXnoPAka8L',
            'AccountUsername': 'mintmediaact',
            'AccountPassword': 'reference',
            'AccountDetailsRequired': 'true',
            'SubscriberID': '224',
        }
        params = {'AccessToken': 'ad381-a59ec-4d7d'}
        async with session.get(url, headers=headers, params=params) as response:
            assert response.status == 200, "Authentication failed"
            data = await response.json()
            return data['Response']['responseDetails']['TokenID']

async def get_metadata(token_id, database_type='consumer'):
    async with aiohttp.ClientSession() as session:
        url = f'http://www.datairis.co/V1/search/metadata/{database_type}'
        headers = {'TokenID': token_id, 'Content-Type': 'application/json'}
        async with session.get(url, headers=headers) as response:
            assert response.status == 200, "Failed to fetch metadata"
            data = await response.json()
            if 'Response' in data and 'responseDetails' in data['Response'] and 'Metadata' in data['Response']['responseDetails']:
                search_fields = [field['id'] for field in data['Response']['responseDetails']['Metadata']]
                return search_fields
            else:
                print("Search fields not found in metadata response")
                return []

async def main():
    token_id = await get_auth_token()
    search_fields = await get_metadata(token_id, 'consumer')
    
    # Save to Excel
    df = pd.DataFrame(search_fields, columns=['Search Fields'])
    excel_path = 'C:\\Users\\rp122\\OneDrive\\Desktop\\dataUSA\\Excel\\Metadata.xlsx'
    df.to_excel(excel_path, index=False)
    print(f"Search Fields saved to {excel_path}")

if __name__ == '__main__':
    asyncio.run(main())
