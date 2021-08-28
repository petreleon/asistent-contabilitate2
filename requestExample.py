import requests

url = 'https://webservicesp.anaf.ro/PlatitorTvaRest/api/v5/ws/tva'
myobj = [
    {
        "cui": 31220820, "data":"2021-08-26"
    },
    {
        "cui": 31233499, "data":"2021-08-26"
    }		
]


x = requests.post(url, json = myobj).json()


print(x)