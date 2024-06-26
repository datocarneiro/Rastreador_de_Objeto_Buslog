import requests

url = "https://api.track3r.com.br/v2/api/Tracking"

payload = {
    "Sessao": "34766b43-ecce-4080-9eea-d6263fdf12aa",
    "CodigoServico": 1,
    "DataInicial": "2020-11-11",
    "DataFinal": "2020-11-11",
    "Pedidos": [
        {
            "ChaveNfe": "",
            "NotaFiscal": "",
            "Encomenda": ""
        },
        {
            "ChaveNfe": "",
            "NotaFiscal": "",
            "Encomenda": ""
        }
    ]
}

headers = {
    'Content-Type': 'application/json'
}

response = requests.post(url, headers=headers, json=payload)

print(response.text)
