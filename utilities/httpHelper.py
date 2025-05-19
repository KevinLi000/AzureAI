class httpHelper:
    @staticmethod
    def get(url):
        import requests
        response = requests.get(url)
        return response.text

    @staticmethod
    def post(url, data):
        import requests
        response = requests.post(url, json=data)
        return response.text