import requests

def main():
    URL = "https://www.geeksforgeeks.org/data-structures/"
    r = requests.get(URL)
    print(r.content)

if __name__ == '__main__':
    main()

