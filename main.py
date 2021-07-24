import requests
from bs4 import BeautifulSoup

def main():
    URL = "https://www.geeksforgeeks.org/data-structures/"
    r = requests.get(URL)
    soup = BeautifulSoup(r.content, 'html5lib')
    print(soup.prettify())

if __name__ == '__main__':
    main()

