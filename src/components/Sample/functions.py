import io
import requests
from os import system
from PIL import Image
from bs4 import BeautifulSoup

def openImage(url):
    soup = BeautifulSoup(requests.get(url).text, 'lxml')
    # print(soup)
    # img = soup.select('app > section > div.container > div > div > div > img')
    img=soup.find('img',alt='formbuilder')
    print(img['src'])
    # return Image.open(io.BytesIO(requests.get(url).content))
    

if __name__ == "__main__":
    openImage('https://forms.app/recordimage/6308f6c70169057963106ed8/435a4a43-255d-11ed-8d27-4201c0a80002')
    # a.save('test.png')