import requests
from lxml.html import fromstring

src = 'https://api64.ipify.org'
response = requests.get('https://free-proxy-list.net/')
parser = fromstring(response.text)
proxies = set()

for i in parser.xpath('//tbody/tr'):
    if i.xpath('.//td[7][contains(text(),"yes")]'):
        proxy = ":".join([i.xpath('.//td[1]/text()')[0],
                          i.xpath('.//td[2]/text()')[0]])
        proxies.add(proxy)

if proxies:
    for proxy in proxies:
        try:
            print("IN")
            r = requests.get(
				src, proxies={"http": proxy, "https": proxy}, timeout=5)
            file_request_succeed = r.ok
            if file_request_succeed:
                print('Rotated IP %s succeed' % proxy)
                break
        except Exception as e:
            print('Rotated IP %s failed (%s)' % (proxy, str(e)))

