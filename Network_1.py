import socket
s=socket.socket(socket.AF_INET, socket.SOC)
s.gethostname()[0]


from dns import reversename, resolver

rev_name = reversename.from_address('120.165.247.93')
reversed_dns = str(resolver.query(rev_name,"PTR")[0])
# 'crawl-203-208-60-1.googlebot.com.'

import socket

reversed_dns = socket.gethostbyaddr('120.165.247.93')
# ('crawl-203-208-60-1.googlebot.com', ['1.60.208.203.in-addr.arpa'], ['203.208.60.1'])
reversed_dns[0]
# 'crawl-203-208-60-1.googlebot.com'