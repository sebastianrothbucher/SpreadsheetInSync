from http.client import HTTPSConnection
from http.client import HTTPConnection
from ssl import _create_unverified_context
from base64 import b64encode

def httpclient(host, port, path, method='GET', contentType=None, content=None, user=None, password=None, ssl=False):
    if ssl:
        conn = HTTPSConnection(host, port, context=_create_unverified_context())
    else:
        conn = HTTPConnection(host, port)
    headers = {}
    if content and contentType: 
        headers['Content-type']=contentType
    if user and password:
        headers['Authorization']='Basic '+b64encode((user+':'+password).encode()).decode()
    conn.request(method, path, (content if content else None), headers)
    return conn.getresponse().read().decode()

#print(httpclient('localhost', 5984, '/test/_all_docs', 'GET', '', '', 'testuser', 'secret', False))
#print(httpclient('breezy-moose-84.localtunnel.me', 443, '/test/_all_docs', 'GET', '', '', 'testuser', 'secret', True))
