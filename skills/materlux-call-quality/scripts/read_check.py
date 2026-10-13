"""Read-only Materlux API client. Secret supplied by the local process environment."""
import os, json, re, urllib.request, urllib.error
from urllib.parse import urlsplit
class SafeError(Exception):
    pass
class NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, *args, **kwargs):
        raise SafeError('API_REDIRECT_BLOCKED')
class Client:
    def __init__(self):
        base=os.environ.get('MATERLUX_WEBHOOK','').strip()
        p=urlsplit(base)
        expected_host=os.environ.get('MATERLUX_PORTAL_HOST','').strip().lower()
        if (p.scheme!='https' or not expected_host or p.hostname!=expected_host or p.username or p.password
            or p.query or p.fragment or not re.fullmatch(r'/rest/[0-9]+/[A-Za-z0-9_-]+/?',p.path)):
            raise SafeError('WEBHOOK_MISSING_OR_INVALID')
        self._base=base.rstrip('/')+'/'
        self._opener=urllib.request.build_opener(NoRedirect())
        self.verified=True
    def call(self, method, params):
        if method not in {'profile','voximplant.statistic.get','user.get','disk.file.get','crm.activity.get'}:
            raise SafeError('METHOD_NOT_ALLOWED')
        req=urllib.request.Request(self._base+method+'.json',data=json.dumps(params).encode(),
             headers={'Content-Type':'application/json'},method='POST')
        try:
            with self._opener.open(req,timeout=25) as response:
                result=json.loads(response.read(4000000))
        except urllib.error.HTTPError as error:
            raise SafeError('HTTP_'+str(error.code)) from None
        except Exception:
            raise SafeError('NETWORK_OR_RESPONSE_ERROR') from None
        if result.get('error'):
            code=str(result['error'])
            raise SafeError('API_'+(code if re.fullmatch(r'[A-Za-z0-9_.-]{1,80}',code) else 'ERROR'))
        if 'result' not in result:
            raise SafeError('UNEXPECTED_RESULT')
        return result
