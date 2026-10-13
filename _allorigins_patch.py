"""Patch requests.Session.get to relay KR open-data hosts via public CORS proxies.

Primary: proxy.cors.dev (path prefix). Fallback: api.allorigins.win/raw.
"""
import urllib.parse
import requests

_ORIG_GET = requests.Session.get
_RELAY_HOSTS = ("apis.data.go.kr", "www.bizinfo.go.kr", "bizinfo.go.kr")


def _patched_get(self, url, **kwargs):
    parsed = urllib.parse.urlparse(url)
    if parsed.hostname in _RELAY_HOSTS:
        params = kwargs.pop("params", None)
        if params:
            q = urllib.parse.urlencode(params, doseq=True, safe="")
            joiner = "&" if parsed.query else "?"
            url = url + joiner + q
        kwargs = dict(kwargs)
        kwargs.pop("proxies", None)
        if "timeout" in kwargs:
            t = kwargs["timeout"]
            if isinstance(t, (int, float)):
                kwargs["timeout"] = max(t, 90)
        relay = "https://proxy.cors.dev/" + url
        try:
            resp = _ORIG_GET(self, relay, **kwargs)
            if resp.status_code == 200 and len(resp.content) > 0:
                return resp
        except Exception as e:
            print(f"[relay_patch] cors.dev failed: {type(e).__name__}; trying allorigins", flush=True)
        relay2 = "https://api.allorigins.win/raw?url=" + urllib.parse.quote(url, safe="")
        return _ORIG_GET(self, relay2, **kwargs)
    return _ORIG_GET(self, url, **kwargs)


requests.Session.get = _patched_get
print("[relay_patch] Session.get -> proxy.cors.dev (+ allorigins fallback) for", _RELAY_HOSTS, flush=True)
