import base64


def encode(s):
    to_bytes = s
    if isinstance(s, str):
        to_bytes = s.encode('utf-8')

    return base64.b64encode(to_bytes).decode('utf-8')


def decode(s):
    return base64.b64decode(s).decode('utf-8')
