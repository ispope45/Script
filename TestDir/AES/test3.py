import time
import base64_2
import binascii
import hashlib

from Crypto.Cipher import AES
from Crypto import Random

from .securecookie import SecureCookie


def secure_cookie_serialize(value, secret_key, *, expires=None):
    """Secure Cookie Serialize"""
    x = SecureCookie(value, secret_key=secret_key)
    return x.serialize(expires=expires)


def secure_cookie_unserialize(value, secret_key):
    """Secure Cookie Unserialize"""

    return SecureCookie.unserialize(value, secret_key)


def make_key():
    now = int(time.time())

    key = str(now).encode('utf-8')
    return hashlib.sha256(key).hexdigest()


class DecryptionError(Exception):
    """복호화 오류"""

    def __init__(self, cause):
        self.cause = cause

    def __str__(self):
        return self.cause


class AESCipher:

    def __init__(self, key):
        self.bs = 16
        self.key = key[:32]

    def encrypt(self, raw):
        """암호화"""
        raw = self._pad(raw)
        iv = Random.new().read(AES.block_size)
        cipher = AES.new(self.key, AES.MODE_CBC, iv)
        hex_data = (iv.hex() + cipher.encrypt(raw).hex()).encode('utf-8')
        return base64_2.b64encode(hex_data)

    def decrypt(self, enc):
        """복호화"""
        try:
            enc = base64_2.b64decode(enc).decode('utf-8')
            enc = bytes(bytearray.fromhex(enc))
            iv = enc[:AES.block_size]
            cipher = AES.new(self.key, AES.MODE_CBC, iv)

            return self._unpad(cipher.decrypt(enc[AES.block_size:]).decode('utf-                                                                                                                                                             8'))
        except binascii.Error:
            raise DecryptionError('Incorrect padding')
        except UnicodeDecodeError:
            raise DecryptionError('Invalid key')
        except Exception as e:
            raise DecryptionError(e)

    def decrypt_and_b64encode(self, enc):
        """복화화된 평문을 base64 인코딩"""
        to_bytes = self.decrypt(enc)
        if isinstance(to_bytes, str):
            to_bytes = to_bytes.encode('utf-8')

        return base64_2.b64encode(to_bytes).decode('utf-8')

    def _pad(self, s):
        return s + (self.bs - len(s) % self.bs) * chr(self.bs - len(s) % self.bs                                                                                                                                                             )

    @staticmethod
    def _unpad(s):
        return s[:-ord(s[len(s) - 1:])]
