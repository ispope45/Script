import time
import base64
import binascii
import hashlib

from Crypto.Cipher import AES
from Crypto import Random

# from .securecookie import SecureCookie
#
#
# def secure_cookie_serialize(value, secret_key, *, expires=None):
#     """Secure Cookie Serialize"""
#     x = SecureCookie(value, secret_key=secret_key)
#     return x.serialize(expires=expires)
#
#
# def secure_cookie_unserialize(value, secret_key):
#     """Secure Cookie Unserialize"""
#
#     return SecureCookie.unserialize(value, secret_key)


def make_key():
    now = int(time.time())

    key = str(now).encode('utf-8')
    print(type(hashlib.sha256(key).hexdigest()))
    return hashlib.sha256(key).hexdigest()


class DecryptionError(Exception):
    """蹂듯샇???ㅻ쪟"""

    def __init__(self, cause):
        self.cause = cause

    def __str__(self):
        return self.cause


class AESCipher:

    def __init__(self, key):
        self.bs = 16
        self.key = key[:32]

    def encrypt(self, raw):
        """?뷀샇??"""
        raw = self._pad(raw)
        iv = Random.new().read(AES.block_size)
        print("iv " + str(iv))
        cipher = AES.new(self.key, AES.MODE_CBC, iv)
        hex_data = (iv.hex() + cipher.encrypt(raw).hex()).encode('utf-8')
        print(hex_data)
        return base64.b64encode(hex_data)

    def decrypt(self, enc):
        """蹂듯샇??"""
        try:
            enc = base64.b64decode(enc).decode('utf-8')
            enc = bytes(bytearray.fromhex(enc))
            iv = enc[:AES.block_size]
            cipher = AES.new(self.key, AES.MODE_CBC, iv)

            return self._unpad(cipher.decrypt(enc[AES.block_size:]).decode('utf-8'))
        except binascii.Error:
            raise DecryptionError('Incorrect padding')
        except UnicodeDecodeError:
            raise DecryptionError('Invalid key')
        except Exception as e:
            raise DecryptionError(e)

    def decrypt_and_b64encode(self, enc):
        """蹂듯솕?붾맂 ?됰Ц??base64 ?몄퐫??"""
        to_bytes = self.decrypt(enc)
        if isinstance(to_bytes, str):
            to_bytes = to_bytes.encode('utf-8')

        return base64.b64encode(to_bytes).decode('utf-8')


    def _pad(self, s):
        return s + (self.bs - len(s) % self.bs) * chr(self.bs - len(s) % self.bs)

    @staticmethod
    def _unpad(s):
        return s[:-ord(s[len(s) - 1:])]


if __name__ == "__main__":
    # @staticmethod
    def _pad(s):
        bs = 16
        return s + (bs - len(s) % bs) * chr(bs - len(s) % bs)

    def _unpad(s):
        return s[:-ord(s[len(s) - 1:])]

    def _decrypt_and_b64encode(key, enc):
        """AES256 蹂듯샇??洹몃━怨?Base64 ?몄퐫??"""

        try:
            cipher1 = AESCipher(key)
            return cipher1.decrypt_and_b64encode(enc)
        except Exception as e:
            print(('Failed decryption: {}' + str(e)))

        return ''

    key = make_key()
    key2 = "e673931e72845037646a64761e2f193d6d58314cf76cbe4b22c344808cc11fef"
    print(key)
    # print(type(hashlib.sha256(key2).hexdigest()))
    # print(type(key))
    enc = "secui00@!"
    enc2 = "MzBjMTI0ZWU2ZjRkNGMxZDY4OTQ4YzVlZTMyMTViZjlkMGU3N2JkMDQ2MGY3YTMxZWU4MjE3ZGUyM2JiMGE4OQ=="

    csrf_token = "e673931e72845037646a64761e2f193d6d58314cf76cbe4b22c344808cc11fef"
    login_pw = "secui00@!"

    res1 = base64.b64decode(enc2).decode('utf-8')
    print(res1)
    res2 = bytes(bytearray.fromhex(res1))
    print(res2)
    iv = res2[:AES.block_size]
    print(iv)
    print(Random.new().read(AES.block_size))
    print(AES.block_size)
    keyv = bytes(key2[:32], 'utf-8')
    print(keyv)
    print(len(keyv))
    cipher = AES.new(keyv, AES.MODE_CBC, iv)
    print(cipher)
    enc3 = res2[16:]
    data = _unpad(cipher.decrypt(enc3).decode('utf-8'))
    print(data)

    # dec_b64_token = base64.b64decode(csrf_token).decode('utf-8')
    # hex_token = bytes(bytearray.fromhex(dec_b64_token))
    # iv = hex_token[:AES.block_size]
    iv = Random.new().read(AES.block_size)
    key = bytes(csrf_token[:32], 'utf-8')
    raw = bytes(_pad(login_pw), 'utf-8')
    cipher = AES.new(key, AES.MODE_CBC, iv)
    hex_data = (iv.hex() + cipher.encrypt(raw).hex()).encode('utf-8')
    print(hex_data)
    print(base64.b64encode(hex_data))


    # res = _decrypt_and_b64encode(key, enc2)
    # print(res)

    #
    # cipher = AESCipher('e673931e72845037646a64761e2f193d6d58314cf76cbe4b22c344808cc11fef')
    # print(cipher.key)
    # enc = "secui00@!"
    # cipher2 = cipher.encrypt(enc)
    # print(cipher2)
    # bs = 16
    # s = "secui00@!"
    # len1 = bs - len(s) % bs
    # print(len1)
    # res = s + (bs - len(s) % bs) * chr(bs - len(s) % bs)
    # print(res)
    # # cipher2 = cipher.encrypt(enc)
    # # print(cipher2)
    # cipher2 = cipher.decrypt_and_b64encode(res)
    # print(cipher2)


    # print(enc.key)
    # res = enc.encrypt("secui00@!")
    # print(res)
    # cipher = AESCipher(b"e673931e72845037646a64761e2f193d6d58314cf76cbe4b22c344808cc11fef")
    #
    # query = "MzBjMTI0ZWU2ZjRkNGMxZDY4OTQ4YzVlZTMyMTViZjlkMGU3N2JkMDQ2MGY3YTMxZWU4MjE3ZGUyM2JiMGE4OQ=="
    #
    # print(cipher.decrypt_and_b64encode(query))
