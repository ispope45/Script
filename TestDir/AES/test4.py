import time
import hashlib
import base64
from Cryptodome.Cipher import AES
from Cryptodome import Random
import binascii

def make_key():
    now = int(time.time())
    print(now)
    now = 1646571469

    key = str(now).encode('utf-8')
    print(key)
    return hashlib.sha256(key).hexdigest()


def _pad(s):
    bs = 16
    return s + (bs - len(s) % bs) * chr(bs - len(s) % bs)


class AESCipher:
    def __init__(self, key):
        self.bs = 16
        self.key = key[:32]

    def encrypt(self, raw):
        """?뷀샇??"""
        raw = self._pad(raw)
        iv = Random.new().read(AES.block_size)
        cipher = AES.new(self.key, AES.MODE_CBC, iv)
        hex_data = (iv.hex() + cipher.encrypt(raw).hex()).encode('utf-8')
        return base64.b64encode(hex_data)

    def decrypt(self, enc):
        """蹂듯샇??"""
        # try:
        enc = base64.b64decode(enc).decode('utf-8')
        enc = bytes(bytearray.fromhex(enc))
        iv = enc[:AES.block_size]
        cipher = AES.new(self.key, AES.MODE_CBC, iv)

        return self._unpad(cipher.decrypt(enc[AES.block_size:]).decode('utf-8'))
        # except binascii.Error:
        #     raise DecryptionError('Incorrect padding')
        # except UnicodeDecodeError:
        #     raise DecryptionError('Invalid key')
        # except Exception as e:
        #     raise DecryptionError(e)

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


# print(make_key())
# key = make_key()
# # key = 'cebb417615cf507d331229901cb3ec4f8cb3509e3a9f83279ac6f7cb2f0cb4c2'
# print(len(key))
# print(type(key.encode('utf-8')))
# print(len(key.encode('utf-8')))
# iv = Random.new().read(AES.block_size)
# print(iv)
# print(len(iv))
# cipher = AES.new(key[:32].encode('utf-8'), AES.MODE_CBC, iv)
pw = 'qhdks00@!'
# raw = bytes(_pad(pw), 'utf-8')
# hex_data = (iv.hex() + cipher.encrypt(raw).hex()).encode('utf-8')
# user_pw = base64.b64encode(hex_data).decode('utf-8')

# print(user_pw)
pw_val = 'ZmYwYjdhMTRiZDVmZmFhOGQ1N2QwYjllY2I1MTY1NjQ0ZmYxNTdmZDkxOTk2ZGRmNzAwYjY1YTY3ODEzY2YzMw=='
# print(pw_val)
# dec_pw = base64.b64decode(pw_val)
dec_pw = pw_val
print(dec_pw)
print(base64.b64decode(pw_val))

key = '6aa3e9f1293a43404469ac0ffcfb8665b6effd81ef44acb9101c47c9908523ca'

a = AESCipher(key.encode('utf-8'))
plane1 = a.decrypt_and_b64encode(dec_pw)
print(plane1)

iv = Random.new().read(AES.block_size)
key2 = bytes(key[:32], 'utf-8')
cipher = AES.new(key2, AES.MODE_CBC, iv)
raw = bytes(_pad(pw), 'utf-8')
hex_data = (iv.hex() + cipher.encrypt(raw).hex()).encode('utf-8')
user_pw = base64.b64encode(hex_data).decode('utf-8')
print(user_pw)
# print(type(plane1.encode('utf-8')))
# enc1 = a.encrypt(plane1)
# print(type(enc1))

