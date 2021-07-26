import hashlib

text = 'admin_password_b1f62d4e'
text1 = 'admi1n_password_b1f62d4e'
enc = hashlib.md5()
enc.update(text.encode('utf-8'))
enc.update(text1.encode('utf-8'))
encText = enc.hexdigest()

print(encText)