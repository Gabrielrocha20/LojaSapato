from hashlib import md5

texto = "Admin1".encode("utf8")
print(texto)
hash = md5(texto).hexdigest()
print(hash)
