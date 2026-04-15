import os
import json
import bson
import codecs
from Crypto.Cipher import AES
from Crypto.Hash import SHA256
import gzip


class EncoderForBytesObj(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, (bytes, bytearray)):
            return obj.decode("utf-8")  # <- or any other encoding of your choice
        # Let the base class default method raise the TypeError
        return json.JSONEncoder.default(self, obj)


def decrypt(input_file, output_file, key, pretty):
    with open(input_file, 'rb') as f:
        ciphertext = f.read()
    
    decrypted_data = aes_decrypt(ciphertext, key)
    decompressed_data = gzip.decompress(decrypted_data)
    json_data = decode_bson(decompressed_data)

    with open(output_file, 'w', encoding='utf8') as f:
        indent = 2 if pretty else None
        json.dump(json_data, f, cls = EncoderForBytesObj, indent = indent)


def encrypt(input_file, output_file, key):
    with codecs.open(input_file, 'r') as f:
        json_data = json.load(f)

    bson_data = encode_bson(json_data)
    encrypted_data = aes_encrypt(bson_data, key)

    with open(output_file, 'wb') as f:
        f.write(encrypted_data)


def aes_decrypt(ciphertext, key):
    nonce, authtag = ciphertext[:12], ciphertext[-16:]
    cipher = AES.new(key, AES.MODE_GCM, nonce)
    try:
        data = cipher.decrypt_and_verify(ciphertext[12:-16], authtag)
    except Exception as e:
        exit(f'Error while decrypt: {str(e)}\nCheck your password')
    return data


def aes_encrypt(text, key):
    cipher = AES.new(key, AES.MODE_GCM, nonce=os.urandom(12))
    nonce = cipher.nonce
    ciphertext, tag = cipher.encrypt_and_digest(text)
    return nonce + ciphertext + tag


def password_to_key(password):
    h = SHA256.new()
    h.update(password.encode())
    return h.digest()


def decode_bson(data):
    decoded_data = bson.BSON(data).decode()
    return decoded_data


def encode_bson(data):
    encoded_data = bson.BSON.encode(data)
    return encoded_data
