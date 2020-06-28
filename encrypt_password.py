from cryptography.fernet import Fernet
import os
import sys
import argparse


class EncryptPassword:
    def __init__(self):
        if not os.path.isfile('secret.key'):
            self.generate_key()
        return

    def generate_key(self):
        """
        Generates a key and save it into a file
        """
        key = Fernet.generate_key()
        with open("secret.key", "wb") as key_file:
            key_file.write(key)
        return

    def load_key(self):
        """
        Load the previously generated key
        """
        return open("secret.key", "rb").read()

    def encrypt_message(self, password):
        """
        Encrypts a message
        """
        key = self.load_key()
        encoded_password = password.encode()
        f = Fernet(key)
        encrypted_password = f.encrypt(encoded_password)
        print("Your Excrypted Password is :\n{}".format(encrypted_password))

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--password', help='password to be encrypted', required=True)
    args = parser.parse_args()
    encry_obj = EncryptPassword()
    encry_obj.encrypt_message(args.password)

