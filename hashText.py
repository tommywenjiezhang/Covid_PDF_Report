import hashlib
import pickle

def hash_message(message):
    result  = hashlib.sha256(message.encode())
    return result.hexdigest()

def save_hash_to_file(hashed_message,file_path):
    payload = {}
    if hashed_message:
        payload["last_hash_messsage"] = hashed_message
        with open(file_path,'wb') as hashed_file:
            pickle.dump(payload, hashed_file)

def load_hash_to_file(file_path):
    with open(file_path, 'rb') as hashed_file:
        payload = pickle.load(hashed_file)
        return payload

if __name__ == "__main__":
    pass

