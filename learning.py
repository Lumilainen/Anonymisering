import json
import os

LEARNING_FOLDER = "learning"
IGNORE_FILE = os.path.join(LEARNING_FOLDER, "ignore_words.json")
FORCED_FILE = os.path.join(LEARNING_FOLDER, "forced_names.json")


def ensure_learning_folder():

    if not os.path.exists(LEARNING_FOLDER):
        os.makedirs(LEARNING_FOLDER)


def load_list(file):

    ensure_learning_folder()

    if not os.path.exists(file):
        return set()

    with open(file, "r", encoding="utf-8") as f:
        return set(json.load(f))


def save_list(file, data):

    ensure_learning_folder()

    with open(file, "w", encoding="utf-8") as f:
        json.dump(sorted(list(data)), f, ensure_ascii=False, indent=2)


def load_ignore():

    return load_list(IGNORE_FILE)


def load_forced():

    return load_list(FORCED_FILE)


def add_ignore(word):

    data = load_ignore()
    data.add(word)
    save_list(IGNORE_FILE, data)


def add_forced(name):

    data = load_forced()
    data.add(name)
    save_list(FORCED_FILE, data)