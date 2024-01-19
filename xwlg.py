import docx.document, docx.table
from requests.api import get
import pandas as pd
from pandas import DataFrame
from pathlib import Path
import json
from random import randint
import re
from distutils.dir_util import mkpath
from typing import Dict, Any
from os import getenv
from dotenv import load_dotenv
load_dotenv()

word = None
page = None

THESAURUS_API_KEY = getenv('THESAURUS_API_KEY')
if (not THESAURUS_API_KEY):
    print("Missing Thesaurus API key. Get one from https://www.dictionaryapi.com/products/api-collegiate-thesaurus/")
    exit()

def get_word_data(word: str) -> Dict[Any, Any]:
    cached_path = Path(__file__).parent.joinpath(f"cache/thesaurus/{word}.json")
    if cached_path.exists():
        # Get from cache
        with open(cached_path) as f:
            word_data = json.load(f)
    else:
        # Request word data from dictionaryapi.com
        request = get(f"https://www.dictionaryapi.com/api/v3/references/thesaurus/json/{word}?key={THESAURUS_API_KEY}")
        word_data = request.json()
        # Cache result
        mkpath(str(Path(__file__).parent.joinpath("cache/thesaurus")))
        with open(cached_path, "x") as f:
            json.dump(word_data, f)
    return word_data

def synonyms(previous_value: str, word_data: Dict[Any, Any]) -> str:
    previous_value = previous_value.strip()
    if previous_value:
        return previous_value
    try:
        s = word_data["meta"]["syns"][0][:randint(1,2)]
        return ", ".join(s)
    except KeyError:
        global word
        global page
        return input(f"Could not find synonyms for '{word}'{page}, please provide: ")

def format_sentence(raw: str):
    raw = raw.strip().capitalize()
    if raw[-1] not in [".", "!", "?"]:
        raw += "."
    return re.sub(r"\{.*?\}", "", raw)

def collocation(previous_value: str, word_data: Dict[Any, Any]) -> str:
    previous_value = previous_value.strip()
    if previous_value:
        return previous_value
    try:
        raw_collocation: str = word_data["def"][0]["sseq"][0][0][1]["dt"][1][1][0]["t"]
        return format_sentence(raw_collocation)
    except:
        global word
        global page
        return input(f"Could not find collocation for '{word}'{page}, please provide: ")

def get_definition(previous_value: str, word_data: Dict[Any, Any]) -> str:
    previous_value = previous_value.strip()
    if previous_value:
        return previous_value
    try:
        raw_definition = word_data["shortdef"][0]
        return format_sentence(raw_definition)
    except KeyError:
        global word
        global page
        return input(f"Could not find definition for '{word}'{page}, please provide: ")

def generate_word_list(table: docx.table.Table, output_path: Path|None = None) -> DataFrame:    
    word_list = DataFrame({cell.text: [] for cell in table.rows[0].cells[1:]})
    for row_number, row in enumerate(table.rows[1:]):
        # Row format: [index, page, word, synonyms, collocation, definition]
        row = [cell.text for cell in row.cells]
        word = row[2].strip()
        page = '' if row[1] == '' else f' ({row[1]})'
        print(word + page)
        if "" in row[3:]:
            # At least one of the last 3 cells in the row are empty
            skip_word = False
            word_data: Dict[Any, Any] = {}
            while True:
                # Get word data
                if " " in word:
                    print(f"Skipping '{word}' because it contains a space.")
                    skip_word = True
                    break
                word_data = get_word_data(word)
                # Word not found, give suggestions
                if type(word_data[0]) == str:
                    print(f"\n'{word}'{page} was not found, did you mean any of the following?")
                    for j, other in enumerate(word_data[:3]):
                        print(f"{j+1}: {other}")
                    selected = input(f"New word (1-{len(word_data[:3])}): ")
                    if selected == "":
                        skip_word = True
                        print("Skipped word")
                        break
                    print()
                    try:
                        word = word_data[int(selected)-1]
                    except:
                        word = selected
                else:
                    break
            if not skip_word:
                if len(word_data) > 1:
                    # Multiple definitions
                    selected = -1
                    while selected < 1 or selected > len(word_data):
                        print(f"\nMultiple definitions found for '{word}'{page}:")
                        for definition in word_data:
                            print(f"{row_number+1}: {definition['shortdef'][0]}")
                        try:
                            selected = int(input(f"Please select a definition (1-{len(word_data)}): "))
                            print()
                            if selected < 1 or selected > len(word_data):
                                raise Exception
                        except:
                            print("Invalid input.")
                    word_data = word_data[selected-1]
                else:
                    word_data = word_data[0]
                print()
                # Synonyms
                row[3] = synonyms(row[3], word_data)
                # Collocation
                row[4] = collocation(row[4], word_data)
                # Definition
                row[5] = get_definition(row[5], word_data)
                print()

        word_list = pd.concat([word_list, DataFrame({table.rows[0].cells[i].text: [row[i]]
                for i in range(len(table.rows[0].cells))[1:]}, index=[row_number+1])])
        if (output_path):
            word_list.to_excel(output_path)
    return word_list