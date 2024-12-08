#!/usr/bin/env python3
# -*- coding: utf-8 -*-
#
# Eli Words Processor.
#
# Just give a cool description to your program ;)
#
import argparse
import pandas as pd
import openpyxl
import yaml as YAML
import re

config = {}
words_map = {}
word_exclusions = []
stanza_nlp = None
stanza_phases_map = {}
cli_args = {}

def parse_cli_args():
    ''' Command line argument processing '''

    parser = argparse.ArgumentParser(description='Process messages from XLS data')
    parser.add_argument('-i', '--input-file', action='store', default='ewop-data.xlsx',
                        help='Input data file name')
    parser.add_argument('-o', '--ouput-file', action='store', default='ewop-results.xlsx',
                        help='Results file name')
    parser.add_argument('-c', '--config-file', action='store', default='ewop-config.yaml',
                        help='Configuration file name')    
    parser.add_argument('-v', '--verbose', action='store_true', default=False,
                        help='Show additional debug info')
    
    cli_args = parser.parse_args()
    return cli_args

def load_config(config_file):
    ''' Load config from YAML file (ewop-config.yaml) '''

    with open(config_file, "r") as f:
        config = YAML.safe_load(f)

    # Compile exclusion patterns
    for p in config["word"]["exclusions"]:
        rep = re.compile(p)
        word_exclusions.append(rep)

    return config

def filter_word(word):
    ''' Filter unwanted words '''

    # Check miminal length
    if (len(word) < 2): return True
    
    # Check exclusion patterns
    for rep in word_exclusions:
        if rep.match(word): return True

    return False

def capture_words(sentence):
    ''' Capture non-filtered words from a sentence '''

    wlist = re.split('\W+', sentence)
    current_keys = words_map.keys()

    # Count captured words
    for w in wlist:
        if filter_word(w): 
            continue

        if w in current_keys:
            words_map[w] += 1
        else:      
            words_map[w] = 1

    return wlist

def run_basic_processor():
    ''' Count word repetitions '''

    print("=== Basic processor ===")

    # Applying the function row-wise
    result = xlsdata.apply(total_words, axis=1)

    if cli_args.verbose:
        for res in result:
            print(res)

    sorted_by_values = dict(sorted(words_map.items(), key=lambda item: item[1], reverse=True))
    if cli_args.verbose: print(sorted_by_values)

    print(f"Words count: {len(sorted_by_values)}")

    df = pd.DataFrame.from_dict(sorted_by_values, orient='index', columns=['count'])
    return df

def run_stanza_processor(sentence):
    ''' Run Stanford NLP 'Stanza' analisys '''

    doc = stanza_nlp(sentence)

    for sent in doc.sentences:
        for dep in sent.dependencies:
            if dep[1] == 'amod':
                phrase = f"{dep[0].text} {dep[2].text}"    

                if phrase in stanza_phases_map.keys():
                    stanza_phases_map[phrase] += 1
                else:      
                    stanza_phases_map[phrase] = 1

        # doc.sentences[0].print_dependencies() 
    return

def generate_basic_results():
    ''' Count word repetitions '''

    print("=== Basic processing ===")

    sorted_by_values = dict(sorted(words_map.items(), key=lambda item: item[1], reverse=True))
    if cli_args.verbose: print(sorted_by_values)

    print(f"Words count: {len(sorted_by_values)}")

    df = pd.DataFrame.from_dict(sorted_by_values, orient='index', columns=['count'])
    return df

def generate_stanza_results():
    ''' Count repetitions of phrases '''

    print("=== Stanza processing ===")

    sorted_by_values = dict(sorted(stanza_phases_map.items(), key=lambda item: item[1], reverse=True))
    if cli_args.verbose: print(sorted_by_values)

    print(f"Phrases count: {len(sorted_by_values)}")

    df = pd.DataFrame.from_dict(sorted_by_values, orient='index', columns=['count'])
    return df

def process_row(row):
    ''' List total non-filtered words '''

    wlist = []

    sentence = str(row[0]).lower()

    if config['processors']['basic']:
        wlist = capture_words(sentence)

    if config['processors']['stanza']:
        run_stanza_processor(sentence)

    return wlist

def process_data():
    df_results = []

    # Applying the function row-wise
    result = xlsdata.apply(process_row, axis=1)

    if cli_args.verbose:
        for r in result:
            print(r)

    if config['processors']['basic']:
        df_words = generate_basic_results()
        df_results.append(['WordCount', df_words])

    if config['processors']['stanza']:
        df_stanza = generate_stanza_results()
        df_results.append(['StanfordNLP', df_stanza])

    return df_results

def write_xls(dataframes=[]):
    ''' Save results to xls file '''

    with pd.ExcelWriter(cli_args.ouput_file) as writer:
        for df in dataframes:
            df[1].to_excel(writer, sheet_name=df[0])
        
# Main 
cli_args = parse_cli_args()
config = load_config(cli_args.config_file)

if config['processors']['stanza']:
    import stanza
    stanza_nlp = stanza.Pipeline('es', download_method=stanza.DownloadMethod.REUSE_RESOURCES)

xlsdata = pd.read_excel(cli_args.input_file, usecols='A,B', skiprows=1)

results = process_data()
write_xls(results)
