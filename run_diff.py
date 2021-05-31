import os
import sys

FILE_PATH = os.path.dirname(__file__)
PROJECT_PATH = os.path.abspath(os.path.join(FILE_PATH, '..', '..'))
sys.path.extend([FILE_PATH, PROJECT_PATH])


import argparse

parser = argparse.ArgumentParser(description='M03_03: Compare the contract to be signed with the signed one.')
parser.add_argument('--input_dir', type = str, help='This dir contains contracts to be compared.')
parser.add_argument('--output_dir', type = str, help = 'This dir saves the diff results.')
parser.add_argument('--libre_timeout', type = int, default= 120)
parser.add_argument('--output_scores', action= 'store_true')
args = parser.parse_args()

import pandas as pd
from tqdm import tqdm
from glob import glob
from cimc_core.m03_03.compare import Compare_Difflib3 as Compare_Difflib
from cimc_core.m03_03 import config as m03_03_config
m03_03_config.LIBRE_TIMEOUT = args.libre_timeout # This line should be set before DocReader is imported
from cimc_core.m03_03.docreader import DocReader


def get_unsigned(files):
    unsigned_files = []
    for fp in files:
        fn_ext = os.path.basename(fp)
        fn, ext = os.path.splitext(fn_ext)
        if (ext in ['.doc', '.docx']) and ('待签' in fn):
            unsigned_files.append(fp)

    return unsigned_files


def get_signed(files):
    signed_files = []
    for fp in files:
        fn_ext = os.path.basename(fp)
        fn, ext = os.path.splitext(fn_ext)
        if ext in ['.docx'] and "已签" in fn:
            signed_files.append(fp)

    return signed_files


def compare_file(unsigned_file, signed_file):
    unsigned_docreader = DocReader(unsigned_file)
    signed_docreader = DocReader(signed_file)
    error_msg_list = []
    if unsigned_docreader.data == []:
        error_msg_list.append('Failed to read data from unsigned file: {}'.format(unsigned_file))

    if signed_docreader.data == []:
        error_msg_list.append('Faild to read data from signed file: {}'.format(signed_file))

    if error_msg_list != []:
        error_msg = '\n'.join(error_msg_list) + '\n'
        return -1, error_msg, None, None, None

    comp = Compare_Difflib()
    diff_df, score = comp(unsigned_docreader.data, signed_docreader.data)
    avg_score = round(diff_df['score'].mean())
    return 0, '', diff_df, score, avg_score


def compare(contract_dir, output_dir):
    contract_dirname = os.path.basename(contract_dir)
    res_dir = os.path.join(output_dir, contract_dirname)

    if not os.path.exists(res_dir):
        os.makedirs(res_dir)

    error_file = os.path.join(res_dir, 'error_msg.txt')

    files = [f for f in glob(os.path.join(contract_dir, '*')) if os.path.isfile(f)]
    unsigned_files = get_unsigned(files)
    signed_files = get_signed(files)

    error_msg_list = []
    if unsigned_files == []:
        error_msg_list.append('No unsigned files are provided or found.')

    if signed_files == []:
        error_msg_list.append('No signed files are provided or found.')

    if error_msg_list != []:
        error_msg = '\n'.join(error_msg_list) + '\n'
        with open(error_file, 'w', encoding='utf8') as f:
            f.write(error_msg)

        return

    diff_scores = []
    for uf in unsigned_files:
        for sf in signed_files:
            error_code, error_msg, diff_df, score, avg_score = compare_file(uf, sf)
            if error_code < 0:
                if not os.path.exists(error_file):
                    with open(error_file, 'w') as f:
                        pass

                with open(error_file, 'a') as f:
                    f.write(error_msg)

                continue

            ufn_ext = os.path.basename(uf)
            ufn, uext = os.path.splitext(ufn_ext)
            sfn_ext = os.path.basename(sf)
            sfn, sext = os.path.splitext(sfn_ext)

            diff_xlsx = os.path.join(res_dir, '{}_vs_{}.xlsx'.format(ufn, sfn))
            diff_df.to_excel(diff_xlsx)
            diff_scores.append([contract_dirname, ufn_ext, sfn_ext, score, avg_score])

    return diff_scores


def batch_compare(input_dir, output_dir, output_scores):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Traverse the input_dir, and compare each contract
    diff_scores = []
    contract_dirs = [cdir for cdir in glob(os.path.join(input_dir, '*')) if os.path.isdir(cdir)]
    for cdir in tqdm(contract_dirs):
        print("It is comparing the contact from: {}".format(cdir))
        _diff_scores =  compare(cdir, output_dir)
        if _diff_scores is not None:
            diff_scores.extend(_diff_scores)

    if output_scores:
        diff_scores_df = pd.DataFrame(diff_scores, columns= ['contract', 'ufn', 'sfn', 'score', 'avg_score'])
        diff_scores_xlsx = os.path.join(output_dir, 'diff_scores.xlsx')
        diff_scores_df.to_excel(diff_scores_xlsx)


def main(args):
    input_dir = args.input_dir
    output_dir = args.output_dir
    output_scores = args.output_scores
    batch_compare(input_dir, output_dir, output_scores)


if __name__ == '__main__':
    main(args)