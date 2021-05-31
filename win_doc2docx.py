import os
import sys

FILE_PATH = os.path.dirname(__file__)
PROJECT_PATH = os.path.abspath(os.path.join(FILE_PATH, '..', '..'))
sys.path.extend([FILE_PATH, PROJECT_PATH])


import argparse

parser = argparse.ArgumentParser(description='M03_03: Convert the doc files to docx files')
parser.add_argument('--input_dir', type = str, help='This dir contains the original contract files.')
parser.add_argument('--output_dir', type = str, help = 'This dir saves the converted contract files.')
args = parser.parse_args()

import shutil
from glob import glob
from tqdm import tqdm
from cimc_core.win_utils import doc2docx


def convert(contract_dir, output_dir):
    contract_dirname = os.path.basename(contract_dir)
    res_dir = os.path.join(output_dir, contract_dirname)

    if not os.path.exists(res_dir):
        os.makedirs(res_dir)

    files = [f for f in glob(os.path.join(contract_dir, '*')) if os.path.isfile(f)]
    for fp in files:
        fn_ext = os.path.basename(fp)
        fn, ext = os.path.splitext(fn_ext)
        if ext == '.doc':
            fpath = fp
            tpath = os.path.join(res_dir, '{}.docx').format(fn)
            doc2docx(fpath, tpath)
        else:
            fpath = fp
            tpath = os.path.join(res_dir, fn_ext)
            shutil.copyfile(fpath, tpath)


def batch_convert(input_dir, output_dir):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Traverse the input_dir, convert the doc files to docx files and keep other files as they are.
    contract_dirs = [cdir for cdir in glob(os.path.join(input_dir, '*')) if os.path.isdir(cdir)]
    for cdir in tqdm(contract_dirs):
        print("It is converting the contract from: {}".format(cdir))
        convert(cdir, output_dir)


def main(args):
    input_dir = args.input_dir
    output_dir = args.output_dir
    batch_convert(input_dir, output_dir)


if __name__ == '__main__':
    main(args)