from concurrent.futures import ThreadPoolExecutor, wait
import subprocess
import logging
import os

def extract(infile: str, outfile: str) -> None:
    logging.info(f"Begin extract ${infile}")
    process = subprocess.run(['ffmpeg', '-y', '-i', infile, outfile], capture_output=True)
    if process.returncode != 0:
        logging.error(f'Extract {infile} failed!')
        logging.error(process.stderr.decode('utf8'))
    else:
        logging.info(f'Extract {infile} sucessfully!')


def get_file_list(path):
    extract_file_list = []
    if not os.path.isdir(path):
        logging.error(f'{path} is not a directory!')
        return None
    for filename in os.listdir(path):
        file = os.path.join(path, filename)
        if os.path.isfile(file) and filename.endswith('mp4'):
            extract_file_list.append(file)
    logging.info(f'Detect {len(extract_file_list)} files from {path}')
    return extract_file_list

def process(extract_file_list: list, out_path: str):
    if not os.path.exists(out_path):
        os.mkdir(out_path)
    with ThreadPoolExecutor(max_workers=8) as pool:
        feature_list = []
        for infile in extract_file_list:
            filename = os.path.basename(infile)[:-4]
            filename += '.mp3'
            outfile = os.path.join(out_path, filename)
            feature = pool.submit(extract, infile, outfile)
            feature_list.append(feature)

        wait(feature_list)
    logging.info('All task done!')

if __name__ == '__main__':
    logging.basicConfig(format='[%(asctime)s] %(lineno)s %(levelname)s: %(message)s', level=logging.INFO)
    import sys
    in_path = sys.argv[1]
    out_path = sys.argv[2]
    extract_file_list = get_file_list(in_path)
    process(extract_file_list, out_path)

