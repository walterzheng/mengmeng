import logging
import os
import re

def rename(path):
    matcher = re.compile(r'P(?P<num>\d+)\.(?P<filename>.+?)\.mp4')
    if not os.path.isdir(path):
        logging.error(f'{path} is not a directory!')
        return None

    mv_info = []
    page_num_len = 0
    for filename in os.listdir(path):
        matcher_result = matcher.match(filename)
        file = os.path.join(path, filename)
        if matcher_result and os.path.isfile(file):
            page_num = matcher_result.group('num')
            real_name = matcher_result.group('filename')
            mv_info.append((filename, page_num, real_name))
            page_num_len = max(page_num_len, len(str(page_num)))
    
    for filename, page_num, real_name in mv_info:
        align_page_num = '0' * (page_num_len - len(str(page_num))) + str(page_num)
        new_file_name = f"P{align_page_num}.{real_name}.mp4"
        logging.info(f"Rename '{filename}' to {new_file_name}")

        file = os.path.join(path, filename)
        new_file = os.path.join(path, new_file_name)
        os.rename(file, new_file)


if __name__ == '__main__':
    logging.basicConfig(format='[%(asctime)s] %(lineno)s %(levelname)s: %(message)s', level=logging.INFO)
    import sys
    in_path = sys.argv[1]
    rename(in_path)