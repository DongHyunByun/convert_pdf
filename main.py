import argparse
import logging
import traceback
from datetime import datetime
import os

from to_pdf import ConvertPdf

if __name__ == '__main__':
    args = argparse.ArgumentParser()
    args.add_argument("--from_path", type=str, default="C:/convert_to_pdf_3/from_folder",
                      help="변환할 파일들을 읽을 절대경로")
    args.add_argument("--to_path", type=str, default="C:/convert_to_pdf_3/PDF",
                      help="변환된 파일들을 쓸 절대경로")
    args.add_argument("--log_path", type=str, default="C:/convert_to_pdf_3/log_folder",
                      help="에러로그 파일을 쓸 경로, 날짜(YYYYMMDD)형식으로 로그파일이 저장된다")
    config = args.parse_args()

    error_log_path = config.log_path+"/"+datetime.today().strftime("%Y%m%d")+".csv"
    # logging.basicConfig(filename=log, level=logging.ERROR)

    converter = ConvertPdf(config.from_path, config.to_path, logging)

    converter.to_csv_error_file(error_log_path)
