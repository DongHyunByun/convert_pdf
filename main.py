import argparse

from datetime import datetime,timedelta
import time

from to_pdf import ConvertPdf
from sftp_connect import Sftp

if __name__ == '__main__':
    args = argparse.ArgumentParser()
    args.add_argument("--env", type=str, default="test",
                      help="환"
                           "경, 운영의 경우 prod")
    args.add_argument("--d", type=str, default=datetime.today().strftime("%Y%m%d"),
                      help="수행할 날짜 YYYYMMDD 형식")
    config = args.parse_args()

    local_from_folder = "C:/convert_to_pdf_3/from_folder"
    local_to_folder = "C:/convert_to_pdf_3/PDF"
    error_log_path = "C:/convert_to_pdf_3/log_folder/" + config.d + ".csv"

    start = time.time()

    # [sftp]
    sftp = Sftp(config.env, config.d, local_from_folder, local_to_folder)
    sftp.get_file_from_sftp() # remote -> local 파일복사

    # [변환]`
    converter = ConvertPdf(local_from_folder, local_to_folder, config.d)
    converter.to_csv_error_file(error_log_path)

    sftp.put_file_to_sftp()  # local -> remote 파일복사

    print(str(timedelta(seconds = time.time()-start)).split(".")[0])