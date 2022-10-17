# PDF변환 프로그램

1. **기간 : 2021.09.15 ~ 2022.10.17** 

2. **기술스택**
    1. 개발언어 : python
    2. IDE : pycharm
 
3. **내용**
    1. SFTP서버를 통해 파일들을 다운로드한다. (from_folder에 저장)
    2. 파일들을 PDF로 변환한다.변환할 수 있는 형식은 다음과 같다. (to_folder에 저장)
    ![image](https://user-images.githubusercontent.com/50386280/196077412-5016985a-47cd-4bbd-a521-0d7b99cd0c3e.png)
    3. 변환한 파일을 sftp서버로 보낸다.  

4. **파일설명**
    1. main.py : main 파일  
    2. sftp_connect.py : sfpt연결
    3. to_pdf.py : pdf변환
    4. 보안모듈(Automation) 한글 보안 경고창 제거용 보안모듈  
