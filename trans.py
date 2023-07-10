import os

# 디렉토리 경로 입력 받기
directory = input("디렉토리 경로를 입력하세요: ")

# 디렉토리 내의 파일 목록 얻기
file_list = os.listdir(directory)

# 파일 확장자 변경
for file_name in file_list:
    if file_name.endswith('.txt'):
        # 파일 이름 변경
        new_file_name = file_name[:-4] + '.py'
        old_file_path = os.path.join(directory, file_name)
        new_file_path = os.path.join(directory, new_file_name)
        os.rename(old_file_path, new_file_path)

print("확장자 변경이 완료되었습니다.")
