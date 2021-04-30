import os

def check_file(file_path,typeall=None):
    os.chdir(file_path)
    print(os.path.abspath(os.curdir))
    all_file = os.listdir()
    files = []
    for f in all_file:
        if os.path.isdir(f):
            files.extend(check_file(file_path+'\\'+f,typeall))
            os.chdir(file_path)
        else:
            _, type1 = os.path.splitext(f)
            if not typeall or type1 in typeall:
                files.append(f)
    return files
