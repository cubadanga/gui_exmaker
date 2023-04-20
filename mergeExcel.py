import os
import pandas as pd
import time
from datetime import date
import sys

print('\033[96m' + '파일 합치기를 시작합니다.'+ '\033[0m' + '\n')

tday = date.today()
current_dir = os.getcwd()
save_folder = './excel/'
data_file_folder = current_dir + save_folder

tday_s = time.strftime('%Y%m%d', time.localtime(time.time()))
def publish_excel(jointype):
    df = []
    try:
        for file in os.listdir(data_file_folder):
            if file.endswith('.xlsx') and jointype in file:
                print('Loading file...{0}'.format(file))
                df.append(pd.read_excel(os.path.join(data_file_folder,file),dtype = {'판매자 상품코드':'str','원산지 코드':'str','제품코드':'str'} ))
        df_master = pd.concat(df, axis=0, ignore_index=True)
        
        #shutil.move(src, dst)
        fileName = jointype+'_merge_naver_' + tday_s + '.xlsx'
        df_master.to_excel('./excel/' + fileName, index=False)
        print(f'\n* 합치기 완료!: [{fileName}]\n')
        return df
    
    except FileNotFoundError as e:
        print('\n' + '\033[31m \033[43m' + "오류 -"+jointype+" 파일을 찾을 수 없습니다."+'\033[0m')
        print(e)
        print('\033[31m' + "엔터를 누르면 종료합니다." + '\033[0m')
        aInput = input("")
        sys.exit()
    
    except PermissionError as e2:
        print('\n' + '\033[31m \033[43m' + "오류 - 병합 될 파일이 엑셀에서 열려있는 것 같습니다."+'\033[0m')
        print(e2)
        print('\033[31m' + "엔터를 누르면 종료합니다." + '\033[0m')
        aInput = input("")
        sys.exit()
    
    except ValueError as e3:
        print('\n' + '\033[31m \033[43m' + "오류 - excel 폴더에 병합 될 파일이 없는 것 같습니다."+'\033[0m')
        print(e3)
        print('\033[31m' + "엔터를 누르면 종료합니다." + '\033[0m')
        aInput = input("")
        sys.exit()


publish_excel('배포용')
publish_excel('개인용')


print('\n'+'\033[96m' + '엔터키를 누르면 종료합니다.'+ '\033[0m')
aaa = input()