string = '□會計事務 -人工記帳  □網頁設計'
print(string[int(len(string))-int(string[::-1].index('□'))-2])
