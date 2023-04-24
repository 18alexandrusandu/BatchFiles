cd C /:Users//Desktop
for /F "delims="  %%a in ('dir /s /b *pass*') do  ((echo  | (type "%%a") & echo:) >> find_passwords.txt)