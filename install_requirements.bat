cd install
wget --no-check-certificate https://www.python.org/ftp/python/3.5.1/python-3.5.1-amd64.exe
python-3.5.1-amd64.exe
del python-3.5.1-amd64.exe
call install_python_libs.bat
cd ..
python fcie.py
