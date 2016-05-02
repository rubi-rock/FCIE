IF "%1"=="" GOTO Help

python fcie.py %* 
GOTO Continue

:Help
python fcie.py --help
GOTO Continue

:Continue


