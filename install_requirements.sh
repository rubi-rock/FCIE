#!/usr/bin/env bash
curl -O https://www.python.org/ftp/python/3.5.1/python-3.5.1-macosx10.6.pkg
sudo installer -pkg /python-3.5.1-macosx10.6.pkg /Applications
pip3 install click
pip3 install dotmap
pip3 install xlsxwriter

