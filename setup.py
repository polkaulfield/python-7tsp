from distutils.core import setup # Need this to handle modules
import py2exe 
import os, sys, shutil, subprocess, ctypes, psutil
sys.argv.append('py2exe')

setup(
    options = {'py2exe': {'bundle_files': 1, 'compressed': True}},
    windows = [{'script': "main.py"}],
    zipfile = None,
)