from distutils.core import setup
import py2exe

setup(
    options= {'py2exe': {'bundle_files': 1, 'compressed': True}},
    windows=[{'script': "expo-invoice-generator.py"}],
    zipfile=None,
    console=['expo-invoice-generator.py']
)
