from setuptools import setup

setup(
    name='DICT',
    version='1.0',
    description="DICT.py sert à envoyer automatiquement les requêtes au concessionaires à partir du fichier .zip. fourni lors d'une demande",
    author='erlouche',
    author_email='erle.marec@gmail.com',
    packages=['your_package_name'],
    install_requires=[
        'pdfplumber',
        'pandas',
        'tkinter',
    ],
)
