from setuptools import setup

with open('README.md', 'r') as fh:
    long_description = fh.read()

setup(
    name='py_canoe',
    version='0.0.1',
    description='Python Library for controlling Vector CANoe tool',
    long_description=long_description,
    long_description_content_type='text/markdown',
    py_modules=['py_canoe'],
    package_dir={'': 'src'},
    classifiers=[
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: 3.9',
        'License :: OSI Approved :: MIT License',
        'Operating System :: Microsoft :: Windows',
    ],
    url='https://github.com/chaitu-ycr/py_canoe',
    author='chaitu-ycr',
    author_email='chaitu.ycr@gmail.com',
)
