from setuptools import setup

setup(
    name='py_canoe',
    version='0.1',
    packages=['src',],
    url='https://github.com/chaitu-ycr/py_canoe',
    license='MIT',
    author='chaitu-ycr',
    author_email='chaitu.ycr@gmail.com',
    description='Python Library for controlling Vector CANoe tool',
    install_requires=[
        'pywin32'
    ]
)