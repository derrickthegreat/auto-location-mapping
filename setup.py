from setuptools import setup
setup(
    name='auto-location-mapping',
    version='1.1.0',
    author='Derrick Alvarez',
    author_email='derrickcanbereached@gmail.com',
    install_requires=['pandas','pyautogui'],
    entry_points='''
        [console_scripts]
        automap=main:automap
    ''',
 )
